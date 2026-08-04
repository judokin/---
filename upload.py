# 使用提醒:
# 1. xbot包提供软件自动化、数据表格、Excel、日志、AI等功能
# 2. package包提供访问当前应用数据的功能，如获取元素、访问全局变量等
# 3. 当此模块作为流程独立运行时执行 main 函数

from __future__ import annotations

import importlib.util
import json
import os
import re
from pathlib import Path

import requests

try:
    import xbot
    from xbot import print, sleep
    from . import package
    from .package import variables as glv
except Exception:
    glv = {
        'gvar_对应运营id': '',
        'gvar_err_msg': '',
        'gvar_shop_name': '',
    }


PROJECT_DIR = Path(r'D:\项目文件\AI自动上架')
OUTPUT_PATH_LIST_FILE = Path(r'D:\NAS_download\nas_path_out_path.txt')
UPLOAD_HISTORY_FILE = Path(r'D:\NAS_download\uploaded_files.json')
FEISHU_PARENT_NODE = 'RE9OfQ3PFlwErcdrJXUceUKZn5c'
GENERATED_FILE_RE = re.compile(r'.+_\d{14}\.xlsm$', re.IGNORECASE)

config_file_path = r'D:\rpa_tools\feishu\config.py'
spec = importlib.util.spec_from_file_location('config', config_file_path)
config = importlib.util.module_from_spec(spec)
spec.loader.exec_module(config)


def get_tenant_access_token() -> str:
    """取得飞书 tenant_access_token；日志中不输出访问令牌。"""
    response = requests.post(
        'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/',
        json={
            'app_id': config.app_id,
            'app_secret': config.app_secret,
        },
        timeout=30,
    )
    response.raise_for_status()
    result = response.json()
    token = result.get('tenant_access_token', '')
    if not token:
        raise RuntimeError(
            f"获取飞书访问令牌失败：code={result.get('code')}，"
            f"msg={result.get('msg')}"
        )
    return token


def upload_file(
    file_path: Path | str,
    parent_node: str = FEISHU_PARENT_NODE,
    tenant_access_token: str | None = None,
) -> dict:
    """上传单个文件，成功返回飞书响应；失败直接抛出异常。"""
    file_path = Path(file_path).resolve()
    if not file_path.is_file():
        raise FileNotFoundError(f"待上传文件不存在：{file_path}")

    token = tenant_access_token or get_tenant_access_token()
    file_size = file_path.stat().st_size
    headers = {'Authorization': f'Bearer {token}'}
    with file_path.open('rb') as file_stream:
        files = {
            'file_name': (None, file_path.name),
            'parent_type': (None, 'explorer'),
            'parent_node': (None, parent_node),
            'size': (None, str(file_size)),
            'file': (file_path.name, file_stream),
        }
        response = requests.post(
            'https://open.feishu.cn/open-apis/drive/v1/files/upload_all',
            headers=headers,
            files=files,
            timeout=180,
        )

    try:
        result = response.json()
    except ValueError as exc:
        raise RuntimeError(
            f"飞书上传返回非 JSON：HTTP {response.status_code}"
        ) from exc
    if not response.ok or result.get('code') not in (None, 0):
        raise RuntimeError(
            f"飞书上传失败：HTTP {response.status_code}，"
            f"code={result.get('code')}，msg={result.get('msg')}"
        )
    if str(result.get('msg', '')).lower() not in ('success', ''):
        raise RuntimeError(f"飞书上传失败：{result.get('msg')}")
    return result


def send_message(
    message: str,
    webhook_url: str = (
        'https://open.feishu.cn/open-apis/bot/v2/hook/'
        '53eda273-cc3d-4092-bfb8-01d6d5122aa5'
    ),
) -> None:
    """发送上传结果通知；通知失败不改变文件的上传成功状态。"""
    payload = {
        'msg_type': 'text',
        'content': {'text': message},
    }
    try:
        response = requests.post(
            webhook_url,
            headers={'Content-Type': 'application/json'},
            data=json.dumps(payload, ensure_ascii=False).encode('utf-8'),
            timeout=30,
        )
        if response.ok:
            print('消息发送成功')
        else:
            print(f'消息发送失败：HTTP {response.status_code}')
    except requests.RequestException as exc:
        print(f'消息发送失败：{exc}')


def normalize_path(file_path: Path | str) -> str:
    """生成不区分 Windows 路径大小写的上传记录键。"""
    return os.path.normcase(os.path.abspath(str(file_path)))


def file_fingerprint(file_path: Path) -> dict[str, int]:
    """同路径文件内容被替换后允许重新上传。"""
    stat = file_path.stat()
    return {
        'size': stat.st_size,
        'mtime_ns': stat.st_mtime_ns,
    }


def load_upload_history() -> dict[str, dict]:
    if not UPLOAD_HISTORY_FILE.is_file():
        return {}
    try:
        data = json.loads(UPLOAD_HISTORY_FILE.read_text(encoding='utf-8'))
    except (OSError, ValueError) as exc:
        raise RuntimeError(f"上传记录文件损坏：{UPLOAD_HISTORY_FILE}") from exc
    if isinstance(data, list):
        return {normalize_path(path): {} for path in data}
    if not isinstance(data, dict):
        raise RuntimeError(f"上传记录格式错误：{UPLOAD_HISTORY_FILE}")
    return data


def save_upload_history(history: dict[str, dict]) -> None:
    """原子保存，避免流程中断把上传记录写坏。"""
    UPLOAD_HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    temp_path = UPLOAD_HISTORY_FILE.with_suffix('.json.tmp')
    temp_path.write_text(
        json.dumps(history, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )
    os.replace(temp_path, UPLOAD_HISTORY_FILE)


def is_already_uploaded(file_path: Path, history: dict[str, dict]) -> bool:
    record = history.get(normalize_path(file_path))
    if record is None:
        return False
    # 兼容旧的“只记录路径”格式；新格式同时校验文件指纹。
    if not record:
        return True
    fingerprint = file_fingerprint(file_path)
    return (
        record.get('size') == fingerprint['size']
        and record.get('mtime_ns') == fingerprint['mtime_ns']
    )


def read_output_path_list() -> list[Path]:
    """兼容 v2.py 写入的逗号分隔和换行分隔路径。"""
    if not OUTPUT_PATH_LIST_FILE.is_file():
        return []
    content = OUTPUT_PATH_LIST_FILE.read_text(encoding='utf-8-sig').strip()
    if not content:
        return []
    return [
        Path(value.strip())
        for value in re.split(r'[,\r\n]+', content)
        if value.strip()
    ]


def discover_generated_files() -> list[Path]:
    """发现 v2.py 已生成的时间戳结果，并合并输出路径清单。"""
    candidates = read_output_path_list()
    if PROJECT_DIR.is_dir():
        candidates.extend(
            path
            for path in PROJECT_DIR.glob('*.xlsm')
            if GENERATED_FILE_RE.fullmatch(path.name)
        )

    result = []
    seen = set()
    for file_path in candidates:
        resolved_path = file_path.resolve()
        path_key = normalize_path(resolved_path)
        if path_key in seen:
            continue
        seen.add(path_key)
        if resolved_path.is_file():
            result.append(resolved_path)
        else:
            print(f'清单中的文件不存在，跳过：{resolved_path}')
    return sorted(result, key=lambda path: (path.stat().st_mtime_ns, path.name))


def build_success_message(file_path: Path) -> str:
    shop_name = str(glv.get('gvar_shop_name', '') or '').strip()
    operator_id = str(glv.get('gvar_对应运营id', '') or '').strip()
    error_message = str(glv.get('gvar_err_msg', '') or '').strip()
    prefix = f'【{shop_name}】' if shop_name else ''
    mention = f'<at user_id="{operator_id}"></at>' if operator_id else ''
    return f'{prefix}{file_path} 已经上传到共享盘{mention}{error_message}'


def main(args=None):
    generated_files = discover_generated_files()
    if not generated_files:
        print('没有发现待上传的已生成 .xlsm 文件')
        return []

    history = load_upload_history()
    pending_files = [
        file_path
        for file_path in generated_files
        if not is_already_uploaded(file_path, history)
    ]
    skipped_count = len(generated_files) - len(pending_files)
    for file_path in generated_files:
        if is_already_uploaded(file_path, history):
            print(f'已上传，跳过：{file_path.name}')

    if not pending_files:
        print(f'全部文件均已上传，本次跳过 {skipped_count} 个')
        return []

    token = get_tenant_access_token()
    uploaded_files = []
    failed_files = []
    for file_index, file_path in enumerate(pending_files, start=1):
        print(
            f'\n========== 上传 [{file_index}/{len(pending_files)}]：'
            f'{file_path.name} =========='
        )
        try:
            result = upload_file(
                file_path,
                parent_node=FEISHU_PARENT_NODE,
                tenant_access_token=token,
            )
        except Exception as exc:
            failed_files.append((file_path, str(exc)))
            print(f'上传失败，保留待下次重试：{file_path.name}，原因：{exc}')
            send_message(f'{file_path} 上传失败：{exc}')
            continue

        fingerprint = file_fingerprint(file_path)
        history[normalize_path(file_path)] = {
            **fingerprint,
            'file_name': file_path.name,
            'file_token': result.get('data', {}).get('file_token', ''),
        }
        save_upload_history(history)
        uploaded_files.append(file_path)
        print(f'上传成功：{file_path.name}')
        send_message(build_success_message(file_path))

    print(
        f'\n批量上传完成：成功 {len(uploaded_files)} 个，'
        f'已上传跳过 {skipped_count} 个，失败 {len(failed_files)} 个'
    )
    if failed_files:
        failure_message = '；'.join(
            f'{file_path.name}: {reason}'
            for file_path, reason in failed_files
        )
        raise RuntimeError(f'以下文件上传失败：{failure_message}')
    return uploaded_files


if __name__ == '__main__':
    main()
