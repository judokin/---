"""按材质方向读取飞书多维表格视图，并生成 SKCS.txt。"""

from __future__ import annotations

import argparse
import time
from pathlib import Path
from typing import Any, Iterable

import requests

import config


BASE_TOKEN = "Y264bSCfHayWKis39MDcL5eQnab"
TABLE_ID = "tbluo8xCboMg3b0T"
OUTPUT_PATH = Path(r"D:\NAS_download\SKCS.txt")
MISSING_IMAGE_SKCS_PATH = OUTPUT_PATH.with_name("数据库无图片的SKC.txt")
MISSING_IMAGE_DETAILS_PATH = Path(__file__).resolve().parent / "缺图明细.txt"
WEBHOOK_URL = (
    "https://open.feishu.cn/open-apis/bot/v2/hook/"
    "53eda273-cc3d-4092-bfb8-01d6d5122aa5"
)

MATERIAL_VIEWS = {
    "印花地毯": "vew1r9vIKo",
    "仿羊绒厨房垫": "vewyE7tvJK",
    "三明治户外垫": "vew0xHeRAu",
}

REQUEST_TIMEOUT = 30
MAX_RETRIES = 5
PAGE_SIZE = 500
MAX_MESSAGE_LENGTH = 3500


def request_json(
    session: requests.Session,
    method: str,
    url: str,
    *,
    operation: str,
    **kwargs: Any,
) -> dict[str, Any]:
    """请求飞书接口，遇到网络错误或服务端错误时自动重试。"""
    last_error = "未知错误"
    for attempt in range(MAX_RETRIES):
        try:
            response = session.request(
                method,
                url,
                timeout=REQUEST_TIMEOUT,
                **kwargs,
            )
            if response.status_code >= 500:
                last_error = f"HTTP {response.status_code}: {response.text[:500]}"
            else:
                response.raise_for_status()
                result = response.json()
                if result.get("code", 0) != 0:
                    raise RuntimeError(f"{operation}失败：{result}")
                return result
        except (requests.RequestException, ValueError) as exc:
            last_error = f"{type(exc).__name__}: {exc}"

        if attempt < MAX_RETRIES - 1:
            wait_seconds = 2**attempt
            print(
                f"{operation}第 {attempt + 1} 次失败，"
                f"{wait_seconds} 秒后重试：{last_error}"
            )
            time.sleep(wait_seconds)

    raise RuntimeError(f"{operation}连续 {MAX_RETRIES} 次失败：{last_error}")


def get_tenant_access_token(session: requests.Session) -> str:
    result = request_json(
        session,
        "POST",
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/",
        operation="获取飞书 tenant_access_token",
        json={"app_id": config.app_id, "app_secret": config.app_secret},
    )
    token = result.get("tenant_access_token")
    if not token:
        raise RuntimeError("飞书鉴权响应中缺少 tenant_access_token")
    return str(token)


def extract_cell_texts(value: Any) -> list[str]:
    """兼容飞书文本字段可能返回的字符串、富文本列表或对象。"""
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        result: list[str] = []
        for item in value:
            result.extend(extract_cell_texts(item))
        return result
    if isinstance(value, dict):
        for key in ("text", "name", "value"):
            if key in value:
                return extract_cell_texts(value[key])
    return [str(value)]


def normalize_skcs(values: Iterable[Any]) -> list[str]:
    """清理空行并按出现顺序去重。"""
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        for text in extract_cell_texts(value):
            for skc in text.splitlines():
                skc = skc.strip()
                if skc and skc not in seen:
                    seen.add(skc)
                    result.append(skc)
    return result


def fetch_skcs(
    session: requests.Session,
    token: str,
    view_id: str,
) -> list[str]:
    """分页读取指定视图中的 SKC 字段。"""
    url = (
        f"https://open.feishu.cn/open-apis/bitable/v1/apps/{BASE_TOKEN}"
        f"/tables/{TABLE_ID}/records/search?page_size={PAGE_SIZE}"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    page_token = ""
    raw_values: list[Any] = []

    while True:
        request_url = url
        if page_token:
            request_url += f"&page_token={page_token}"
        result = request_json(
            session,
            "POST",
            request_url,
            operation=f"读取飞书视图 {view_id}",
            headers=headers,
            json={"view_id": view_id, "field_names": ["SKC"]},
        )
        data = result.get("data") or {}
        for record in data.get("items") or []:
            raw_values.append((record.get("fields") or {}).get("SKC"))

        if not data.get("has_more"):
            break
        page_token = str(data.get("page_token") or "")
        if not page_token:
            raise RuntimeError(f"飞书视图 {view_id} 返回 has_more=true 但缺少 page_token")

    return normalize_skcs(raw_values)


def write_skcs(skcs: list[str], output_path: Path = OUTPUT_PATH) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    content = "\n".join(skcs)
    if content:
        content += "\n"
    output_path.write_text(content, encoding="utf-8")


def build_notification_parts(material_direction: str, skcs: list[str]) -> list[str]:
    header = (
        "【SKCS.txt 更新通知】\n"
        f"材质方向：{material_direction}\n"
        f"SKC 数量：{len(skcs)}\n"
        f"保存位置：{OUTPUT_PATH}\n"
    )
    if not skcs:
        return [header + "最终 SKCS.txt 没有可写入的 SKC。"]

    lines = [f"{index}. {skc}" for index, skc in enumerate(skcs, start=1)]
    parts: list[str] = []
    current = header + "SKC 明细：\n"
    for line in lines:
        candidate = current + line + "\n"
        if len(candidate) > MAX_MESSAGE_LENGTH and current.strip():
            parts.append(current.rstrip())
            current = f"【SKC 明细续】材质方向：{material_direction}\n{line}\n"
        else:
            current = candidate
    if current.strip():
        parts.append(current.rstrip())
    return parts


def send_notifications(
    session: requests.Session,
    material_direction: str,
    skcs: list[str],
) -> None:
    for index, content in enumerate(
        build_notification_parts(material_direction, skcs), start=1
    ):
        request_json(
            session,
            "POST",
            WEBHOOK_URL,
            operation=f"发送飞书群通知第 {index} 段",
            json={"msg_type": "text", "content": {"text": content}},
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="按材质方向读取飞书视图并生成 D:\\NAS_download\\SKCS.txt"
    )
    parser.add_argument(
        "--material-direction",
        required=True,
        choices=tuple(MATERIAL_VIEWS),
        help="要读取的材质方向",
    )
    parser.add_argument(
        "--check-images",
        action="store_true",
        help="生成 SKCS.txt 后复用 pg_set_excel_by_skc_size.py 检查图片完整性",
    )
    parser.add_argument(
        "--no-group-notification",
        action="store_true",
        help="不发送图片检查结果；SKCS.txt 更新通知仍会发送",
    )
    return parser.parse_args()


def check_images(
    material_direction: str,
    skcs: list[str],
    notify: bool,
) -> dict[str, object]:
    """复用 PostgreSQL 图片匹配规则，检查当前 SKC 清单的图片完整性。"""
    from pg_set_excel_by_skc_size import (
        DEFAULT_PG_CONFIG_PATH,
        check_skc_image_completeness,
        save_image_check_reports,
        send_feishu_text,
        send_missing_image_notifications,
    )

    result = check_skc_image_completeness(
        skcs,
        material_direction,
        DEFAULT_PG_CONFIG_PATH,
    )
    detail_lines = save_image_check_reports(
        result,
        MISSING_IMAGE_SKCS_PATH,
        MISSING_IMAGE_DETAILS_PATH,
    )
    checked_count = len(result["checked_skcs"])
    incomplete_count = len({skc for skc, _ in result["missing_image_details"]})
    complete_count = checked_count - incomplete_count
    no_database_count = len(result["missing_image_skcs"])
    summary = (
        "【SKC 图片完整性检查】\n"
        f"材质方向：{material_direction}\n"
        f"检查 SKC：{checked_count} 个\n"
        f"图片齐全：{complete_count} 个\n"
        f"图片不完整：{incomplete_count} 个\n"
        f"数据库无有效图片：{no_database_count} 个\n"
        f"检查尺寸：{', '.join(result['checked_sizes'])}\n"
        f"缺图报告：{MISSING_IMAGE_DETAILS_PATH}"
    )
    print(summary.replace("【SKC 图片完整性检查】\n", ""))
    if notify:
        send_feishu_text(WEBHOOK_URL, summary)
        if detail_lines:
            send_missing_image_notifications(detail_lines)
        print("图片检查结果已发送到飞书群")
    else:
        print("已按参数关闭图片检查群通知")
    return result


def main() -> None:
    args = parse_args()
    view_id = MATERIAL_VIEWS[args.material_direction]
    with requests.Session() as session:
        token = get_tenant_access_token(session)
        skcs = fetch_skcs(session, token, view_id)
    fetched_count = len(skcs)
    if args.check_images:
        check_result = check_images(
            args.material_direction,
            skcs,
            notify=not args.no_group_notification,
        )
        incomplete_skcs = {
            skc for skc, _required_image in check_result["missing_image_details"]
        }
        skcs = [skc for skc in skcs if skc not in incomplete_skcs]
        print(
            f"图片检查筛选：候选 {fetched_count} 个，"
            f"排除缺图 {len(incomplete_skcs)} 个，保留 {len(skcs)} 个"
        )

    with requests.Session() as session:
        write_skcs(skcs)
        print(f"已从视图 {view_id} 读取 {fetched_count} 个候选 SKC")
        print(f"已保存到：{OUTPUT_PATH}")
        send_notifications(session, args.material_direction, skcs)
        print("飞书群通知发送成功")


if __name__ == "__main__":
    main()
