r"""从 PostgreSQL 读取图片，按 SKC、材质方向和尺寸匹配上架图片 URL。

默认测试文件：
    D:\项目文件\AI自动上架\
    output_feishu_table_data_亚马逊--冬豚--北美（子账号）_AR1002-NEW.xlsx

图片数据来源：
    PostgreSQL public.ods_fbm_image_upload
    数据库连接配置默认读取脚本同目录的 pg.config。

默认另存为“原文件名_尺寸图片已匹配.xlsx”；传入 --in-place 可覆盖源文件。
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import unquote, urlsplit
from urllib.request import Request, urlopen

import pandas as pd
import psycopg2


DEFAULT_INPUT_PATH = Path(
    r"D:\项目文件\AI自动上架\output_feishu_table_data_亚马逊-北蓉-北美（子账号）_Minimalist-AR3-POD.xlsx"
)
DEFAULT_INPUT_DIR = DEFAULT_INPUT_PATH.parent
DEFAULT_INPUT_PATTERN = "output_feishu_table_data_*.xlsx"
MATCHED_FILE_SUFFIX = "_尺寸图片已匹配"
DEFAULT_PG_CONFIG_PATH = Path(__file__).resolve().parent / "pg.config"
DEFAULT_MISSING_IMAGE_SKCS_PATH = DEFAULT_INPUT_DIR / "缺少图片的SKC.txt"
DEFAULT_MISSING_IMAGE_DETAILS_PATH = Path(__file__).resolve().parent / "缺图明细.txt"
FEISHU_MISSING_IMAGE_WEBHOOK = (
    "https://open.feishu.cn/open-apis/bot/v2/hook/"
    "53eda273-cc3d-4092-bfb8-01d6d5122aa5"
)
RUG_MATERIAL = "印花地毯"
CHRISTMAS_RUG_MATERIAL = "印花地毯圣诞专用"
CHRISTMAS_IMAGE_VERSION = "2026圣诞专题v1"
KITCHEN_MAT_MATERIAL = "仿羊绒厨房垫"
OUTDOOR_MAT_MATERIAL = "三明治户外垫"
SUPPORTED_MATERIALS = {
    RUG_MATERIAL,
    KITCHEN_MAT_MATERIAL,
    OUTDOOR_MAT_MATERIAL,
}

# 尺寸及图片规则统一读取 public.listing_image_versions，不在此处预设。

IMAGE_COLUMNS = [
    "主图片 URL",
    "其他图片 URL1",
    "其他图片 URL2",
    "其他图片 URL3",
    "其他图片 URL4",
    "其他图片 URL5",
    "其他图片 URL6",
    "其他图片 URL7",
    "其他图片 URL8",
    "样本图片 URL",
]


def normalize_text(value: object) -> str:
    """把 Excel 值转换为去除首尾空格的字符串。"""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def normalize_size(value: object) -> str:
    """统一尺寸写法，兼容 2x3、2 X 3、2×3、带 FT/IN 单位的值。"""
    size = normalize_text(value).upper().replace("×", "X").replace(" ", "")
    for suffix in ("FEET", "FOOT", "FT", "INCHES", "INCH", "IN"):
        if size.endswith(suffix):
            size = size[: -len(suffix)]
            break
    return size


def normalize_image_name(value: object) -> str:
    """统一图片文件名，供关键词、尺寸和错误裸文件名识别共用。"""
    return (
        normalize_text(value)
        .upper()
        .replace("×", "X")
        .replace("＋", "+")
        .replace(" ", "")
    )


def size_from_image_name(file_name: object) -> str:
    """从图片文件名中提取第一个 AxB 尺寸，例如白底图5X7.jpg。"""
    normalized_name = normalize_image_name(file_name)
    match = re.search(
        r"(?<!\d)(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)(?!\d)",
        normalized_name,
    )
    if not match:
        return ""
    return normalize_size(f"{match.group(1)}X{match.group(2)}")


def is_bare_skc_size_image(
    file_name: object,
    skc: object,
    size: object,
) -> bool:
    """识别误上传的“<SKC>-<尺寸>.<扩展名>”裸命名图片。"""
    normalized_skc = normalize_image_name(skc)
    normalized_size = normalize_size(size)
    if not normalized_skc or not normalized_size:
        return False

    file_stem = Path(normalize_text(file_name)).stem
    normalized_stem = normalize_image_name(file_stem)
    bare_suffix = normalize_image_name(f"{normalized_skc}-{normalized_size}")
    # URL 对象名可能在真实文件名前附加 SKC/目录前缀，因此使用结尾匹配。
    return normalized_stem.endswith(bare_suffix)


def first_url(rows: pd.DataFrame) -> str:
    """取得数据集合中的第一个非空图片地址。"""
    if rows.empty or "图片地址" not in rows.columns:
        return ""
    for value in rows["图片地址"].tolist():
        url = normalize_text(value)
        if url:
            return url
    return ""


def urls(rows: pd.DataFrame) -> list[str]:
    """按源文件顺序取得非空图片地址。"""
    if rows.empty or "图片地址" not in rows.columns:
        return []
    return [
        url
        for url in (normalize_text(value) for value in rows["图片地址"].tolist())
        if url
    ]


def pick_named_url(image_df: pd.DataFrame, *keywords: str) -> str:
    """按文件名关键词优先级取得图片地址。"""
    file_names = image_df["文件名"].map(normalize_image_name)
    for keyword in keywords:
        normalized_keyword = normalize_image_name(keyword)
        matched = image_df[
            file_names.str.contains(normalized_keyword, regex=False)
        ]
        url = first_url(matched)
        if url:
            return url
    return ""


def pick_sample_url(image_df: pd.DataFrame) -> str:
    """优先按图片类型选择样本图，再兼容常见样本图文件名。"""
    if "图片类型" in image_df.columns:
        image_types = image_df["图片类型"].map(normalize_text)
        url = first_url(image_df[image_types == "样本图片"])
        if url:
            return url
    return pick_named_url(image_df, "SWATCH", "SWACH", "SWITCH", "SWICH")








def assign_urls(
    target_df: pd.DataFrame,
    row_index: int,
    values: dict[str, str],
) -> None:
    """只写入非空 URL，缺失图片保持空白。"""
    for column, value in values.items():
        if value:
            target_df.at[row_index, column] = value


def load_pg_config(config_path: Path) -> dict[str, object]:
    """读取 pg.config，兼容中文配置名和常见 PostgreSQL 配置名。"""
    if not config_path.is_file():
        raise FileNotFoundError(f"PostgreSQL 配置文件不存在：{config_path}")

    key_aliases = {
        "连接地址": "host",
        "地址": "host",
        "主机": "host",
        "host": "host",
        "端口": "port",
        "port": "port",
        "数据库": "dbname",
        "数据库名": "dbname",
        "database": "dbname",
        "dbname": "dbname",
        "账号": "user",
        "用户名": "user",
        "user": "user",
        "密码": "password",
        "password": "password",
    }
    config: dict[str, object] = {}
    for raw_line in config_path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith(("#", ";", "[")):
            continue
        separator = "：" if "：" in line else "=" if "=" in line else ":"
        if separator not in line:
            continue
        raw_key, raw_value = line.split(separator, 1)
        key = key_aliases.get(raw_key.strip().lower())
        if key:
            config[key] = raw_value.strip()

    required_keys = {"host", "port", "dbname", "user", "password"}
    missing_keys = sorted(required_keys - set(config))
    if missing_keys:
        raise ValueError(
            f"PostgreSQL 配置文件缺少字段：{missing_keys}，文件：{config_path}"
        )
    config["port"] = int(config["port"])
    config["connect_timeout"] = 10
    return config


def decoded_image_name(url: object, fallback_name: object) -> str:
    """优先从 URL 路径还原中文文件名，规避数据库文件名乱码。"""
    normalized_url = normalize_text(url)
    if normalized_url:
        decoded_path = unquote(urlsplit(normalized_url).path)
        decoded_name = Path(decoded_path).name
        if decoded_name:
            return decoded_name
    return normalize_text(fallback_name)


def infer_image_type(file_name: str, raw_image_type: object) -> str:
    """从解码后的文件名恢复图片类型，数据库原字段正常时作为回退。"""
    normalized_name = file_name.upper()
    if any(
        keyword in normalized_name
        for keyword in ("SWATCH", "SWACH", "SWITCH", "SWICH")
    ):
        return "样本图片"
    if "白底图" in file_name:
        return "白底图"
    if "封面图" in file_name or "8X10客厅" in normalized_name:
        return "主图片"

    normalized_type = normalize_text(raw_image_type)
    if normalized_type in {"主图片", "其他图片", "白底图", "样本图片"}:
        return normalized_type
    return "其他图片"


def load_image_data_from_pg(
    skcs: list[str],
    config_path: Path,
) -> dict[str, pd.DataFrame]:
    """一次查询所有目标 SKC，并转换成现有图片匹配规则所需的字段。"""
    if not skcs:
        return {}

    query = """
        SELECT skc, file_name, url, size_label, image_type
        FROM public.ods_fbm_image_upload
        WHERE skc = ANY(%s)
          AND status = 'success'
          AND COALESCE(url, '') <> ''
          AND (url_expire_at IS NULL OR url_expire_at > NOW())
        ORDER BY skc, id
    """
    with psycopg2.connect(**load_pg_config(config_path)) as connection:
        connection.set_session(readonly=True)
        with connection.cursor() as cursor:
            cursor.execute("SET LOCAL statement_timeout = '15s'")
            cursor.execute(query, (skcs,))
            rows = cursor.fetchall()

    records = []
    for skc, file_name, url, size_label, image_type in rows:
        restored_name = decoded_image_name(url, file_name)
        effective_type = infer_image_type(restored_name, image_type)
        effective_size = normalize_size(size_label)
        if effective_type == "白底图":
            # 白底图优先相信文件名中的尺寸，数据库 size_label 仅作为回退。
            effective_size = size_from_image_name(restored_name) or effective_size
        records.append({
            "SKC": normalize_text(skc),
            "文件名": restored_name,
            "图片地址": normalize_text(url),
            "尺寸": effective_size,
            "图片类型": effective_type,
        })

    if not records:
        return {}
    image_df = pd.DataFrame(records)
    return {
        skc: group.reset_index(drop=True)
        for skc, group in image_df.groupby("SKC", sort=False)
    }




def value_at(values: list[str], index: int) -> str:
    """安全读取列表项；图片不足时返回空字符串。"""
    return values[index] if index < len(values) else ""








def match_rug_images(
    input_path: Path,
    output_path: Path,
    pg_config_path: Path = DEFAULT_PG_CONFIG_PATH,
    image_version: str | None = None,
) -> dict[str, object]:
    """按数据库规则填图，逐行支持不同材质，缺少尺寸配置时停止。"""
    from pg_v2_set_excel_by_skc_size import load_enabled_image_version, version_image_values, rules_for_size
    target = pd.read_excel(input_path)
    required = {'SKC', '材质方向', 'size_text'}
    if not required.issubset(target.columns):
        raise ValueError(f"目标 Excel 缺少字段：{sorted(required - set(target.columns))}")
    for column in IMAGE_COLUMNS:
        if column not in target:
            target[column] = ''
        target[column] = target[column].astype(object)
    target['SKC'] = target['SKC'].map(normalize_text)
    target['材质方向'] = target['材质方向'].map(normalize_text)
    versions = {}
    eligible = []
    for index, row in target.iterrows():
        size, skc, material = normalize_size(row['size_text']), row['SKC'], row['材质方向']
        if not skc or not size:
            continue
        if material not in versions:
            versions[material] = load_enabled_image_version(image_version, pg_config_path, material)
        rules = rules_for_size(versions[material], size)
        eligible.append((index, skc, material, size, rules))
    images = load_image_data_from_pg(list(dict.fromkeys(r[1] for r in eligible)), pg_config_path)
    processed, missing, details = [], [], []
    matched_rows = 0
    for index, skc, material, size, rules in eligible:
        frame = images.get(skc)
        values, required_names = version_image_values(frame, rules, size)
        target.loc[index, IMAGE_COLUMNS] = ''
        assign_urls(target, index, values)
        details.extend(dict(skc=skc, material=material, size=size, target_column=slot, required_image=name)
                       for slot, name in required_names.items() if not values[slot])
        if frame is None or frame.empty:
            missing.append(skc)
        else:
            processed.append(skc)
            matched_rows += 1
    output_path.parent.mkdir(parents=True, exist_ok=True)
    target.to_excel(output_path, index=False)
    return dict(output_path=output_path, processed_skcs=list(dict.fromkeys(processed)),
                missing_image_skcs=list(dict.fromkeys(missing)), missing_image_details=details,
                skipped_skcs={}, matched_rows=matched_rows)


def check_skc_image_completeness(
    skcs: Iterable[str],
    material: str,
    pg_config_path: Path = DEFAULT_PG_CONFIG_PATH,
    sizes: Iterable[str] | None = None,
    image_version: str | None = None,
) -> dict[str, object]:
    """尺寸、必需图片和排列全部取数据库唯一启用版本。"""
    from pg_v2_set_excel_by_skc_size import load_enabled_image_version, version_image_values, rules_for_size
    version = load_enabled_image_version(image_version, pg_config_path, normalize_text(material))
    normalized_skcs = list(dict.fromkeys(normalize_text(v) for v in skcs if normalize_text(v)))
    concrete_sizes = [s for s in version['rules_by_size'] if s not in {'其他', '全部尺寸'}]
    check_sizes = list(dict.fromkeys(normalize_size(s) for s in (sizes or concrete_sizes or list(version['rules_by_size']))))
    for size in check_sizes:
        rules_for_size(version, size)
    images = load_image_data_from_pg(normalized_skcs, pg_config_path)
    missing_skcs, details = [], []
    for skc in normalized_skcs:
        frame = images.get(skc)
        if frame is None or frame.empty:
            missing_skcs.append(skc)
        for size in check_sizes:
            values, required = version_image_values(frame, rules_for_size(version, size), size)
            details.extend((skc, name) for slot, name in required.items() if not values[slot])
    return dict(checked_skcs=normalized_skcs, checked_sizes=check_sizes,
                missing_image_skcs=missing_skcs, missing_image_details=list(dict.fromkeys(details)),
                image_version=version['version_name'])


def check_christmas_skc_images(skcs, pg_config_path, sizes=None):
    return check_skc_image_completeness(skcs, CHRISTMAS_RUG_MATERIAL, pg_config_path, sizes)


def save_image_check_reports(
    result: dict[str, object],
    missing_skc_path: Path,
    detail_path: Path,
) -> list[str]:
    """保存完整性检查报告，并返回适合群通知的缺图明细文本。"""
    missing_skcs = list(result["missing_image_skcs"])
    missing_skc_path.parent.mkdir(parents=True, exist_ok=True)
    missing_skc_path.write_text(
        "".join(f"{skc}\n" for skc in missing_skcs),
        encoding="utf-8-sig",
    )

    details_by_skc: dict[str, list[str]] = {}
    for skc, required_image in result["missing_image_details"]:
        details_by_skc.setdefault(skc, []).append(required_image)

    detail_lines: list[str] = []
    for skc, required_images in details_by_skc.items():
        detail_lines.append(f"SKC: {skc}")
        for required_image in required_images:
            detail_lines.append(f"  - 缺少图片: {required_image}")
        detail_lines.append("")

    detail_path.parent.mkdir(parents=True, exist_ok=True)
    detail_path.write_text("\n".join(detail_lines), encoding="utf-8-sig")
    return detail_lines


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从 PostgreSQL 按 SKC、材质方向和尺寸匹配图片 URL"
    )
    parser.add_argument('--image-version', '--version-name', help='指定数据库已启用版本；省略时按 Excel 材质自动选择唯一启用版本')
    parser.add_argument(
        "input_path",
        nargs="?",
        type=Path,
        help=(
            "待处理的单个 output_feishu_table_data Excel；"
            "不传时批量扫描默认目录"
        ),
    )
    parser.add_argument(
        "--pg-config",
        type=Path,
        default=DEFAULT_PG_CONFIG_PATH,
        help="PostgreSQL 配置文件路径，默认使用脚本同目录的 pg.config",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help=(
            "单文件模式的另存路径；"
            "批量模式不能使用，默认增加“_尺寸图片已匹配”"
        ),
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="直接覆盖输入文件",
    )
    parser.add_argument(
        "--missing-image-output",
        type=Path,
        default=DEFAULT_MISSING_IMAGE_SKCS_PATH,
        help="数据库中没有可用图片记录的 SKC 文本清单路径",
    )
    parser.add_argument(
        "--missing-image-details-output",
        type=Path,
        default=DEFAULT_MISSING_IMAGE_DETAILS_PATH,
        help="按 SKC 记录缺少的具体图片名称的文本报告路径",
    )
    parser.add_argument(
        "--no-group-notification",
        "--no-feishu-notification",
        dest="no_group_notification",
        action="store_true",
        help="不把缺图明细发送到飞书群；默认发送",
    )
    return parser.parse_args()


def discover_input_files(input_path: Path | None) -> list[Path]:
    """传参时处理单文件；未传参时发现全部尚未匹配的源 Excel。"""
    if input_path is not None:
        return [input_path]
    if not DEFAULT_INPUT_DIR.is_dir():
        raise FileNotFoundError(f"默认输入目录不存在：{DEFAULT_INPUT_DIR}")
    return sorted(
        path
        for path in DEFAULT_INPUT_DIR.glob(DEFAULT_INPUT_PATTERN)
        if path.is_file()
        and not path.name.startswith("~$")
        and not path.stem.endswith(MATCHED_FILE_SUFFIX)
    )


def default_output_path(input_path: Path) -> Path:
    """生成稳定的匹配结果文件名，避免重复叠加后缀。"""
    if input_path.stem.endswith(MATCHED_FILE_SUFFIX):
        return input_path
    return input_path.with_name(
        f"{input_path.stem}{MATCHED_FILE_SUFFIX}{input_path.suffix}"
    )


def print_result(result: dict[str, object]) -> None:
    print(f"已保存：{result['output_path']}")
    print(f"已处理 SKC：{result['processed_skcs']}")
    print(f"匹配数据行：{result['matched_rows']}")
    if result["skipped_skcs"]:
        print("跳过的 SKC：")
        for skc, reason in result["skipped_skcs"].items():
            print(f"  {skc}: {reason}")


def send_feishu_text(webhook: str, content: str) -> None:
    """通过飞书自定义机器人发送一条文本消息。"""
    payload = json.dumps(
        {"msg_type": "text", "content": {"text": content}},
        ensure_ascii=False,
    ).encode("utf-8")
    request = Request(
        webhook,
        data=payload,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    try:
        with urlopen(request, timeout=15) as response:
            response_data = json.loads(response.read().decode("utf-8"))
    except (HTTPError, URLError, TimeoutError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"飞书群通知请求失败：{exc}") from exc
    if response_data.get("code", 0) != 0:
        raise RuntimeError(f"飞书群通知发送失败：{response_data}")


def send_missing_image_notifications(detail_lines: list[str]) -> None:
    """存在缺图时，将明细按消息长度分段发送到飞书群。"""
    if not detail_lines:
        return

    header = "【亚马逊批量上架缺图提醒】\n"
    max_content_length = 3500
    message_parts: list[str] = []
    current_lines: list[str] = []
    current_length = len(header)
    for line in detail_lines:
        line_length = len(line) + 1
        if current_lines and current_length + line_length > max_content_length:
            message_parts.append("\n".join(current_lines))
            current_lines = []
            current_length = len(header)
        current_lines.append(line)
        current_length += line_length
    if current_lines:
        message_parts.append("\n".join(current_lines))

    total = len(message_parts)
    for index, message_part in enumerate(message_parts, start=1):
        page_text = f"（{index}/{total}）\n" if total > 1 else ""
        send_feishu_text(
            FEISHU_MISSING_IMAGE_WEBHOOK,
            f"{header}{page_text}{message_part}",
        )


def main() -> None:
    args = parse_args()
    if args.in_place and args.output:
        raise ValueError("--in-place 和 --output 不能同时使用")
    if args.input_path is None and args.output:
        raise ValueError("批量模式不能使用 --output，请传入单个 input_path")

    input_files = discover_input_files(args.input_path)
    if not input_files:
        print(
            f"没有匹配到待处理文件："
            f"{DEFAULT_INPUT_DIR / DEFAULT_INPUT_PATTERN}"
        )
        return

    succeeded_files = []
    failed_files = []
    missing_image_skcs: list[str] = []
    missing_image_details: list[dict[str, str]] = []
    for file_index, input_path in enumerate(input_files, start=1):
        print(
            f"\n========== 处理 [{file_index}/{len(input_files)}]："
            f"{input_path.name} =========="
        )
        if args.in_place:
            output_path = input_path
        elif args.output:
            output_path = args.output
        else:
            output_path = default_output_path(input_path)

        try:
            result = match_rug_images(input_path, output_path, args.pg_config, image_version=args.image_version)
        except Exception as exc:
            failed_files.append((input_path, str(exc)))
            print(f"处理失败，继续下一文件：{input_path.name}，原因：{exc}")
        else:
            succeeded_files.append(Path(result["output_path"]))
            missing_image_skcs.extend(result["missing_image_skcs"])
            missing_image_details.extend(result["missing_image_details"])
            print_result(result)

    unique_missing_image_skcs = list(dict.fromkeys(missing_image_skcs))
    args.missing_image_output.parent.mkdir(parents=True, exist_ok=True)
    args.missing_image_output.write_text(
        "".join(f"{skc}\n" for skc in unique_missing_image_skcs),
        encoding="utf-8",
    )
    print(
        f"数据库中没有可用图片记录的 SKC：{len(unique_missing_image_skcs)} 个，"
        f"清单已保存至：{args.missing_image_output}"
    )

    unique_missing_image_details = list(dict.fromkeys(
        (
            item["skc"],
            item["required_image"],
        )
        for item in missing_image_details
    ))
    details_by_skc: dict[str, list[str]] = {}
    for skc, required_image in unique_missing_image_details:
        details_by_skc.setdefault(skc, []).append(required_image)
    detail_lines = []
    for skc, required_images in details_by_skc.items():
        detail_lines.append(f"SKC: {skc}")
        for required_image in required_images:
            detail_lines.append(f"  - 缺少图片: {required_image}")
        detail_lines.append("")
    args.missing_image_details_output.parent.mkdir(parents=True, exist_ok=True)
    args.missing_image_details_output.write_text(
        "\n".join(detail_lines),
        encoding="utf-8-sig",
    )
    print(
        f"缺图明细：{len(unique_missing_image_details)} 项，"
        f"报告已保存至：{args.missing_image_details_output}"
    )
    if unique_missing_image_details and args.no_group_notification:
        print("已按参数关闭飞书群缺图通知")
    elif unique_missing_image_details:
        try:
            send_missing_image_notifications(detail_lines)
        except RuntimeError as exc:
            print(f"缺图明细发送群通知失败：{exc}")
        else:
            print("缺图明细已发送到飞书群")

    print(
        f"\n批量处理完成：成功 {len(succeeded_files)} 个，"
        f"失败 {len(failed_files)} 个"
    )
    if failed_files:
        failure_message = "；".join(
            f"{path.name}: {reason}" for path, reason in failed_files
        )
        raise RuntimeError(f"以下文件处理失败：{failure_message}")


if __name__ == "__main__":
    main()
