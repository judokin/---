r"""从 PostgreSQL 读取已启用图片版本，按 SKC 和尺寸填入上架图片 URL。

默认批量目录：D:\项目文件\AI自动上架\

图片数据来源：
    PostgreSQL public.ods_fbm_image_upload
    数据库连接配置默认读取脚本同目录的 pg.config。

默认另存为“原文件名_尺寸图片已匹配.xlsx”；传入 --in-place 可覆盖源文件。
必须传入 --image-version（如 "2026圣诞专题v1"）；图片名称及顺序仅取自数据库。
"""

from __future__ import annotations

import argparse
from contextlib import closing
import json
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import unquote, urlsplit
from urllib.request import Request, urlopen

import pandas as pd
import psycopg2


DEFAULT_INPUT_DIR = Path(r"D:\项目文件\AI自动上架")
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
    """统一图片文件名和数据库关键词的大小写、空格及乘号。"""
    return (
        normalize_text(value)
        .upper()
        .replace("×", "X")
        .replace("＋", "+")
        .replace(" ", "")
    )


def first_url(rows: pd.DataFrame) -> str:
    """取得数据集合中的第一个非空图片地址。"""
    if rows.empty or "图片地址" not in rows.columns:
        return ""
    for value in rows["图片地址"].tolist():
        url = normalize_text(value)
        if url:
            return url
    return ""


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


def load_image_data_from_pg(
    skcs: list[str],
    config_path: Path,
) -> dict[str, pd.DataFrame]:
    """读取目标 SKC 的可用图片，只还原文件名，不推断图片类型或内置顺序。"""
    if not skcs:
        return {}

    query = """
        SELECT skc, file_name, url
        FROM public.ods_fbm_image_upload
        WHERE skc = ANY(%s)
          AND status = 'success'
          AND COALESCE(url, '') <> ''
          AND (url_expire_at IS NULL OR url_expire_at > NOW())
        ORDER BY skc, id
    """
    with closing(psycopg2.connect(**load_pg_config(config_path))) as connection:
        connection.set_session(readonly=True)
        with connection.cursor() as cursor:
            cursor.execute("SET LOCAL statement_timeout = '15s'")
            cursor.execute(query, (skcs,))
            rows = cursor.fetchall()

    records = [
        {
            "SKC": normalize_text(skc),
            "文件名": decoded_image_name(url, file_name),
            "图片地址": normalize_text(url),
        }
        for skc, file_name, url in rows
    ]
    if not records:
        return {}
    image_df = pd.DataFrame(records)
    return {
        skc: group.reset_index(drop=True)
        for skc, group in image_df.groupby("SKC", sort=False)
    }


def load_enabled_image_version(
    version_name: str,
    pg_config_path: Path = DEFAULT_PG_CONFIG_PATH,
) -> dict[str, object]:
    """读取唯一启用的 Amazon 印花地毯版本；一次批处理使用同一规则快照。"""
    version_name = normalize_text(version_name)
    if not version_name:
        raise ValueError("图片版本名称不能为空")
    with closing(psycopg2.connect(**load_pg_config(pg_config_path))) as connection:
        connection.set_session(readonly=True)
        with connection.cursor() as cursor:
            cursor.execute("SET LOCAL statement_timeout = '15s'")
            cursor.execute("""
                SELECT rules_schema_version, rules_config
                FROM public.listing_image_versions
                WHERE version_name = %s AND is_enabled = TRUE
                  AND platform = %s AND material_direction = %s
            """, (version_name, 'amazon', RUG_MATERIAL))
            rows = cursor.fetchall()
    if len(rows) != 1:
        raise ValueError(
            f"需要唯一的已启用 Amazon / {RUG_MATERIAL} / {version_name} 图片版本，"
            f"实际找到 {len(rows)} 个；请检查版本名称和启用状态"
        )
    schema, config = rows[0]
    if (schema != 1 or not isinstance(config, dict)
            or config.get('slot_order') != IMAGE_COLUMNS
            or not isinstance(config.get('groups'), list) or not config['groups']):
        raise ValueError(f"图片版本 {version_name} 规则为空或结构不兼容")

    rules_by_size = {}
    for group in config['groups']:
        if (not isinstance(group, dict) or not isinstance(group.get('sizes'), list)
                or not group['sizes'] or not isinstance(group.get('rules'), dict)
                or set(group['rules']) != set(IMAGE_COLUMNS)):
            raise ValueError(f"图片版本 {version_name} 尺寸或图片列配置不完整")
        for slot in IMAGE_COLUMNS:
            rule = group['rules'][slot]
            if (not isinstance(rule, dict) or not isinstance(rule.get('primary'), str)
                    or not rule['primary'].strip() or not isinstance(rule.get('type'), str)
                    or not isinstance(rule.get('alternatives'), list)
                    or not all(isinstance(name, str) and name.strip() for name in rule['alternatives'])):
                raise ValueError(f"图片版本 {version_name} 的 {slot} 文件名规则无效")
        for raw_size in group['sizes']:
            if not isinstance(raw_size, str) or not normalize_size(raw_size):
                raise ValueError(f"图片版本 {version_name} 包含空尺寸或无效尺寸")
            size = normalize_size(raw_size)
            if size in rules_by_size:
                raise ValueError(f"图片版本 {version_name} 重复配置尺寸：{size}")
            rules_by_size[size] = group['rules']
    return {
        'version_name': version_name,
        'platform': 'amazon',
        'material_direction': RUG_MATERIAL,
        'rules_by_size': rules_by_size,
    }


def version_image_values(
    image_df: pd.DataFrame | None,
    rules: dict[str, dict],
) -> tuple[dict[str, str], dict[str, str]]:
    """完全按数据库目标列的名称规则取图；source 历史信息不参与匹配。"""
    values = dict.fromkeys(IMAGE_COLUMNS, '')
    required_names = {}
    has_images = image_df is not None and not image_df.empty
    normalized_names = (
        image_df['文件名'].map(lambda value: normalize_image_name(value).replace('封面图', ''))
        if has_images else None
    )
    for slot in IMAGE_COLUMNS:
        rule = rules[slot]
        if rule['primary'].strip() == '留空' or rule['type'].strip() == '不写入':
            continue
        names = [rule['primary'], *rule['alternatives']]
        required_names[slot] = ' 或 '.join(dict.fromkeys(names))
        if not has_images:
            continue
        for name in names:
            keyword = normalize_image_name(name).replace('封面图', '')
            if keyword:
                values[slot] = first_url(image_df[normalized_names.str.contains(keyword, regex=False)])
            if values[slot]:
                break
    return values, required_names


def match_rug_images(
    input_path: Path,
    output_path: Path,
    pg_config_path: Path = DEFAULT_PG_CONFIG_PATH,
    image_version: dict[str, object] | None = None,
) -> dict[str, object]:
    """读取目标 Excel，必须使用已加载的数据库版本给 SKC 匹配图片。"""
    if image_version is None:
        raise ValueError("必须通过 --image-version 指定数据库中已启用的图片版本")
    if not input_path.is_file():
        raise FileNotFoundError(f"目标 Excel 不存在：{input_path}")

    target_df = pd.read_excel(input_path)
    required_columns = {"SKC", "材质方向", "size_text"}
    missing_columns = sorted(required_columns - set(target_df.columns))
    if missing_columns:
        raise ValueError(f"目标 Excel 缺少字段：{missing_columns}")

    for column in IMAGE_COLUMNS:
        if column not in target_df.columns:
            target_df[column] = ""
        target_df[column] = target_df[column].astype(object)
    target_df["SKC"] = target_df["SKC"].map(normalize_text)
    target_df["材质方向"] = target_df["材质方向"].map(normalize_text)

    # 圣诞模式生成的 Excel 材质仍为印花地毯，兼容未归一化的历史文件。
    version_materials = {image_version['material_direction']}
    if image_version['material_direction'] == RUG_MATERIAL:
        version_materials.add(CHRISTMAS_RUG_MATERIAL)
    skcs = list(dict.fromkeys(value for value in target_df["SKC"] if value))
    processed_skcs: list[str] = []
    missing_image_skcs: list[str] = []
    missing_image_details: list[dict[str, str]] = []
    skipped_skcs: dict[str, str] = {}
    matched_rows = 0

    eligible_rows_by_skc = {}
    for skc in skcs:
        skc_rows = target_df[target_df["SKC"] == skc]
        materials = list(dict.fromkeys(
            value for value in skc_rows["材质方向"].tolist() if value
        ))
        if not materials or any(material not in version_materials for material in materials):
            skipped_skcs[skc] = f"材质方向与图片版本不匹配：{materials or ['<空>']}"
            continue
        sized_rows = skc_rows[skc_rows['size_text'].map(normalize_size).ne('')]
        if sized_rows.empty:
            continue
        unknown_sizes = sorted({
            normalize_size(size) for size in sized_rows['size_text']
            if normalize_size(size) not in image_version['rules_by_size']
        })
        if unknown_sizes:
            raise ValueError(
                f"图片版本 {image_version['version_name']} 未配置 Excel 尺寸：{unknown_sizes}；SKC：{skc}"
            )
        eligible_rows_by_skc[skc] = sized_rows

    image_data_by_skc = load_image_data_from_pg(list(eligible_rows_by_skc), pg_config_path)
    for skc, skc_rows in eligible_rows_by_skc.items():
        image_df = image_data_by_skc.get(skc)
        has_images = image_df is not None and not image_df.empty
        if not has_images:
            missing_image_skcs.append(skc)
            skipped_skcs[skc] = "数据库 public.ods_fbm_image_upload 中没有可用图片记录"
        for row_index in skc_rows.index:
            size = normalize_size(target_df.at[row_index, 'size_text'])
            values, required_names = version_image_values(
                image_df, image_version['rules_by_size'][size]
            )
            # 清掉旧版本链接，包括配置留空和当前缺失的图片。
            target_df.loc[row_index, IMAGE_COLUMNS] = ''
            assign_urls(target_df, row_index, values)
            for slot, name in required_names.items():
                if not values[slot]:
                    missing_image_details.append({
                        'skc': skc, 'material': image_version['material_direction'],
                        'size': size, 'target_column': slot, 'required_image': name,
                    })
            if has_images:
                matched_rows += 1
        if has_images:
            processed_skcs.append(skc)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    target_df.to_excel(output_path, index=False)
    return {
        "output_path": output_path,
        "processed_skcs": processed_skcs,
        "missing_image_skcs": missing_image_skcs,
        "missing_image_details": missing_image_details,
        "skipped_skcs": skipped_skcs,
        "matched_rows": matched_rows,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从 PostgreSQL 按 SKC、材质方向和尺寸匹配图片 URL"
    )
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
        "--image-version", "--version-name",
        dest="image_version",
        required=True,
        help="必须指定已启用的 Amazon 印花地毯图片版本，如 2026圣诞专题v1",
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

    image_version = load_enabled_image_version(args.image_version, args.pg_config)
    print(
        f"使用已启用图片版本：{image_version['version_name']}，"
        f"平台：{image_version['platform']}，材质：{image_version['material_direction']}，"
        f"尺寸：{', '.join(image_version['rules_by_size'])}"
    )

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
            result = match_rug_images(
                input_path, output_path, args.pg_config, image_version=image_version
            )
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
