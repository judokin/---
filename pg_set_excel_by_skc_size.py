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
from pathlib import Path
from typing import Iterable
from urllib.parse import unquote, urlsplit

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
RUG_MATERIAL = "印花地毯"
KITCHEN_MAT_MATERIAL = "仿羊绒厨房垫"
OUTDOOR_MAT_MATERIAL = "三明治户外垫"
SUPPORTED_MATERIALS = {
    RUG_MATERIAL,
    KITCHEN_MAT_MATERIAL,
    OUTDOOR_MAT_MATERIAL,
}

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

KITCHEN_MAT_IMAGE_NAMES = {
    "主图片 URL": "20X32+20X48厨房垫",
    "其他图片 URL1": "封面图2X5厨房 / 2X5厨房",
    "其他图片 URL2": "封面图2X3门口 / 2X3门口",
    "其他图片 URL3": "2X5走廊",
    "其他图片 URL4": "120X170单椅",
    "其他图片 URL5": "2X5尺寸白底图 / 白底图2X5 / 2X5白底图",
    "样本图片 URL": "样本图片 / SWATCH / SWITCH / SWICH",
}

OUTDOOR_MAT_IMAGE_NAMES = {
    "主图片 URL": "5X7户外庭院",
    "其他图片 URL1": "白底图5X7",
    "其他图片 URL2": "3X5户外门口",
    "其他图片 URL3": "2.5X8户外走廊",
    "其他图片 URL4": "封面图5X7客厅",
    "其他图片 URL5": "120X170单椅",
    "其他图片 URL6": "5X7卧室",
    "样本图片 URL": "样本图片 / SWATCH / SWITCH / SWICH",
}


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
    def normalize_image_name(value: object) -> str:
        return (
            normalize_text(value)
            .upper()
            .replace("×", "X")
            .replace("＋", "+")
            .replace(" ", "")
        )

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
    return pick_named_url(image_df, "SWATCH", "SWITCH", "SWICH")


def kitchen_mat_image_values(image_df: pd.DataFrame) -> dict[str, str]:
    """按文件名给仿羊绒厨房垫匹配固定顺序的图片 URL。"""
    required_columns = {"文件名", "图片地址"}
    missing_columns = sorted(required_columns - set(image_df.columns))
    if missing_columns:
        raise ValueError(f"图片数据文件缺少字段：{missing_columns}")

    return {
        "主图片 URL": pick_named_url(image_df, "20X32+20X48厨房垫"),
        "其他图片 URL1": pick_named_url(
            image_df,
            "封面图2X5厨房",
            "2X5厨房",
        ),
        "其他图片 URL2": pick_named_url(
            image_df,
            "封面图2X3门口",
            "2X3门口",
        ),
        "其他图片 URL3": pick_named_url(image_df, "2X5走廊"),
        "其他图片 URL4": pick_named_url(image_df, "120X170单椅"),
        "其他图片 URL5": pick_named_url(
            image_df,
            "2X5尺寸白底图",
            "白底图2X5",
            "2X5白底图",
        ),
        "其他图片 URL6": "",
        "其他图片 URL7": "",
        "其他图片 URL8": "",
        "样本图片 URL": pick_sample_url(image_df),
    }


def outdoor_mat_image_values(image_df: pd.DataFrame) -> dict[str, str]:
    """按文件名给三明治户外垫匹配固定顺序的图片 URL。"""
    required_columns = {"文件名", "图片地址"}
    missing_columns = sorted(required_columns - set(image_df.columns))
    if missing_columns:
        raise ValueError(f"图片数据文件缺少字段：{missing_columns}")

    return {
        "主图片 URL": pick_named_url(image_df, "5X7户外庭院"),
        "其他图片 URL1": pick_named_url(image_df, "白底图5X7"),
        "其他图片 URL2": pick_named_url(image_df, "3X5户外门口"),
        "其他图片 URL3": pick_named_url(image_df, "2.5X8户外走廊"),
        "其他图片 URL4": pick_named_url(image_df, "封面图5X7客厅"),
        "其他图片 URL5": pick_named_url(image_df, "120X170单椅"),
        "其他图片 URL6": pick_named_url(image_df, "5X7卧室"),
        "其他图片 URL7": "",
        "其他图片 URL8": "",
        "样本图片 URL": pick_sample_url(image_df),
    }


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
    if any(keyword in normalized_name for keyword in ("SWATCH", "SWITCH", "SWICH")):
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
        with connection.cursor() as cursor:
            cursor.execute(query, (skcs,))
            rows = cursor.fetchall()

    records = []
    for skc, file_name, url, size_label, image_type in rows:
        restored_name = decoded_image_name(url, file_name)
        records.append({
            "SKC": normalize_text(skc),
            "文件名": restored_name,
            "图片地址": normalize_text(url),
            "尺寸": normalize_size(size_label),
            "图片类型": infer_image_type(restored_name, image_type),
        })

    if not records:
        return {}
    image_df = pd.DataFrame(records)
    return {
        skc: group.reset_index(drop=True)
        for skc, group in image_df.groupby("SKC", sort=False)
    }


def build_image_pool(image_df: pd.DataFrame) -> dict[str, object]:
    """参考 set_excel.py，按尺寸和图片类型整理印花地毯图片。"""
    required_columns = {"文件名", "图片地址", "尺寸", "图片类型"}
    missing_columns = sorted(required_columns - set(image_df.columns))
    if missing_columns:
        raise ValueError(f"图片数据文件缺少字段：{missing_columns}")

    image_df = image_df.copy()
    image_df["标准尺寸"] = image_df["尺寸"].map(normalize_size)
    image_df["标准图片类型"] = image_df["图片类型"].map(normalize_text)

    def by(size_values: Iterable[str], image_type: str) -> pd.DataFrame:
        return image_df[
            image_df["标准尺寸"].isin(set(size_values))
            & (image_df["标准图片类型"] == image_type)
        ]

    main_57 = urls(by(["5X7", "120X170"], "主图片"))
    other_57 = urls(by(["5X7", "120X170"], "其他图片"))
    main_25 = urls(by(["2X5"], "主图片"))
    other_25 = urls(by(["2X5"], "其他图片"))

    main_by_size = {
        size: first_url(by([size], "主图片"))
        for size in image_df["标准尺寸"].dropna().unique()
    }
    white_by_size = {
        size: first_url(by([size], "白底图"))
        for size in image_df["标准尺寸"].dropna().unique()
    }

    return {
        "sample": first_url(by(["特殊"], "样本图片")),
        "main_57": main_57,
        "other_57": other_57,
        "main_25": main_25,
        "other_25": other_25,
        "main_by_size": main_by_size,
        "white_by_size": white_by_size,
        "main_23_named": pick_named_url(image_df, "2X3门口", "2X3玄关"),
    }


def value_at(values: list[str], index: int) -> str:
    """安全读取列表项；图片不足时返回空字符串。"""
    return values[index] if index < len(values) else ""


def required_rug_image_names(size: str) -> dict[str, str]:
    """返回印花地毯指定尺寸各图片列期望的图片名称。"""
    names = {"样本图片 URL": "特殊尺寸的样本图片"}
    if size == "2X3":
        names.update({
            "主图片 URL": "2X3门口 / 2X3玄关 / 2X3主图片",
            "其他图片 URL1": "2X3白底图",
            "其他图片 URL2": "5X7或120X170主图片1",
            "其他图片 URL3": "5X7或120X170其他图片1",
            "其他图片 URL4": "5X7或120X170其他图片2",
            "其他图片 URL5": "2X5主图片1",
            "其他图片 URL6": "2X5其他图片1",
        })
    elif size in {"2X5", "2X6", "2X7"}:
        names.update({
            "主图片 URL": f"{size}主图片",
            "其他图片 URL1": "2X5其他图片1",
            "其他图片 URL2": f"{size}白底图 / 2X5白底图",
            "其他图片 URL3": "5X7或120X170主图片1",
            "其他图片 URL4": "5X7或120X170其他图片1",
            "其他图片 URL5": "5X7或120X170其他图片2",
            "其他图片 URL6": "2X3门口 / 2X3玄关 / 2X3主图片",
        })
    elif size in {"5X7", "8X10"}:
        names.update({
            "主图片 URL": "5X7或120X170主图片1",
            "其他图片 URL1": "5X7或120X170其他图片1",
            "其他图片 URL2": "5X7或120X170其他图片2",
            "其他图片 URL3": "5X7白底图",
            "其他图片 URL4": "2X5主图片1",
            "其他图片 URL5": "2X5其他图片1",
            "其他图片 URL6": "2X3门口 / 2X3玄关 / 2X3主图片",
        })
    else:
        names["主图片 URL"] = f"{size}主图片"
    return names


def collect_missing_images(
    skc: str,
    material: str,
    size: str,
    values: dict[str, str],
) -> list[dict[str, str]]:
    """收集一个 SKC 当前尺寸按规则应有但未匹配到的图片。"""
    if material == KITCHEN_MAT_MATERIAL:
        required_names = KITCHEN_MAT_IMAGE_NAMES
    elif material == OUTDOOR_MAT_MATERIAL:
        required_names = OUTDOOR_MAT_IMAGE_NAMES
    else:
        required_names = required_rug_image_names(size)

    return [
        {
            "skc": skc,
            "material": material,
            "size": size,
            "target_column": target_column,
            "required_image": required_image,
        }
        for target_column, required_image in required_names.items()
        if not normalize_text(values.get(target_column, ""))
    ]


def image_values_for_size(size: str, pool: dict[str, object]) -> dict[str, str]:
    """生成某个商品尺寸应写入的图片列，规则与 set_excel.py 保持一致。"""
    main_57 = pool["main_57"]
    other_57 = pool["other_57"]
    main_25 = pool["main_25"]
    other_25 = pool["other_25"]
    main_by_size = pool["main_by_size"]
    white_by_size = pool["white_by_size"]
    main_23 = pool["main_23_named"] or main_by_size.get("2X3", "")

    values = {"样本图片 URL": pool["sample"]}

    if size == "2X3":
        values.update({
            "主图片 URL": main_23,
            "其他图片 URL1": white_by_size.get("2X3", ""),
            "其他图片 URL2": value_at(main_57, 0),
            "其他图片 URL3": value_at(other_57, 0),
            "其他图片 URL4": value_at(other_57, 1),
            "其他图片 URL5": value_at(main_25, 0),
            "其他图片 URL6": value_at(other_25, 0),
        })
    elif size in {"2X5", "2X6", "2X7"}:
        values.update({
            "主图片 URL": main_by_size.get(size, ""),
            "其他图片 URL1": value_at(other_25, 0),
            "其他图片 URL2": white_by_size.get(size, "")
            or white_by_size.get("2X5", ""),
            "其他图片 URL3": value_at(main_57, 0),
            "其他图片 URL4": value_at(other_57, 0),
            "其他图片 URL5": value_at(other_57, 1),
            "其他图片 URL6": main_23,
        })
    elif size in {"5X7", "8X10"}:
        values.update({
            "主图片 URL": value_at(main_57, 0),
            "其他图片 URL1": value_at(other_57, 0),
            "其他图片 URL2": value_at(other_57, 1),
            "其他图片 URL3": white_by_size.get("5X7", ""),
            "其他图片 URL4": value_at(main_25, 0),
            "其他图片 URL5": value_at(other_25, 0),
            "其他图片 URL6": main_23,
        })
    else:
        values["主图片 URL"] = main_by_size.get(size, "")

    return values


def match_rug_images(
    input_path: Path,
    output_path: Path,
    pg_config_path: Path = DEFAULT_PG_CONFIG_PATH,
) -> dict[str, object]:
    """读取目标 Excel，并按材质方向给 SKC 匹配图片。"""
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

    processed_skcs: list[str] = []
    missing_image_skcs: list[str] = []
    missing_image_details: list[dict[str, str]] = []
    skipped_skcs: dict[str, str] = {}
    matched_rows = 0

    skcs = list(dict.fromkeys(value for value in target_df["SKC"] if value))
    image_data_by_skc = load_image_data_from_pg(skcs, pg_config_path)
    for skc in skcs:
        skc_rows = target_df[target_df["SKC"] == skc]
        materials = list(dict.fromkeys(
            value for value in skc_rows["材质方向"].tolist() if value
        ))
        if len(materials) != 1 or materials[0] not in SUPPORTED_MATERIALS:
            skipped_skcs[skc] = (
                f"材质方向不受支持或不唯一：{materials or ['<空>']}"
            )
            continue
        material = materials[0]

        image_df = image_data_by_skc.get(skc)
        if image_df is None or image_df.empty:
            missing_image_skcs.append(skc)
            for row_index in skc_rows.index:
                size = normalize_size(target_df.at[row_index, "size_text"])
                if size:
                    missing_image_details.extend(
                        collect_missing_images(skc, material, size, {})
                    )
            skipped_skcs[skc] = (
                "数据库 public.ods_fbm_image_upload 中没有可用图片记录"
            )
            continue

        if material == RUG_MATERIAL:
            pool = build_image_pool(image_df)
            fixed_values = None
        elif material == KITCHEN_MAT_MATERIAL:
            pool = None
            fixed_values = kitchen_mat_image_values(image_df)
        else:
            pool = None
            fixed_values = outdoor_mat_image_values(image_df)

        for row_index in skc_rows.index:
            size = normalize_size(target_df.at[row_index, "size_text"])
            if not size:
                continue
            values = (
                image_values_for_size(size, pool)
                if material == RUG_MATERIAL
                else fixed_values
            )
            target_df.loc[row_index, IMAGE_COLUMNS] = ""
            assign_urls(target_df, row_index, values)
            missing_image_details.extend(
                collect_missing_images(skc, material, size, values)
            )
            matched_rows += 1
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
            result = match_rug_images(input_path, output_path, args.pg_config)
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
            item["material"],
            item["size"],
            item["target_column"],
            item["required_image"],
        )
        for item in missing_image_details
    ))
    details_by_skc: dict[str, list[tuple[str, str, str, str]]] = {}
    for skc, material, size, target_column, required_image in unique_missing_image_details:
        details_by_skc.setdefault(skc, []).append(
            (material, size, target_column, required_image)
        )
    detail_lines = []
    for skc, details in details_by_skc.items():
        detail_lines.append(f"SKC: {skc}")
        for material, size, target_column, required_image in details:
            detail_lines.append(
                f"  - 材质方向: {material}；尺寸: {size}；"
                f"字段: {target_column}；缺少图片: {required_image}"
            )
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
