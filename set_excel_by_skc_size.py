r"""按 SKC、材质方向和尺寸，把印花地毯图片 URL 匹配到上架数据 Excel。

默认测试文件：
    D:\项目文件\AI自动上架\
    output_feishu_table_data_亚马逊--冬豚--北美（子账号）_AR1002-NEW.xlsx

图片数据文件约定：
    D:\NAS_download\{SKC}\{SKC}data.xlsx

默认另存为“原文件名_尺寸图片已匹配.xlsx”；传入 --in-place 可覆盖源文件。
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

import pandas as pd


DEFAULT_INPUT_PATH = Path(
    r"D:\项目文件\AI自动上架\output_feishu_table_data_亚马逊-北蓉-北美（子账号）_Minimalist-AR3-POD.xlsx"
)
DEFAULT_INPUT_DIR = DEFAULT_INPUT_PATH.parent
DEFAULT_INPUT_PATTERN = "output_feishu_table_data_*.xlsx"
MATCHED_FILE_SUFFIX = "_尺寸图片已匹配"
DEFAULT_IMAGE_ROOT = Path(r"D:\NAS_download")
RUG_MATERIAL = "印花地毯"

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
    file_names = image_df["文件名"].map(lambda value: normalize_text(value).upper())
    for keyword in keywords:
        matched = image_df[file_names.str.contains(keyword.upper(), regex=False)]
        url = first_url(matched)
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
    image_root: Path = DEFAULT_IMAGE_ROOT,
) -> dict[str, object]:
    """读取目标 Excel，并给材质方向为印花地毯的 SKC 匹配尺寸图片。"""
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
    skipped_skcs: dict[str, str] = {}
    matched_rows = 0

    skcs = list(dict.fromkeys(value for value in target_df["SKC"] if value))
    for skc in skcs:
        skc_rows = target_df[target_df["SKC"] == skc]
        materials = list(dict.fromkeys(
            value for value in skc_rows["材质方向"].tolist() if value
        ))
        if materials != [RUG_MATERIAL]:
            skipped_skcs[skc] = (
                f"材质方向不是“{RUG_MATERIAL}”：{materials or ['<空>']}"
            )
            continue

        image_data_path = image_root / skc / f"{skc}data.xlsx"
        if not image_data_path.is_file():
            skipped_skcs[skc] = f"图片数据文件不存在：{image_data_path}"
            continue

        image_df = pd.read_excel(image_data_path)
        if "SKC" in image_df.columns:
            image_df = image_df[
                image_df["SKC"].map(normalize_text) == skc
            ].copy()
        if image_df.empty:
            skipped_skcs[skc] = "图片数据文件中没有该 SKC 的记录"
            continue

        pool = build_image_pool(image_df)
        for row_index in skc_rows.index:
            size = normalize_size(target_df.at[row_index, "size_text"])
            if not size:
                continue
            target_df.loc[row_index, IMAGE_COLUMNS] = ""
            assign_urls(
                target_df,
                row_index,
                image_values_for_size(size, pool),
            )
            matched_rows += 1
        processed_skcs.append(skc)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    target_df.to_excel(output_path, index=False)
    return {
        "output_path": output_path,
        "processed_skcs": processed_skcs,
        "skipped_skcs": skipped_skcs,
        "matched_rows": matched_rows,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="按 SKC 和尺寸给印花地毯上架数据匹配图片 URL"
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
        "--image-root",
        type=Path,
        default=DEFAULT_IMAGE_ROOT,
        help="SKC 图片数据根目录",
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
            result = match_rug_images(input_path, output_path, args.image_root)
        except Exception as exc:
            failed_files.append((input_path, str(exc)))
            print(f"处理失败，继续下一文件：{input_path.name}，原因：{exc}")
        else:
            succeeded_files.append(Path(result["output_path"]))
            print_result(result)

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
