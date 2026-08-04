# 使用提醒:
# 1. xbot包提供软件自动化、数据表格、Excel、日志、AI等功能
# 2. package包提供访问当前应用数据的功能，如获取元素、访问全局变量、获取资源文件等功能
# 3. 当此模块作为流程独立运行时执行main函数
# 4. 可视化流程中可以通过"调用模块"的指令使用此模块
try:
    import xbot
    from xbot import print, sleep
    from .import package
    from .package import variables as glv
except:
    glv = {}
    pass
import pandas as pd
import xlwings as xw
from pathlib import Path

# ================================
#  配置部分（请根据实际情况修改路径）
# ================================
input_dir = Path(r'D:\项目文件\AI自动上架')
input_pattern = '*尺寸图片已匹配.xlsx'
template_dir = Path(r'D:\项目文件\AI自动上架\模板文件')
output_path = input_dir / '印花地毯-模板文件_v2__工作文件.xlsm'

_active_app = None
_active_wb = None
_active_work_path = None

# skc_srt = open(r"D:\NAS_download\\nas_path.txt", "r", encoding="utf-8").read()
# glv['gvar_img_path'] = skc_srt
# skc_src_list = []
# for ss in skc_srt.split(","):
#     skc_src_list.append(ss.split("\\")[-1])
# print(skc_src_list)
# glv['gvar_skc'] = skc_srt.split("\\")[-1]

# output_skc_path = r'D:\项目文件\AI自动上架\\'
# print("output_skc_path =", output_skc_path)
start_row = 7  # 数据写入起始行

# ================================
# 辅助函数：列字母 → 列号
# ================================
def col_letter_to_num(column_letter):
    """Excel列字母 → 从1开始的列号，例如 A→1、Z→26、AA→27。"""
    normalized = str(column_letter).strip().upper()
    if not normalized or not normalized.isalpha():
        raise ValueError(f"无效的 Excel 列字母：{column_letter}")
    result = 0
    for char in normalized:
        result = result * 26 + ord(char) - ord('A') + 1
    return result

def cleanup_active_excel():
    """仅清理本脚本创建的 Excel 实例和工作文件。"""
    global _active_app, _active_wb, _active_work_path
    if _active_wb is not None:
        try:
            _active_wb.close()
        except Exception:
            pass
        _active_wb = None
    if _active_app is not None:
        try:
            _active_app.quit()
        except Exception:
            pass
        _active_app = None
    if _active_work_path is not None:
        try:
            Path(_active_work_path).unlink(missing_ok=True)
        except Exception:
            pass
        _active_work_path = None


def process_file(data_path):
    # ================================
    # 主程序
    # ================================
    global _active_app, _active_wb, _active_work_path
    import datetime

    # 步骤1: 读取数据
    df_data = pd.read_excel(data_path)
    data_columns = df_data.columns.tolist()

    # Header 名称是唯一的数据定位依据，不能依赖 Excel 中的列顺序。
    header_to_idx = {}
    for idx, column_name in enumerate(data_columns):
        normalized_name = str(column_name).strip()
        if normalized_name in header_to_idx:
            raise ValueError(f"源数据存在重复 Header：{normalized_name}")
        header_to_idx[normalized_name] = idx

    def get_required_value(row_idx, header_name):
        if header_name not in header_to_idx:
            raise ValueError(f"源数据缺少生成文件名所需的 Header：{header_name}")
        if row_idx >= len(df_data):
            raise ValueError(f"源数据不足 {row_idx + 1} 行，无法读取 Header：{header_name}")
        value = df_data.iloc[row_idx, header_to_idx[header_name]]
        if pd.isna(value) or str(value).strip() == "":
            raise ValueError(f"源数据第 {row_idx + 2} 行的 {header_name} 为空")
        return str(value).strip()

    def get_unique_required_value(header_name):
        """取得指定 Header 的唯一非空值，防止一次任务混用多个店铺模板。"""
        if header_name not in header_to_idx:
            raise ValueError(f"源数据缺少选择模板所需的 Header：{header_name}")

        values = []
        for value in df_data.iloc[:, header_to_idx[header_name]].tolist():
            if pd.isna(value) or str(value).strip() == "":
                continue
            normalized_value = str(value).strip()
            if normalized_value not in values:
                values.append(normalized_value)

        if not values:
            raise ValueError(f"源数据的 {header_name} 全部为空，无法选择模板")
        if len(values) > 1:
            raise ValueError(
                f"源数据包含多个 {header_name}：{values}，一次任务只能使用一个店铺模板"
            )
        return values[0]

    # 根据“FBM上架店铺”动态选择对应的账号模板。
    store_name = get_unique_required_value("FBM上架店铺")
    template_path = template_dir / f"印花地毯-模板文件v2_{store_name}.xlsm"
    if not template_path.is_file():
        raise FileNotFoundError(
            f"未找到 FBM 上架店铺“{store_name}”对应的模板：{template_path}"
        )
    print(f"FBM上架店铺：{store_name}")
    print(f"匹配模板：{template_path}")

    # 写入模板前按“卖家 SKU”去重，保留第一次出现的记录和原始顺序。
    sku_header = "卖家 SKU"
    if sku_header not in header_to_idx:
        raise ValueError(f"源数据缺少去重所需的 Header：{sku_header}")

    sku_values = df_data.iloc[:, header_to_idx[sku_header]]
    normalized_skus = sku_values.map(
        lambda value: "" if pd.isna(value) else str(value).strip()
    )
    empty_sku_rows = [
        row_idx + 2
        for row_idx, sku in enumerate(normalized_skus.tolist())
        if sku == ""
    ]
    if empty_sku_rows:
        raise ValueError(
            f"源数据以下 Excel 行的 {sku_header} 为空，无法安全去重：{empty_sku_rows}"
        )

    duplicate_mask = normalized_skus.duplicated(keep="first")
    original_rows_count = len(df_data)
    df_data = df_data.loc[~duplicate_mask].reset_index(drop=True)
    removed_rows_count = original_rows_count - len(df_data)
    print(
        f"按 {sku_header} 去重：原 {original_rows_count} 行，"
        f"删除重复 {removed_rows_count} 行，保留 {len(df_data)} 行（顺序不变）"
    )

    data_values = df_data.values.tolist()
    print(df_data)

    # 生成带时间戳的输出路径；文件名字段也按 Header 获取。
    timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    width_unit = get_required_value(1, "商品宽度单位")
    seller_sku = get_required_value(0, "卖家 SKU")
    invalid_filename_chars = '<>:"/\\|?*'
    safe_store_name = "".join(
        "_" if char in invalid_filename_chars else char
        for char in store_name
    ).strip().rstrip(".")
    if not safe_store_name:
        raise ValueError(f"FBM上架店铺无法用于生成文件名：{store_name}")
    output_name_prefix = f'{width_unit}_{seller_sku}_{safe_store_name}_'
    existing_outputs = sorted(
        path for path in input_dir.glob('*.xlsm')
        if path.name.startswith(output_name_prefix)
        and path.name.endswith('.xlsm')
        and path.name != output_path.name
    )
    if existing_outputs:
        print(
            f"已生成，跳过：{Path(data_path).name} → "
            f"{existing_outputs[-1].name}"
        )
        return None

    output_skc_path = input_dir / (
        f'{output_name_prefix}{timestamp}.xlsm'
    )

    # 步骤2: 复制模板并打开
    import shutil
    shutil.copy(template_path, output_path)
    _active_work_path = output_path
    app = xw.App(visible=False)  # 调试时可改为 True
    _active_app = app
    wb = app.books.open(str(output_path))
    _active_wb = wb
    sheet_names = [sheet_item.name for sheet_item in wb.sheets]
    template_sheet_name = next(
        (
            candidate
            for candidate in ('Template', '模板')
            if candidate in sheet_names
        ),
        None,
    )
    if template_sheet_name is None:
        raise ValueError(
            f"模板文件中找不到 Template/模板 工作表，实际工作表：{sheet_names}"
        )
    sheet = wb.sheets[template_sheet_name]
    print(f"写入工作表：{template_sheet_name}")

    # ================================
    # 飞书表格 → 模板列映射
    # 键：飞书数据 Header 名称
    # 值：新模板中的目标列字母
    #
    # 字典顺序同时定义同一模板列的覆盖优先级：靠后的字段优先。
    # 因此即使源 Excel 调整列顺序，写入结果也保持一致。
    # ================================
    feishu_to_template_mapping = {
        "商品类型": "B",
        "父条目的库存单位": "E",
        "pattern": "BZ",
        "不含税价目表": "ED",
        "设计": "AQ",
        "搜索关键词": "AK",
        "产品类型关键字": "L",
        "关于此艺术品": "AE",
        "Product Description": "AE",
        "制造商": "R",
        "卖家 SKU": "A",
        "品牌": "I",
        "商品名称": "G",
        "商品编号类型": "J",
        "商品 ID": "K",
        "Back Material Type": "CR",
        "Care Instructions1": "BM",
        "Care Instructions2": "BN",
        "Construction Type": "BA",
        "Item Shape": "BF",
        "Material": "AR",
        "Material Features1": "BS",
        "Material Features2": "BT",
        "Material Features3": "BU",
        "Model Name": "Q",
        "Pile Height": "BX",
        "Recommended Uses For Product1": "CL",
        "Recommended Uses For Product2": "CM",
        "Recommended Uses For Product3": "CN",
        "Recommended Uses For Product4": "CO",
        "Recommended Uses For Product5": "CP",
        "Room Type1": "CX",
        "Room Type2": "CY",
        "Room Type3": "CZ",
        "Room Type4": "DA",
        "Room Type5": "DB",
        "Size": "AZ",
        "Special Features1": "AL",
        "Special Features2": "AM",
        "Special Features3": "AN",
        "Special Features4": "AO",
        "Special Features5": "AP",
        "Weave Type": "BL",
        "产品数量": "AW",
        "件数": "BB",
        "内包装数量": "DP",
        "功能": "ES",
        "功能状况": "EX",
        "包装宽度": "GN",
        "包装宽度单位": "GO",
        "包装级别": "M",
        "包装重量": "GR",
        "包装重量单位": "GS",
        "包装长度": "GL",
        "包装长度单位": "GM",
        "包装高度": "GP",
        "包装高度单位": "GQ",
        "原产国": "HB",
        # 旧 Header 作为回退；新 Header 放在后面，有非空值时优先。
        "商品变体主题": "F",
        "变体主题名称": "F",
        "更新删除": "C",
        "商品信息操作": "C",
        "商品厚度十进制数值": "DC",
        "商品厚度单位": "DE",
        "商品宽度单位": "DL",
        "商品是否抗污": "CQ",
        "商品状况": "EB",
        "商品短边的宽度": "DK",
        "商品长度单位": "DJ",
        "商品长边的长度": "DI",
        "地毯款式类型": "DM",
        "型号": "P",
        "处理时间 (US)": "FC",
        "您的价格 USD (在亚马逊上出售, US)": "FF",
        "数量 (US)": "FB",
        "最大订单数量": "EG",
        "父条目的等级": "D",
        "物流渠道代码 (US)": "FA",
        "箱子数量": "GT",
        "适合室内外使用": "DF",
        "配送模板 (US)": "GE",
        "零件编号": "BC",
        "高流量区域使用": "DN",
        "商品尺寸": "AZ",
        "副标题": "H",
        "颜色": "AY",
        "商品特性": "AF",
        "商品特性.1": "AG",
        "商品特性.2": "AH",
        "商品特性.3": "AI",
        "商品特性.4": "AJ",
        "主图片 URL": "U",
        "其他图片 URL1": "V",
        "其他图片 URL2": "W",
        "其他图片 URL3": "X",
        "其他图片 URL4": "Y",
        "其他图片 URL5": "Z",
        "其他图片 URL6": "AA",
        "其他图片 URL7": "AB",
        "其他图片 URL8": "AC",
        "样本图片 URL": "AD",
    }

    # 不同店铺下载的亚马逊模板可能使用不同语言或字段顺序。
    # 默认映射对应“亚马逊--灿东（子账号）”；以下覆盖项按各店铺模板
    # 第 5 行 Amazon 属性 ID 重新核对，而不是依赖显示名称或旧列序。
    store_mapping_overrides = {
        "亚马逊--冬豚--北美（子账号）": {
            "pattern": "CZ",
            "不含税价目表": "FP",
            "Back Material Type": "DW",
            "Care Instructions1": "BT",
            "Care Instructions2": "BU",
            "Construction Type": "BC",
            "Item Shape": "BM",
            "Material Features1": "CB",
            "Material Features2": "CC",
            "Material Features3": "CD",
            "Pile Height": "CG",
            "Recommended Uses For Product1": "DL",
            "Recommended Uses For Product2": "DM",
            "Recommended Uses For Product3": "DN",
            "Recommended Uses For Product4": "DO",
            "Recommended Uses For Product5": "DP",
            "Room Type1": "EI",
            "Room Type2": "EJ",
            "Room Type3": "EK",
            "Room Type4": "EL",
            "Room Type5": "EM",
            "Size": "BB",
            "Weave Type": "BS",
            "件数": "BD",
            "内包装数量": "FF",
            "功能": "GE",
            "功能状况": "GJ",
            "包装宽度": "HZ",
            "包装宽度单位": "IA",
            "包装重量": "ID",
            "包装重量单位": "IE",
            "包装长度": "HX",
            "包装长度单位": "HY",
            "包装高度": "IB",
            "包装高度单位": "IC",
            "原产国": "IP",
            "商品厚度十进制数值": "ES",
            "商品厚度单位": "EU",
            "商品宽度单位": "FB",
            "商品是否抗污": "DV",
            "商品状况": "FN",
            "商品短边的宽度": "FA",
            "商品长度单位": "EZ",
            "商品长边的长度": "EY",
            "地毯款式类型": "FC",
            "处理时间 (US)": "GO",
            "您的价格 USD (在亚马逊上出售, US)": "GR",
            "数量 (US)": "GN",
            "最大订单数量": "FS",
            "物流渠道代码 (US)": "GM",
            "箱子数量": "IH",
            "适合室内外使用": "EV",
            "配送模板 (US)": "HQ",
            "零件编号": "BJ",
            "高流量区域使用": "FD",
            "商品尺寸": "BB",
            "颜色": "AZ",
        },
        "亚马逊--岚风（子账号）": {
            # 根据印花地毯-模板文件v2_亚马逊--岚风（子账号）.xlsm 第5行 Amazon属性ID 重新核对
            "pattern": "BZ",
            "不含税价目表": "ED",
            "Back Material Type": "CR",
            "Care Instructions1": "BM",
            "Care Instructions2": "BN",
            "Construction Type": "BA",
            "Item Shape": "BF",
            "Material": "AR",
            "Material Features1": "BS",
            "Material Features2": "BT",
            "Material Features3": "BU",
            "Pile Height": "BX",
            "Recommended Uses For Product1": "CL",
            "Recommended Uses For Product2": "CM",
            "Recommended Uses For Product3": "CN",
            "Recommended Uses For Product4": "CO",
            "Recommended Uses For Product5": "CP",
            "Room Type1": "CX",
            "Room Type2": "CY",
            "Room Type3": "CZ",
            "Room Type4": "DA",
            "Room Type5": "DB",
            "Size": "AZ",
            "Special Features1": "AL",
            "Special Features2": "AM",
            "Special Features3": "AN",
            "Special Features4": "AO",
            "Special Features5": "AP",
            "Weave Type": "BL",
            "件数": "BB",
            "内包装数量": "DP",
            "功能": "ES",
            "功能状况": "EX",
            "包装宽度": "GN",
            "包装宽度单位": "GO",
            "包装长度": "GL",
            "包装长度单位": "GM",
            "包装高度": "GP",
            "包装高度单位": "GQ",
            "原产国": "HB",
            "商品厚度十进制数值": "DC",
            "商品厚度单位": "DE",
            "商品宽度单位": "DL",
            "商品短边的宽度": "DK",
            "商品长边的长度": "DJ",
            "商品长度单位": "DI",
            "地毯款式类型": "DM",
            "型号": "P",
            "处理时间 (US)": "FC",
            "您的价格 USD (在亚马逊上出售, US)": "FF",
            "数量 (US)": "FB",
            "最大订单数量": "EG",
            "物流渠道代码 (US)": "FA",
            "箱子数量": "GT",
            "适合室内外使用": "DF",
            "配送模板 (US)": "GE",
            "零件编号": "BC",
            "高流量区域使用": "DN",
            "商品尺寸": "AZ",
            "颜色": "AY",
            "商品特性": "AF",
            "商品特性.1": "AG",
            "商品特性.2": "AH",
            "商品特性.3": "AI",
            "商品特性.4": "AJ",
            "关于此艺术品": "AE",
            "Product Description": "AE",
            "搜索关键词": "AK",
            "产品类型关键字": "L",
            "制造商": "R",
            "品牌": "I",
            "商品名称": "G",
            "商品编号类型": "J",
            "商品 ID": "K",
            "卖家 SKU": "A",
            "包装级别": "M",
            "副标题": "H",
            "主图片 URL": "U",
            "其他图片 URL1": "V",
            "其他图片 URL2": "W",
            "其他图片 URL3": "X",
            "其他图片 URL4": "Y",
            "其他图片 URL5": "Z",
            "其他图片 URL6": "AA",
            "其他图片 URL7": "AB",
            "其他图片 URL8": "AC",
            "样本图片 URL": "AD",
            "父条目的等级": "D",
            "父条目的库存单位": "E",
            "商品变体主题": "F",
            "变体主题名称": "F",
            "更新删除": "C",
            "商品信息操作": "C",
            "包装重量": "GR",
            "包装重量单位": "GS",
        },
    }

    # 当前北蓉 v2 与岚风 v2 模板的 101 个 Amazon 属性 ID 列序完全一致。
    # 合并默认映射得到完整 mapping，再分别复制，避免后续修改某个店铺时相互影响。
    lanfeng_mapping = {
        **feishu_to_template_mapping,
        **store_mapping_overrides["亚马逊--岚风（子账号）"],
        "商品长度单位": "DJ",
        "商品长边的长度": "DI",
    }
    store_mapping_overrides["亚马逊--岚风（子账号）"] = lanfeng_mapping.copy()
    store_mapping_overrides["亚马逊-北蓉-北美（子账号）"] = lanfeng_mapping.copy()

    mapping_overrides = store_mapping_overrides.get(store_name, {})
    feishu_to_template_mapping.update(mapping_overrides)
    print(
        f"店铺 mapping：{store_name}，"
        f"按当前模板调整 {len(mapping_overrides)} 个字段"
    )

    # 步骤4: 按 Header 名称构建列映射。
    # col_mapping 的顺序取自上面的映射字典，不受源文件列顺序影响。
    col_mapping = []
    for source_header, target_letter in feishu_to_template_mapping.items():
        if source_header not in header_to_idx:
            print(f"源数据缺少（跳过）：{source_header}")
            continue
        data_idx = header_to_idx[source_header]
        target_col_num = col_letter_to_num(target_letter)
        col_mapping.append((source_header, data_idx, target_letter, target_col_num))
        print(f"映射：{source_header} → 列{target_letter}")

    mapped_headers = {source_header for source_header, _, _, _ in col_mapping}
    for data_col_name in data_columns:
        normalized_name = str(data_col_name).strip()
        if normalized_name not in mapped_headers:
            print(f"未映射：{data_col_name}")

    if not col_mapping:
        raise ValueError("没有任何列可以映射，请检查列名！")

    # 步骤5: 保存并临时删除数据验证（只针对列表类型）
    target_columns = {
        target_letter: target_col_num
        for _, _, target_letter, target_col_num in col_mapping
    }
    validation_backup = {}  # 用于保存原来的验证设置（除IgnoreBlank外）

    print("\n正在备份并临时移除目标列的列表型数据验证...")

    for col_letter, col_num in sorted(target_columns.items(), key=lambda item: item[1]):
        data_range = sheet.range(f"{col_letter}{start_row}:{col_letter}10000")

        try:
            validation = data_range.api.Validation
            val_type = validation.Type

            # 只处理下拉列表类型 (xlValidateList = 3)
            if val_type == 3:
                print(f"  -> 备份并移除列 {col_letter} (第 {col_num} 列) 的下拉列表验证")

                validation_backup[col_num] = {
                    'formula1': validation.Formula1,
                    'alertstyle': validation.AlertStyle,
                    'errortitle': validation.ErrorTitle or "",
                    'errormessage': validation.ErrorMessage or "",
                    'showerror': validation.ShowError
                }

                validation.Delete()

        except:
            pass  # 没有验证或不是列表类型 -> 跳过

    # 步骤6: 写入数据（批量写入，速度快）
    print("\n开始写入数据...")

    # 构建写入矩阵
    rows_count = len(data_values)
    max_col = max(target_columns.values())
    last_target_letter = max(target_columns, key=target_columns.get)
    last_existing_row = max(sheet.used_range.last_cell.row, start_row)
    sheet.range(
        f"A{start_row}:{last_target_letter}{last_existing_row}"
    ).clear_contents()
    write_matrix = [[None] * max_col for _ in range(rows_count)]

    for row_idx, row_data in enumerate(data_values):
        print("row_idx =", row_idx)
        for _, data_col_idx, _, target_col_num in col_mapping:
            value = row_data[data_col_idx]
            final_value = None if (pd.isna(value) or value == "") else value
            # 同一模板列可能有新旧两个来源。空值不能覆盖前面已经写入的
            # 有效回退值；后面的非空新字段仍可覆盖旧字段。
            if final_value is not None:
                write_matrix[row_idx][target_col_num - 1] = final_value

    # FA 写死值：Fulfillment by Merchant (Default)
    fa_col_num = col_letter_to_num("FA")  # 141
    if fa_col_num > max_col:
        # 扩展矩阵列以容纳 FA
        for row in write_matrix:
            row.extend([None] * (fa_col_num - max_col))
        max_col = fa_col_num
    fa_col_idx = fa_col_num - 1  # 0-based index
    for row_idx in range(rows_count):
        write_matrix[row_idx][fa_col_idx] = "Fulfillment by Merchant (Default)"
    print(f"FA 列已写死为：Fulfillment by Merchant (Default)")

    # 一次性写入整块区域
    write_range = sheet.range(f"A{start_row}").resize(rows_count, max_col)
    write_range.value = write_matrix
    print(f"数据写入完成，共 {rows_count} 行")

    # 步骤7: 恢复列表型数据验证，并**强制忽略空值**
    # 注意：此步骤可能影响数据，暂时跳过以确保数据正确写入
    # 如果需要恢复验证，可以后续手动处理
    print("\n跳过数据验证恢复步骤以确保数据正确写入...")
    print(f"已备份 {len(validation_backup)} 个列的验证设置")

    # 步骤8: 保存并退出
    print(f"\n正在保存到：{output_path}")
    wb.save(str(output_path))
    print(f"保存到 output_path 成功")

    # 使用shutil复制文件到 output_skc_path，避免第二次save丢数据
    print(f"\n正在复制到：{output_skc_path}")
    shutil.copy(output_path, output_skc_path)
    print(f"复制到 output_skc_path 成功")

    wb.close()
    _active_wb = None
    app.quit()
    _active_app = None
    Path(output_path).unlink(missing_ok=True)
    _active_work_path = None

    print(f"\n已保存到：\n{output_path}")
    print(f"\n已保存到：\n{output_skc_path}")
    print("全部完成！")
    return output_skc_path


def m():
    """批量处理全部“尺寸图片已匹配”Excel，已生成结果自动跳过。"""
    if not input_dir.is_dir():
        raise FileNotFoundError(f"输入目录不存在：{input_dir}")

    input_files = sorted(
        path for path in input_dir.glob(input_pattern)
        if path.is_file() and not path.name.startswith('~$')
    )
    if not input_files:
        print(f"没有匹配到输入文件：{input_dir / input_pattern}")
        return []

    generated_paths = []
    skipped_count = 0
    failed_files = []
    for file_index, data_path in enumerate(input_files, start=1):
        print(
            f"\n========== 处理文件 [{file_index}/{len(input_files)}]："
            f"{data_path.name} =========="
        )
        try:
            generated_path = process_file(data_path)
        except Exception as exc:
            cleanup_active_excel()
            failed_files.append((data_path.name, str(exc)))
            print(f"处理失败，继续下一文件：{data_path.name}，原因：{exc}")
        else:
            if generated_path is None:
                skipped_count += 1
            else:
                generated_paths.append(Path(generated_path))

    if generated_paths:
        Path(r"D:\NAS_download\nas_path_out_path.txt").write_text(
            ",".join(str(path) for path in generated_paths),
            encoding='utf-8',
        )

    print(
        f"\n批量处理完成：新生成 {len(generated_paths)} 个，"
        f"已生成跳过 {skipped_count} 个，失败 {len(failed_files)} 个"
    )
    if failed_files:
        failure_message = "；".join(
            f"{file_name}: {reason}" for file_name, reason in failed_files
        )
        raise RuntimeError(f"以下文件处理失败：{failure_message}")
    return generated_paths


if __name__ == '__main__':
    m()
def main(args):
    m()
