# 亚马逊自动上架批量处理脚本

这是一套面向 Windows 的亚马逊商品批量上架数据处理脚本。项目从飞书多维表格读取商品与店铺数据，生成上架 Excel，按 SKC 和尺寸补齐图片 URL，再写入对应店铺的宏模板并将生成的 `.xlsm` 文件上传到飞书云盘。

## 处理流程

```text
飞书多维表格
    ↓ module1_v2.py
output_feishu_table_data_<店铺>_<父体SKU>.xlsx
    ↓ set_excel_by_skc_size.py
*_尺寸图片已匹配.xlsx
    ↓ v2.py
<店铺>_<SKU>_<时间戳>.xlsm
    ↓ upload.py
飞书云盘 + 飞书机器人通知
```

各脚本职责：

- `module1_v2.py`：读取飞书商品、固定属性和违禁词等数据，可按材质方向筛选，再按店铺及父体 SKU 分批生成上架源数据；正式模式还会回写飞书子表。
- `set_excel_by_skc_size.py`：读取每个 SKC 的图片数据，根据材质方向匹配主图、其他图片及样本图片 URL；目前支持印花地毯和仿羊绒厨房垫。
- `v2.py`：根据 `材质方向 + FBM上架店铺` 选择对应的 `.xlsm` 模板，通过本机 Excel 生成最终上架文件。
- `upload.py`：发现新生成的 `.xlsm` 文件，上传到指定飞书云盘节点，发送机器人通知，并记录文件指纹以避免重复上传。

## 运行环境

- Windows
- Python 3.10 或更高版本
- Microsoft Excel（`v2.py` 通过 `xlwings` 操作 Excel）
- 可访问飞书 OpenAPI 的网络环境
- RPA 场景运行时还需要项目所属的 `xbot` 与 `package` 模块

安装普通 Python 依赖：

```powershell
python -m pip install pandas openpyxl xlwings requests
```

> 仓库未包含业务数据、Excel 模板和 RPA 运行时，这些资源需要单独准备。

## 目录和文件约定

当前代码使用以下固定路径；如果本机目录不同，请先修改各脚本顶部的配置常量。

```text
D:\项目文件\AI自动上架\
├─ output_feishu_table_data_*.xlsx
├─ *_尺寸图片已匹配.xlsx
├─ <店铺>_<SKU>_<时间戳>.xlsm
└─ 模板文件\
   └─ <材质方向>-模板文件v2_<FBM上架店铺>.xlsm

D:\NAS_download\
├─ SKCS.txt
├─ nas_path_out_path.txt
├─ uploaded_files.json
└─ <SKC>\
   └─ <SKC>data.xlsx

D:\oss\
└─ <SKC>.xlsx
```

关键输入要求：

- `D:\NAS_download\SKCS.txt`：每行一个待处理 SKC。
- `output_feishu_table_data_*.xlsx`：图片匹配至少需要 `SKC`、`材质方向`、`size_text` 三列；模板生成还会读取 `FBM上架店铺`、`卖家 SKU` 等业务列。
- 图片数据优先读取 `D:\NAS_download\<SKC>\<SKC>data.xlsx`；不存在时回退读取 `D:\oss\<SKC>.xlsx`。
- 印花地毯图片表需要 `文件名`、`图片地址`、`尺寸`、`图片类型` 四列；厨房垫至少需要 `文件名`、`图片地址`，建议同时保留 `图片类型`。
- 模板文件名中的材质方向和店铺名称必须与源数据中的 `材质方向`、`FBM上架店铺` 完全一致。

## 配置

在 `config.py` 中配置飞书应用凭证及两个通知 Webhook：

```python
app_id = "your_app_id"
app_secret = "your_app_secret"
webhook_url_upload = "https://open.feishu.cn/open-apis/bot/v2/hook/..."
webhook_url_v2 = "https://open.feishu.cn/open-apis/bot/v2/hook/..."
```

代码中还包含飞书多维表格 token、表 ID、云盘父节点及 Windows 绝对路径。迁移环境时请搜索并检查这些常量：

```powershell
rg "D:\\|Jolyb8Q|tbl|FEISHU_PARENT_NODE" -g "*.py" .
```

## 使用方法

建议依次执行以下步骤。

### 1. 生成上架源数据

```powershell
python module1_v2.py
```

脚本读取 `SKCS.txt`，遍历配置的店铺，并按父体 SKU 分批生成 `output_feishu_table_data_<店铺>_<父体SKU>.xlsx`。

只处理指定材质方向：

```powershell
python module1_v2.py --material-direction "仿羊绒厨房垫"
```

不传 `--material-direction` 时处理全部材质。模块/RPA 调用也支持 `material_direction` 或 `材质方向` 入参。

`module1_v2.py` 中的 `TEST_MODE` 控制是否回写飞书子表：

- `True`：仅生成本地 Excel。
- `False`：生成 Excel，同时新增或更新飞书子表记录。

首次运行或调试时建议先使用测试模式。

### 2. 匹配图片 URL

批量处理默认目录中的所有源数据：

```powershell
python set_excel_by_skc_size.py
```

处理单个文件：

```powershell
python set_excel_by_skc_size.py "D:\项目文件\AI自动上架\output_feishu_table_data_xxx.xlsx"
```

指定图片数据根目录和输出文件：

```powershell
python set_excel_by_skc_size.py input.xlsx --image-root "D:\NAS_download" --output output.xlsx
```

直接覆盖输入文件：

```powershell
python set_excel_by_skc_size.py input.xlsx --in-place
```

`--in-place` 与 `--output` 不能同时使用；批量模式不能使用 `--output`。

图片数据文件不存在的 SKC 会去重后写入：

```text
D:\项目文件\AI自动上架\缺少图片的SKC.txt
```

可以通过 `--missing-image-output` 指定其他清单路径。

仿羊绒厨房垫使用固定的文件名关键词顺序：

| 目标字段 | 匹配关键词 |
|---|---|
| 主图片 URL | `20X32+20X48厨房垫` |
| 其他图片 URL1 | `封面图2X5厨房`，回退 `2X5厨房` |
| 其他图片 URL2 | `封面图2X3门口`，回退 `2X3门口` |
| 其他图片 URL3 | `2X5走廊` |
| 其他图片 URL4 | `120X170单椅` |
| 其他图片 URL5 | `2X5尺寸白底图`、`白底图2X5` 或 `2X5白底图` |
| 其他图片 URL6–8 | 留空 |
| 样本图片 URL | 优先 `图片类型=样本图片`，再回退文件名 `SWATCH`/`SWITCH` |

文件名匹配忽略大小写、空格，并兼容 `x/X/×` 和半角/全角加号。

### 3. 生成最终 XLSM

```powershell
python v2.py
```

脚本批量读取 `*_尺寸图片已匹配.xlsx`，校验材质方向、店铺模板和必要字段，通过独立 Excel 实例生成带时间戳的 `.xlsm` 文件。模板命名规则为：

```text
<材质方向>-模板文件v2_<FBM上架店铺>.xlsm
```

同一个输入文件只能包含一个非空材质方向和一个店铺。不同材质模板的 Amazon 列序可能不同，因此映射配置分为“店铺覆盖”和“材质 + 店铺覆盖”；当前 `仿羊绒厨房垫 + 亚马逊--冬豚--北美（子账号）` 已按模板第 5 行 Amazon 属性 ID 配置并验证专属映射。

本批生成路径会写入：

```text
D:\NAS_download\nas_path_out_path.txt
```

执行期间请勿手动关闭脚本创建的 Excel 实例。

### 4. 上传飞书云盘

```powershell
python upload.py
```

上传器会合并路径清单与输出目录中的时间戳 `.xlsm` 文件。成功记录保存在 `D:\NAS_download\uploaded_files.json`；只有路径、文件大小或修改时间发生变化时才会重新上传。

## 注意事项

- 脚本会访问并修改真实飞书数据，运行正式模式前请确认应用权限、店铺列表和表格 ID。
- `v2.py` 依赖宏模板及桌面版 Excel，不适合直接在无 GUI 的 Linux 服务器运行。
- 批处理采用“单文件/单批失败后继续”的方式，结束时会汇总失败项并抛出异常，请查看完整控制台日志。
- `config.py`、Webhook、应用密钥和业务 token 都属于敏感信息，不应提交到公开仓库。如果凭证曾经进入 Git 历史，请立即在飞书后台轮换，而不只是删除当前文件。
- 上传历史用于幂等控制；删除 `uploaded_files.json` 会使已生成文件在下次运行时重新上传。

## 故障排查

- **找不到输入文件**：检查固定目录、文件名后缀以及文件是否被 Excel 临时占用。
- **找不到店铺模板**：确认模板位于 `模板文件` 目录，且文件名中的材质方向、店铺名称与源数据完全一致。
- **图片未匹配**：检查 NAS 与 `D:\oss` 回退路径、图片表必需列、SKC、材质方向、尺寸及文件名关键词。样本图建议将 `图片类型` 设置为 `样本图片`。
- **字段写入错列**：不同材质的模板列序可能不同，应按模板第 5 行 Amazon 属性 ID 增加对应的“材质 + 店铺”专属覆盖映射。
- **Excel 自动化失败**：确认已安装桌面版 Excel，关闭残留的脚本工作簿后重试。
- **飞书请求失败**：检查应用凭证、权限范围、表格/云盘 token、Webhook 和网络连接。
