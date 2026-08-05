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

- `module1_v2.py`：读取飞书商品、固定属性和违禁词等数据，按店铺及父体 SKU 分批生成上架源数据；正式模式还会回写飞书子表。
- `set_excel_by_skc_size.py`：读取每个 SKC 的图片数据，根据材质方向和尺寸填写主图、其他图片及样本图片 URL。
- `v2.py`：根据 `FBM上架店铺` 选择对应的 `.xlsm` 模板，通过本机 Excel 生成最终上架文件。
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
   └─ 印花地毯-模板文件v2_<FBM上架店铺>.xlsm

D:\NAS_download\
├─ SKCS.txt
├─ nas_path_out_path.txt
├─ uploaded_files.json
└─ <SKC>\
   └─ <SKC>data.xlsx
```

关键输入要求：

- `D:\NAS_download\SKCS.txt`：每行一个待处理 SKC。
- `output_feishu_table_data_*.xlsx`：图片匹配至少需要 `SKC`、`材质方向`、`size_text` 三列；模板生成还会读取 `FBM上架店铺`、`卖家 SKU` 等业务列。
- `<SKC>data.xlsx`：至少包含 `文件名`、`图片地址`、`尺寸`、`图片类型` 四列。
- 店铺模板名称必须与源数据中的 `FBM上架店铺` 完全一致。

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

### 3. 生成最终 XLSM

```powershell
python v2.py
```

脚本批量读取 `*_尺寸图片已匹配.xlsx`，校验店铺模板和必要字段，通过独立 Excel 实例生成带时间戳的 `.xlsm` 文件，并将本批生成路径写入：

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
- **找不到店铺模板**：确认模板位于 `模板文件` 目录，且文件名中的店铺名称与 `FBM上架店铺` 完全一致。
- **图片未匹配**：确认材质方向为 `印花地毯`，并检查 SKC 图片数据文件的必需列和尺寸写法。
- **Excel 自动化失败**：确认已安装桌面版 Excel，关闭残留的脚本工作簿后重试。
- **飞书请求失败**：检查应用凭证、权限范围、表格/云盘 token、Webhook 和网络连接。
