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
<材质方向>_<宽度单位>_<SKU>_<店铺>_<时间戳>.xlsm
    ↓ upload.py
飞书云盘 + 飞书机器人通知
```

各脚本职责：

- `fetch_skcs_by_material.py`：按材质方向读取指定飞书视图的 `SKC` 字段，覆盖生成 `D:\NAS_download\SKCS.txt`，并把生成结果发送到飞书群。
- `module1_v2.py`：读取飞书商品、固定属性和违禁词等数据，可按材质方向筛选，再按店铺及父体 SKU 分批生成上架源数据；正式模式还会回写飞书子表。
- `set_excel_by_skc_size.py`：读取每个 SKC 的图片数据，根据材质方向匹配主图、其他图片及样本图片 URL；目前支持印花地毯、仿羊绒厨房垫和三明治户外垫。
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
├─ <材质方向>_<宽度单位>_<SKU>_<店铺>_<时间戳>.xlsm
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
- 印花地毯图片表需要 `文件名`、`图片地址`、`尺寸`、`图片类型` 四列；厨房垫和三明治户外垫至少需要 `文件名`、`图片地址`，建议同时保留 `图片类型`。
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

### 1. 从飞书视图生成 SKCS.txt

必须通过 `--material-direction` 指定材质方向：

```powershell
python fetch_skcs_by_material.py --material-direction "印花地毯" --check-images --no-group-notification
python fetch_skcs_by_material.py --material-direction "仿羊绒厨房垫" --check-images --no-group-notification
python fetch_skcs_by_material.py --material-direction "三明治户外垫" --check-images --no-group-notification
python fetch_skcs_by_material.py --material-direction "印花地毯圣诞专用" --check-images --no-group-notification
https://wit0jhu6kvu.feishu.cn/wiki/BeTcwtVfJidBvikXhCMcIp0MnU3?table=tbluo8xCboMg3b0T&view=vewJPF7bhh
```

生成清单后立即检查 PostgreSQL 图片是否齐全：

```powershell
python fetch_skcs_by_material.py --material-direction "印花地毯" --check-images
```

这个一键入口会先读取飞书候选 SKC，再按 `public.listing_image_versions` 中该材质唯一启用的 Amazon 版本执行只读图片检查。检查尺寸、图片名称、候选名称、顺序和必需图片全部取数据库，不再使用代码中的固定尺寸表。缺少任一必需图片的 SKC 会从最终 `SKCS.txt` 排除；检查失败不覆盖原有清单。数据库的“其他”是显式通配规则，不作为独立商品尺寸检查；“全部尺寸”用于所有商品尺寸共用图片。

可通过 `--image-version` 指定已启用版本名称：

```powershell
python fetch_skcs_by_material.py --material-direction "印花地毯" --check-images --image-version "批量上架v1" --no-group-notification
python pg_set_excel_by_skc_size.py --image-version "批量上架v1" --no-group-notification
```

省略版本名时按材质自动选择唯一启用版本；缺失、重名、未启用或规则无效都会报错，不回退旧规则。圣诞材质映射数据库的“印花地毯”分类，使用圣诞图片前需在管理站启用相应版本。普通、厨房垫、三明治和圣诞的检查与填图共用 `pg_v2_set_excel_by_skc_size.py` 的数据库规则解析器，支持文件名优先匹配及数据库配置的“尺寸＋图片类型＋第几张”。管理站预览也读取启用版本。6X9 是否共用 8X10，由数据库中的两组配置决定。

检查结果会写入：

```text
D:\NAS_download\数据库无图片的SKC.txt
D:\code\亚马逊自动上架-批量上架脚本\缺图明细.txt
```

默认会把检查汇总和缺图明细发送到飞书群。使用以下参数关闭全部群通知（包含 `SKCS.txt` 更新通知）：

```powershell
python fetch_skcs_by_material.py --material-direction "印花地毯" --check-images --no-group-notification
```

材质与飞书视图的对应关系：

| 材质方向 | 飞书视图 ID |
|---|---|
| 印花地毯 | `vew1r9vIKo` |
| 印花地毯圣诞专用 | `vewJPF7bhh` |
| 仿羊绒厨房垫 | `vewyE7tvJK` |
| 三明治户外垫 | `vew0xHeRAu` |

圣诞材质也按数据库唯一启用的 Amazon / 印花地毯版本检查；可加 `--image-version "2026圣诞专题v1"` 显式指定已启用圣诞版本。按首选与候选名称匹配，兼容有无“封面图”前缀及大小写差异。版本缺失、未启用或规则无效时停止，不覆盖原 SKCS.txt。

脚本分页读取视图中的 `SKC` 字段，清理空行并按出现顺序去重，然后覆盖写入 `D:\NAS_download\SKCS.txt`（UTF-8、每行一个 SKC）。未设置 `--no-group-notification` 时，写入成功后会向配置的飞书机器人群发送材质方向、SKC 数量、保存位置和 SKC 明细；明细过长时自动分段发送。视图没有数据时会生成空文件。

### 2. 生成上架源数据

仅处理一个 SKC（不读取或修改 `SKCS.txt`）：

```powershell
python module1_v2.py --material-direction "印花地毯" --skc "DT5001-V7ORANGE"
```

单 SKC 模式不要求花型记录的 `父体SKU` 非空。已有父体 SKU 时仍使用原值；为空时，用花型记录的 `花型风格`、当前匹配店铺、`运营团队分配`，精确匹配父体表 `tblZOU82tgaLNhAx` 中的 `风格`、`店铺`、`团队`（人员 ID 集合）。只有唯一匹配才继续，并要求父体在五点文案表中有对应记录；无匹配、多匹配或条件缺失会明确报错，不自动回写花型表。其他筛选条件及正式模式的子表写入行为保持不变。

```powershell
python module1_v2.py --material-direction "印花地毯"
```

脚本读取 `SKCS.txt`，遍历配置的店铺，并按父体 SKU 分批生成 `output_feishu_table_data_<店铺>_<父体SKU>.xlsx`。

只处理指定材质方向：

```powershell
python module1_v2.py --material-direction "印花地毯"
python module1_v2.py --material-direction "仿羊绒厨房垫"
python module1_v2.py --material-direction "三明治户外垫"
python module1_v2.py --material-direction "印花地毯圣诞专用"
python module1_v2.py --material-direction "印花地毯" --skc "DT5001-V7ORANGE"

python pg_set_excel_by_skc_size.py --no-feishu-notification
```

`--material-direction` 必须传入且不能为空。模块/RPA 调用也支持 `material_direction` 或 `材质方向` 入参。

圣诞专用模式按父体的“卖家 SKU”匹配[文案库](https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tblYg6tDUuMAKsfT&view=vewoqs4zTj)，子体文案使用 `26年圣诞-` 前缀：

| 输出内容 | 文案库字段 |
| --- | --- |
| 商品名称 | `26年圣诞-商品名称{尺寸}` |
| 副标题 | `26年圣诞-Item Highlight {尺寸}` |
| 五点描述 | `26年圣诞-商品特性1`～`26年圣诞-商品特性5` |
| Product Description | `26年圣诞-Product Description` |
| 搜索关键词 | `26年圣诞-搜索关键词` |
| 设计、pattern | `26年圣诞-设计`、`26年圣诞-Pattern` |

各材质的 `title_size_list` 从对应固定字段表中 `是否取用=是` 的 `size_text` 动态读取：

| 材质方向 | 固定字段表 |
| --- | --- |
| 印花地毯 | [tblNqTrr56XFMxYJ](https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tblNqTrr56XFMxYJ&view=vewBwTjUSL) |
| 印花地毯圣诞专用 | [tblxg4at2pg6lnyU](https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tblxg4at2pg6lnyU&view=vewBwTjUSL) |
| 仿羊绒厨房垫 | [tbl18dyMkn1KS8KH](https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tbl18dyMkn1KS8KH&view=vew294Yp4d) |
| 三明治户外垫 | [tblCffxAeXGnocAY](https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tblCffxAeXGnocAY&view=vewzrq6EcO) |

每个店铺读取一次当前材质的固定字段，标题、副标题、子体新增和 Excel 导出共用这份尺寸数据。新增尺寸或切换“是否取用”后，下次运行自动生效，不再维护代码中的尺寸列表；花型表中的历史尺寸标题也不会额外生成子体。启用记录的 `size_text` 为空或没有任何启用尺寸时明确报错。

尺寸统一为大写并去掉空格，兼容 `×`；厨房垫的 `20X32-20X48` 可匹配 `Item Name 20 x32" +20 x48"` 和对应 `Item Highlight`，新增厨房尺寸同样按尺寸后缀匹配文案。圣诞文案仍取 `26年圣诞-` 字段，缺失时不回退到普通印花地毯文案；缺少任一五点描述会报出字段名、SKC、父体 SKU 和文案记录位置。

父体行文案继续来自父体表；花型查询和最终 Excel 的“材质方向”仍为 `印花地毯`。

圣诞专用模式的子体 SKU 使用 `{SKC}-{尺寸}FT-XMAS26-POD`，例如 `ARB7119-SAFFRON BLOOM-2X3FT-XMAS26-POD`。新增飞书子体和 Excel 导出采用相同规则，Excel 子体“零件编号”也跟随此 SKU。

`module1_v2.py` 中的 `TEST_MODE` 控制是否回写飞书子表：

- `True`：仅生成本地 Excel。
- `False`：生成 Excel，同时新增或更新飞书子表记录。

首次运行或调试时建议先使用测试模式。

### 3. 匹配图片 URL

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

三明治户外垫使用以下固定顺序：

| 目标字段 | 匹配关键词 |
|---|---|
| 主图片 URL | `5X7户外庭院` |
| 其他图片 URL1 | `白底图5X7` |
| 其他图片 URL2 | `3X5户外门口` |
| 其他图片 URL3 | `2.5X8户外走廊` |
| 其他图片 URL4 | `封面图5X7客厅` |
| 其他图片 URL5 | `120X170单椅` |
| 其他图片 URL6 | `5X7卧室` |
| 其他图片 URL7–8 | 留空 |
| 样本图片 URL | 优先 `图片类型=样本图片`，再回退文件名 `SWATCH`/`SWITCH`/`SWICH` |

文件名匹配忽略大小写、空格，并兼容 `x/X/×` 和半角/全角加号。

#### 按数据库中的已启用版本填图

`pg_v2_set_excel_by_skc_size.py` 只使用数据库规则，通过 `--image-version`（别名 `--version-name`）选择已启用图片版本；同名版本分属不同材质时加 `--material-direction`。`pg_set_excel_by_skc_size.py` 同样使用数据库规则，支持不传版本名、按 Excel 各行材质自动选择。例如批量处理默认目录中的待匹配 Excel：

```powershell
python pg_v2_set_excel_by_skc_size.py --image-version "2026圣诞专题v1" --no-group-notification
```

处理单个文件：

```powershell
python pg_v2_set_excel_by_skc_size.py "D:\项目文件\AI自动上架\output_feishu_table_data_店铺_父体SKU.xlsx" --image-version "2026圣诞专题v1" --no-group-notification
```

脚本从 `pg.config` 连接数据库，查询 `public.listing_image_versions` 中 `version_name` 等于传入名称且 `is_enabled = TRUE` 的 Amazon / 印花地毯版本。按 `rules_config.groups` 中的尺寸及各目标列规则，从 `public.ods_fbm_image_upload` 为对应 SKC 匹配图片链接，写入主图、其他图和样本图列。

- 每批只读取一次版本规则。版本未启用、不存在、重名或结构无效时停止，不回退到旧图片顺序。
- Excel 的 `材质方向=印花地毯` 或 `印花地毯圣诞专用` 使用该版本；其他材质行不改图，父体无 SKC 或无尺寸的行不填图。
- 按 `size_text` 匹配尺寸，忽略大小写、空格、乘号差异及 FT/IN 后缀；版本缺少当前文件所需尺寸时，该文件报错且不保存。
- 图片名称和顺序完全来自数据库各目标列的 `primary`、`alternatives` 和 `type`。文件名依次匹配首选及候选名称，兼容有无“封面图”前缀；匹配不到记入缺图明细，配置留空或缺失的图片列会清掉旧链接。
- 已删除脚本内写死的各材质图片顺序、尺寸分组、样本图别名及旧匹配回退。数据库中的 `source` 仅是历史来源信息，不再调用旧规则；拖动后按当前目标列保存的首选/候选文件名匹配。
- 缺图明细继续按 SKC 汇总缺少的图片名称。默认发送群通知，使用 `--no-group-notification` 关闭。
- 默认另存为 `*_尺寸图片已匹配.xlsx`；可用 `--output` 指定单文件输出，或用 `--in-place` 覆盖原文件。

### 4. 生成最终 XLSM

```powershell
python v2.py
```

脚本批量读取 `*_尺寸图片已匹配.xlsx`，校验材质方向、店铺模板和必要字段，通过独立 Excel 实例生成带时间戳的 `.xlsm` 文件。模板命名规则为：

```text
<材质方向>-模板文件v2_<FBM上架店铺>.xlsm
```

最终文件命名规则为：

```text
<材质方向>_<宽度单位>_<首个卖家SKU>_<FBM上架店铺>_<14位时间戳>.xlsm
```

同一个输入文件只能包含一个非空材质方向和一个店铺。不同材质模板的 Amazon 列序可能不同，因此映射配置分为“店铺覆盖”和“材质 + 店铺覆盖”。当前已配置专属映射的组合包括：

- `仿羊绒厨房垫 + 亚马逊--冬豚--北美（子账号）`
- `三明治户外垫 + 亚马逊-北蓉-北美（子账号）`

三明治户外垫模板中的重点字段映射如下：

| 源字段 | 模板列 |
|---|---|
| 型号 | `P` |
| Model Name | `Q` |
| 零件编号 | `BJ` |
| 商品长边的长度 | `EY` |
| 商品长度单位 | `EZ` |
| 商品短边的宽度 | `FA` |
| 商品宽度单位 | `FB` |
| 物流渠道代码 (US, CA) | `GQ`，固定写入 `卖家自配送（默认）` |
| 数量 (US, CA) | `GR` |
| 处理时间 (US, CA) | `GS` |
| 配送模板 | `HU` |

映射应以模板第 5 行 Amazon 属性 ID 为准；模板列顺序调整后，需要同步更新 `v2.py` 中对应的“材质 + 店铺”覆盖配置。

本批生成路径会写入：

```text
D:\NAS_download\nas_path_out_path.txt
```

执行期间请勿手动关闭脚本创建的 Excel 实例。

### 5. 上传飞书云盘

```powershell
python upload.py
```

上传器会合并路径清单与输出目录中的时间戳 `.xlsm` 文件。成功记录保存在 `D:\NAS_download\uploaded_files.json`；只有路径、文件大小或修改时间发生变化时才会重新上传。

## 图片版本管理网站（Flask）

`image_version_manager/image_version_manager_server.py` 提供登录、管理员用户管理、操作日志，以及按 SKC 查询数据库图片和显示缩略图的页面，支持 Amazon、Wayfair、Walmart。

新版通过 `image_version_manager/image_manager_versions.py` 将版本名称、启停和各尺寸图片排列保存到 PostgreSQL 的 `public.listing_image_versions`，所有登录用户共享。修改成功后即时保存，刷新或换电脑仍可读取，并发冲突会提示重做。原图片表只读；`pg_v2_set_excel_by_skc_size.py --image-version` 可按指定的已启用 Amazon / 印花地毯版本生成 Excel 图片链接。

首次需 DBA 执行 `image_version_manager/sql/create_listing_image_versions.sql` 和 `image_version_manager/sql/grant_listing_image_versions.sql`，再同步发布前后端。空库需在页面确认使用默认规则或迁移本机旧配置，不会自动覆盖共享数据。

运行、权限和迁移说明见 [图片管理站部署说明](image_version_manager/IMAGE_MANAGER_DEPLOYMENT.md)。共享持久化已于 2026-09-10 完成正式库建表、授权及线上发布；首次进入空库时仍需在页面确认初始化来源。

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
