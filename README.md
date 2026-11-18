# FACA 品质异常追踪平台

把 `FACA Tracker` Excel 搬上来的可运行平台，包含 **大屏展示** 与 **数据管理（增删改查 + 批量导入导出）** 两个模块。数据落在本地 SQLite，字段按源表第一行建立，支持**富文本 + 图片**、**布尔状态**、**日期类型**。

## 快速开始

```bash
cd faca-platform
npm start          # 仅本机访问
npm run lan        # 同一局域网的其他电脑/手机也能访问
```

- 数据管理：<http://127.0.0.1:8787/>
- 大屏展示：<http://127.0.0.1:8787/screen>

### 同一网络下的其他电脑怎么访问

`npm start` 只监听 `127.0.0.1`，外部连不进来；用 `npm run lan` 启动（等价 `HOST=0.0.0.0 npm start`），启动日志会直接打印局域网地址：

```
FACA 品质异常追踪平台已启动
  数据管理   http://172.23.137.22:8787
  大屏展示   http://172.23.137.22:8787/screen
  本机       http://127.0.0.1:8787
```

查本机 IP：`ipconfig getifaddr en0`（或 系统设置 → 网络）。
打不开时依次排查：两台机器在**同一网段**、macOS 防火墙是否放行 node、公司网络是否开了 AP 隔离、Mac 是否休眠。

首次启动或需要重新灌入源表：

```bash
npm run seed -- "/Users/lienhui/Desktop/FACA Tracker_1012.xlsx"
```

> 运行环境：Node.js ≥ 22（内置 `node:sqlite`，无需 npm 依赖）；Python 3 + `openpyxl`（Excel 读写）、`Pillow`（导出时把图片嵌入 Excel，缺失则只导出文字）。

## 功能

### 一、大屏展示（/screen）

- 1920×1080 设计稿，按窗口等比缩放，任意分辨率自动适配
- 6 个指标：记录总数、进行中、已关闭、逾期未上传、平均 FR、已汇报率
- 6 个面板：状态分布、发生趋势（含进行中虚线）、供应商 TOP、责任人处理量、最新异常、失效率 TOP 排行
- 底部逾期预警滚动条，30 秒自动刷新；图表为自绘 SVG/CSS，**不依赖 CDN，离线可用**

### 二、数据管理（/）

- **查询**：全字段搜索 + 阶段 / 状态 / 供应商 / 类别 / 汇报状态 / 日期区间筛选，服务端分页
- **排序**：发生日期、阶段、供应商、状态、FR 点击表头排序
- **新增 / 编辑 / 删除**：单条、批量删除均有二次确认
- **批量导入 / 批量导出**：见下文

### 三、富文本与图片（NG Details / FA / CA / Remark）

- 这四个字段存的是**富文本 HTML**，不是纯文本，可同时保存文字与图片
- 编辑器工具栏：加粗、斜体、下划线、项目符号、编号、清除格式、插入图片
- 插图三种方式：**Ctrl+V 粘贴截图**、**拖拽图片进编辑区**、点工具栏选文件
- **图片可拖动改大小**：点一下图片出现选框和四角手柄，拖任意一个角即可等比缩放，拖动时实时显示像素尺寸；双击图片看原图
- **多图并排**：图片是行内元素，同一段里放两张会自动并排；不在同一段时，选中图片点选框右上角的「左」/「右」让它环绕，后面的图片和文字就会排到旁边；点「整」恢复独占一行
- 环绕/宽度都会随 HTML 一起保存，查看页按同样版式渲染
- 调整后的宽度以 `width:Npx` 存进 HTML，保存后查看页同样按这个尺寸渲染（上限为编辑器/抽屉宽度，不会溢出）
- 查看时点击图片可放大到原图；表格里显示「图片数量」徽标
- 图片落盘在 `data/uploads/年月/`，按内容哈希命名（重复图片自动复用），通过 `/uploads/...` 访问
- **透明背景的 PNG 会合成到白底**（导出、缩略图、灯箱预览都做了处理），否则透明区会被当成黑色
- 编辑保存不会丢图：编辑器把 HTML（含 `<img>`）整体回传，服务端只做安全清洗，不动图片

**导入时的图文还原规则**（`lib/importer.js` + `tools/xlsx_io.py`）：

1. Excel 里的图片是浮动图形，没有"在文字里的位置"。平台用锚点在**行内的纵向比例**（`rowOff / 行高`）推算它大致在哪一行文字之后，再插回去；超出文字区的图放到末尾。
2. 图片尺寸按 **Excel 里的实际显示尺寸**还原（由锚点/组合变换换算出 px），上限 520px、下限 100px —— 既不会把 100px 的缩略图放大成巨图，也不会把大图压扁。
3. **组合形状（grpSp）会拆开处理**：一组里往往有多张图片＋箭头＋文字标注（例如「改善前 → 改善后」对比图）。现在组内每张图都会单独取出并排显示，标注文字渲染成图注，并与图片行等宽对齐。
4. 图片插入处会吃掉紧随其后的空行——这些空行本来就是 Excel 给浮动图片让位用的，不去掉会出现双倍留白。
5. Excel 里的连续空格转成 `&nbsp;`，保持 `Verification OK @ pic 5;        Completion time 2026/10/9` 这种右侧对齐不被压缩。

### 四、字段类型

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| Status 状态 | **布尔** | `1 = done 已完成`，`0 = ongoing 进行中` |
| Reported 是否已汇报 | **布尔** | `1 = Y 已汇报`，`0 = N 未汇报`，`NULL = 未填写`（源表 26 条 `NA`） |
| Issue Date / FACA Report / Foxtrot Upload | **日期** | 统一存 `YYYY-MM-DD`；`/`、`-` 等占位符视为空 |
| NG Details / FA / CA / Remark | **富文本** | HTML，可含图片 |
| FR | 数值 | 0–1 之间的小数，页面按百分比展示 |
| 其余 | 文本 | |

> 编辑表单里 Status 是「进行中 / 已完成」两段开关，Reported 是「已汇报 / 未汇报 / 未填写」三段开关——保持布尔语义，同时不把源表的 `NA` 强行变成 `N`。

### 五、批量导入规则

1. 支持 `.xlsx / .xlsm / .csv`，一个文件可含多个 Sheet。
2. 每个 Sheet 在前 12 行内自动定位表头行。
3. **按第一行列名判断是否入库**：必需列（`Build/阶段`、`Issue Date/发生日期`、`IPN/ICT料号`、`Vendor/供应商`、`NG Details/不良项`、`Status/状态`）全部命中且总命中 ≥ 10 列 → 导入；否则跳过并给出缺失列原因。
4. 表头支持中英混排（`Issue Date 发生日期`）、纯英文、纯中文、列顺序任意；FA/CA 后缀描述不同也能识别。
5. **单元格内嵌图片会一并导入**：按图片锚点定位到行/列，落入对应的 NG Details / FA / CA / Remark 字段；HEIC 自动转 JPG（macOS `sips`）。
6. 源表多出的列不丢，存 `extra` 字段，导出时追加在末尾。
7. 去重：`阶段 + 料号 + 发生日期 + 供应商 + 不良项` 相同视为重复（按纯文本比对），默认跳过。
8. 每次导入写日志，页面「导入记录」可查。

以桌面这份表为例：`OPT` / `OVB` 各 20/20 列匹配、41 行、30 张图片导入成功；`Pending Review` 因缺 `NG Details` 列被跳过。

### 六、批量导出

- 三种方式：当前筛选结果 / 仅选中记录 / 按 Build 阶段分 Sheet
- 表头还原成源表样式（`EN\nCN`）：**浅灰底 `#BFBFBF` + 黑色粗体**（和源表一致），日期写为真日期（`yyyy-mm-dd`），FR 为百分比格式，冻结首列 + 自动筛选
- **富文本单元格导出为「文字在上、图片在下」**：文字仍是真正的单元格内容（可搜索、可筛选、可再导入），图片按 `oneCellAnchor` 精确贴在文字块下方、所在字段的列内
- **图片尺寸与系统显示一致**：优先取富文本里记录的 `width:Npx`（即页面上的显示宽度），高度按原图比例；列宽会自动撑开容纳最宽的图片，不会被二次压缩
- 同一单元格内按实际宽度从左到右排列，放不下自动换行；行高按「文字高度 + 图片高度」设置，不会压到下一行
- **图注（如「改善前 / 改善后」）会画在图片正下方**：Excel 里图片是浮动对象、文字没法跟着图走，所以导出时用 PIL 把图注渲染成图片底部的一条文字带（中文字体自动查找），正文里不再重复出现这两行
- 行高 = **Chrome 实测的文字高度**（按列宽 + 10pt 雅黑 + 20.5px 行距，避免 Excel 按单词换行后图片压字）+ 图片高度；没有 Chrome 时退回估算公式。图片统一放在**整段文字之后**——Excel 单元格做不到图文流式混排，这是保真的折中
- 受 Excel 单行 409pt 上限约束：极少数「超长文字 + 多张大图」的行会等比缩小图片（41 条样例里 30 张图有 27 张保持原尺寸）
- 已知限制：Excel 的图片是浮动对象（`oneCellAnchor`），**在 Excel 里排序/筛选后图片不会跟着行走**；需要图片随行变化时，用平台内的表格查看，或导出后不要再排序
- `exporter.js` 另保留 `visualCells: true` 模式：整格渲染成图片贴回单元格，版式 100% 一致但文字不可搜索（默认不启用）

## 数据模型

表 `faca_records`：

| 列名 | 含义 | 类型 |
| --- | --- | --- |
| no | NO. 序号 | TEXT |
| stage | Build 阶段 | TEXT |
| where_from | Where 来源 | TEXT |
| issue_date | Issue Date 发生日期 | DATE |
| faca_report_date | FACA Report 提交报告日期 | DATE |
| foxtrot_upload_date | Foxtrot Upload 需上传系统日期 | DATE |
| ipn | IPN ICT料号 | TEXT |
| vendor | Vendor 供应商 | TEXT |
| config | Config 配置 | TEXT |
| component | Component DRP物料名称 | TEXT |
| cos_fun_dim | Cos/Fun/Dim 类别 | TEXT |
| ng_details | NG Details 不良项（富文本） | TEXT(HTML) |
| fr | FR = NG数量/总数量 | REAL |
| status | Status 状态 | INTEGER 布尔 |
| dri | DRI 责任人 | TEXT |
| fa | FA 失效分析（富文本） | TEXT(HTML) |
| ca | CA 纠正措施（富文本） | TEXT(HTML) |
| foxtrot_id | Foxtrot ID | TEXT |
| reported | Reported 是否已汇报 | INTEGER 布尔 |
| remark | Remark 备注（富文本） | TEXT(HTML) |

附加列：`extra`(JSON) / `source_sheet` / `fingerprint`（唯一索引去重）/ `created_at` / `updated_at`。
导入日志表：`import_logs`。数据库文件 `data/faca.db`，图片目录 `data/uploads/`。

**版本迁移**：启动时读取 `PRAGMA user_version`，v1 旧库会自动迁移到 v2（status/reported 文本转布尔、长文本转富文本），迁移过程有事务保护。

## 接口

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET | `/api/schema` | 字段字典（类型、表头别名） |
| GET | `/api/meta` | 筛选下拉、日期范围、布尔计数 |
| GET | `/api/stats` | 大屏聚合数据 |
| GET | `/api/records` | 分页查询 |
| POST | `/api/records` | 新增 |
| GET/PUT/DELETE | `/api/records/:id` | 详情 / 更新 / 删除 |
| POST | `/api/records/batch-delete` | 批量删除 |
| POST | `/api/upload` | 上传图片（富文本插图） |
| POST | `/api/import` | 导入（多 Sheet + 图片） |
| POST | `/api/export` | 导出 xlsx（图文单元格：文字 + 图片在下） |
| GET | `/api/import-logs` | 导入日志 |
| GET | `/uploads/*` | 图片静态访问 |

## 目录结构

```
faca-platform/
├── server.js              HTTP 服务 + REST 路由（零依赖）
├── seed.js                用源表初始化数据库
├── lib/
│   ├── db.js              SQLite 建表、v1→v2 迁移
│   ├── fields.js          20 字段字典 + 表头归一化/匹配算法
│   ├── repo.js            增删改查、富文本清洗、布尔与日期规范化、统计
│   ├── importer.js        多 Sheet 导入、表头校验、图片回填
│   └── exporter.js        导出：文字 + 图片按系统尺寸嵌回单元格
├── tools/xlsx_io.py       openpyxl + zip 解析：提取内嵌图片、生成带样式 xlsx
├── public/
│   ├── index.html         数据管理
│   ├── screen.html        大屏
│   ├── assets/*.css
│   └── js/{common,richtext,app,charts,screen}.js
└── data/
    ├── faca.db
    └── uploads/           富文本图片
```

## 备注

- 服务默认只监听 `127.0.0.1`；局域网/大屏访问：`npm run lan`（或 `HOST=0.0.0.0 PORT=8787 npm start`）
- 端口占用：`PORT=8080 npm start`
- 备份：复制 `data/` 整个目录（含图片）
- 删除记录不会自动清理图片文件（可能被其他记录复用）
