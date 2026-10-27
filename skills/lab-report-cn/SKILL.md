---
name: lab-report-cn
description: 本校课程实验报告（Word/.docx）撰写标准：沿用课程模板排版，封面信息 + 六行信息表 + "实验过程"图文（每步 10-20 字说明 + 逐步截图）+ 实验结论，并做结构与渲染校验。当用户要求"写实验报告/填实验报告/按上次报告的标准写"时使用。
---

# 本校课程实验报告（实验报告 N.docx）

用于把一次真实的实验过程沉淀成课程实验报告。核心是**三板斧**：①先用模板保住排版；②按"图注在上、截图在下"的图文节奏填"实验过程"；③重开+渲染双校验。

> **隐私红线**：本 skill 是模板，**不得写入任何具体姓名、学号、学院、专业、班级、实验室门牌、账号名**。下文用 `{{占位符}}` 表示，生成报告时**从同目录已有报告（实验一/实验二…）与用户口述中读取真实值**并当场填入；skill 自身永远保持占位符状态。

## 0. 先决条件

- 找到三样东西：**实验指导**（PDF/PPT，看步骤与报告要求）、**课程模板**（如 `实验报告模板\实验报告模版.docx`）、**往期报告**（`实验一.docx`、`实验二.docx`）。
- 往期报告是**排版与文风基准**：封面字段、表格行序、字号、图宽、图注长度都照它来，不要自创版式。
- 往期报告同时是**个人信息来源**：封面/表头的姓名、学号等按往期报告原样沿用，不要凭空编造。
- 用 `office-docx` 技能的 Python 环境（`load_workspace_dependencies` 返回的 python），不要另装环境。

## 1. 报告结构规范（逐项照抄，勿改）

### 1.1 封面（3 段，居中/两端对齐沿用模板）
```
        {{课程名}}       课程课内实验实践报告
学院：     {{学院}}     年级：    {{年级}}       专业：   {{专业}}
班级：   {{班级}}        姓名：  {{姓名}}                  学号：  {{学号}}
```
- 字体 10.5pt（五号）；**字号超过 10.5 会让较长的专业名称折行**，破坏版式。
- 年级写 4 位年份（如 2024），不要写 "24"。
- 字段间用空格对齐（照往期报告的空格数量），不要改成表格或制表符。

### 1.2 信息表（模板固定 6 行，合并单元格位置极易踩坑）
| 行 | 内容 | 写入单元格 |
|---|---|---|
| r0 | 实验实践时间 / 地点 | 时间→`cells[2]`，地点→`cells[4]`（`cells[0]==cells[1]` 是合并的**标签**格，写它会把"实验实践时间"覆盖掉） |
| r1 | 实验实践名称 | `cells[1]`（横向合并），格式"实验N：<实验名称>"，"实验名称"取自指导书标题 |
| r2 | 实验实践目的 | `cells[1]`，逐条 `●` 开头，来源=指导书"实验目的" |
| r3 | 实验实践设备与仪器 | `cells[1]`，列"计算机/宿主软件/平台与版本/镜像"，如"计算机、VMware Workstation、VMware ESXi 8.0平台、Windows 10与CentOS 7安装镜像" |
| r4 | **实验过程**（整行合并） | `cells[0]`，图文混排，见 §2 |
| r5 | 实验实践结论 + 指导教师签名 | `cells[0]`，段首"实验实践结论："，正文首行缩进 21pt，末尾留空行 + `指导教师签名：                        年    月    日` |

写入前先 `clear_cell()`：删掉除首段外的段落、清空首段 run，**保留单元格的段落属性**（否则边框/缩进丢失）。

### 1.3 命名规范
- 虚拟机/用户名用**姓名首字母**（如 张三 → `ZS`），拼成 `{{首字母}}-Windows10`、`{{首字母}}-CentOS7`；账号名同理用小写首字母。
- 指导书要求"左上角能看到 首字母-系统名"时，截图必须包含**控制台标题栏/浏览器标签**，别只截客户机画面。

## 2. "实验过程"图文规范（报告的灵魂）

### 2.1 节奏
`实验过程：` → 每步：**图注段（10–20 个汉字）+ 图片段（居中，宽 15cm）** → 下一步。图注在上、图在下，和往期报告一致。

### 2.2 图注怎么写（10–20 字，动词开头，写"做了什么"）
- 合格：`上传Windows与CentOS安装镜像`(15)、`配置2核CPU、4096MB内存与20GB磁盘`(18)、`安装完成进入Windows10桌面`(15)
- 不合格：`截图1`（无信息）、`点击下一步然后选择自定义安装并新建分区应用确定`（超长）
- 一条图注只对应一个动作/一个界面，不要把 3 个步骤塞进一句。

### 2.3 截图时机（关键：错过就要重来）
| 阶段 | 什么时候按下截图 | 必须包含 |
|---|---|---|
| 环境准备 | 平台/宿主程序启动完成、管理界面登录成功 | 版本号、IP、内存/磁盘等可核对信息 |
| 创建虚拟机 | **每一步向导页面**（类型→名称与系统→存储→自定义硬件→完成后的列表/详情） | 虚拟机名、CPU/内存/磁盘数值、ISO 名称 |
| 装系统 | 每个需要用户决策的界面：语言、现在安装、密钥、版本、许可、安装类型、磁盘/分区、进度、OOBE 各页、**最终桌面/登录提示** | 安装器标题、进度百分比、桌面/登录界面 |
| 收尾 | 汇总页（列表、详情、数据存储文件列表） | 名称与容量对应关系 |

规则：
1. **点完就截**，别等流程走完——对话框一旦关闭无法复现。
2. 进度类画面（安装中、上传中）抓一次带百分比的即可。
3. 截图前清理干扰：关通知弹窗（`NOC_GLOBAL_SETTING_TOASTS_ENABLED=0`）、避免其它窗口遮挡、必要时把目标窗口置顶；**注意别把聊天工具、个人信息窗口截进报告**。
4. 文件命名 `r01_阶段_动作.png`、`r02_...`，方便与图注一一对应，也方便返工补图。

### 2.4 指导书"必需图"核对
开始写报告前，把指导书"实验报告要求"里点名的图片列成清单（常见为：创建好的虚拟机、装好系统的界面，每个系统各一张），逐张确认存在且满足附加条件（如"左上角要看到系统名"）。缺图要在交付时明确告知，不要用示意图/他人截图充数。

## 3. 生成脚本模板（python-docx）

```python
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

TPL = r"<课程目录>\实验报告模板\实验报告模版.docx"   # 永远基于模板改，不新建空白文档
OUT = r"<课程目录>\实验报告\实验{{N}}.docx"

def set_font(run, size=10.5, bold=False, name="宋体"):
    run.font.name = name; run.font.size = Pt(size); run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)   # 中文必须显式设 eastAsia

def clear_cell(cell):
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    for r in list(cell.paragraphs[0].runs):
        r._element.getparent().remove(r._element)
    return cell    # 返回 cell，文本用 cell.paragraphs[0].add_run(...) 加

doc = Document(TPL); tbl = doc.tables[0]
# 封面与 r0-r3 的字段值统一放在 CONFIG 字典里，从往期报告读取后传入，不要写死在脚本注释中
# cell = clear_cell(tbl.rows[4].cells[0]); cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
# cell.paragraphs[0].add_run("实验过程：")
for fname, caption in STEPS:                      # STEPS = [(图片文件名, 10-20字说明), ...]
    p = cell.add_paragraph(); run = p.add_run(caption); set_font(run)
    p.paragraph_format.keep_with_next = True       # 图注不要和图片被分页拆开
    ip = cell.add_paragraph(); ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ip.add_run().add_picture(os.path.join(SHOT, fname), width=Cm(15.0))
doc.save(OUT)
```

要点：
- `_Cell` 没有 `add_run`，只能 `cell.paragraphs[0].add_run(...)`。
- 图片一律 `width=Cm(15.0~15.2)`（表格内可用宽度约 15.5cm），不要按原始像素宽度插。
- 图片会嵌入 docx，源文件可留在采集目录，不必搬迁。
- 个人信息（姓名/学号/学院/专业/班级/地点）只在运行时从 `CONFIG` 填入，**不写进脚本、不写进 skill**。

## 4. 交付前校验（两步都要做）

1. **结构校验**：重开文档，打印 `paragraphs/tables/inline_shapes`、封面三段文字、r0–r5 标签与值、实验过程段落数/图片数、结论段数。核对：图片数 == STEPS 条数、封面未被覆盖、r0 标签仍是"实验实践时间/地点"。
2. **渲染校验**：用 `office-render` 的 LibreOffice Kit `render --pages "1,2,<末页>"`，看图确认：封面不折行、表格边框完整、图注在图片上方且未被拆页、结论与签名栏在最后一页。渲染后删除 `.dsh-office-*` 临时目录。

汇报口径：页数、图片数、缺图清单、以及"真实做过"的边界（哪些是实操、哪些未完成）。

## 5. 本机常见坑（踩过的）

- **Windows PowerShell 5.1 读 UTF-8 无 BOM 的 .ps1 会乱码报解析错误**：脚本里别写中文，或把文件转成 `UTF8Encoding($true)`（带 BOM）。同理 `.ps1` 里不要用 `$args` 当函数参数（自动变量冲突）。
- **中文路径**：写 python 脚本文件（UTF-8）比在命令行传中文参数稳；日志/控制台输出乱码只是显示层问题，不影响文件内容。
- **Word 打开过会留 `~$xxx.docx` 锁文件**：写入前确认 WINWORD 未运行，否则保存失败。
- **模板单元格合并**：`row.cells` 会把同一个合并格重复返回，写错索引就覆盖标签（见 §1.2）。
- **LibreOffice 渲染可能提示缺字体**（如 PMingLiU）：只要页数/版式正常，可忽略；正文统一用宋体。

## 6. 素材采集（若实验也要本 skill 一起做）

- 桌面程序：PowerShell + `user32` P/Invoke 截图（`CopyFromScreen`），先 `SetProcessDPIAware()`，否则坐标被 DPI 缩放；用 `SetWindowPos TOPMOST + SetForegroundWindow` 防遮挡。
- Web 管理台：Edge `--remote-debugging-port=9222 --ignore-certificate-errors --disable-popup-blocking`，用 CDP 驱动：`Page.captureScreenshot` 截图、`Input.dispatchMouseEvent` 真实点击（JS `element.click()` 不算用户手势，`window.open` 会被拦）、`Page.navigate` 换页。
- 虚拟机控制台（WebMKS canvas）：先量 `canvas.getBoundingClientRect()` 与 `canvas.width/height`，把"客户机坐标"换算成页面 CSS 坐标再点击；客户机 GUI 里**点击要按住 200ms 以上**、必要时同点连点两次；键盘用 `rawKeyDown/keyUp` + `windowsVirtualKeyCode`，大写字母要带 `modifiers=8`（Shift）。
- 有 VMware 宿主时，给客户机 `RemoteDisplay.vnc.enabled=TRUE` + 端口，用 VNC 抓画面/发键更稳（免受窗口焦点影响）；ESXi 需在防火墙放行 `vnc` 规则集。
- 资源保护：只用一个客户机同时在跑；截图前确认宿主可用内存，低内存时降低窗口尺寸、关闭多余标签页。