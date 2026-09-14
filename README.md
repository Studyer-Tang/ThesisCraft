# ThesisCraft · 学研排版

**让论文排版更省心。** 面向本科生、硕士生和博士生的本地论文排版工具，也保留公文与普通文档排版。作者与维护者：**Study-Tang**；开发辅助：**GPT**。

[下载 Windows 免安装版](https://github.com/Studyer-Tang/ThesisCraft/releases/latest) · [完整使用说明](docs/THESIS-GUIDE.md) · [反馈问题](https://github.com/Studyer-Tang/ThesisCraft/issues)

## Windows 快速开始

1. 在发布页下载 `ThesisCraft.v4.0.6.Windows-portable.zip`，解压到固定文件夹，无需安装 Python。
2. 双击 `Paper-Studio.cmd` 打开论文工作台，选择论文和模板，先检查，再生成排版副本。
3. 在 Word 或 WPS 中检查生成的文档；需要目录与引用更新、PDF 时，选择对应的 Office 宿主。

| 入口 | 用途 |
| --- | --- |
| `Paper-Studio.cmd` | 论文工作台：模板、字体、编号、检查和导出 |
| `Word-Plugin.cmd` | 打开 Word 学研排版操作窗口 |
| `WPS-Plugin.cmd` | 打开 WPS 学研排版操作窗口 |
| `Install-Native-Addin.cmd` | 安装原生“学研排版”选项卡，之后随 Word/WPS 启动 |
| `Uninstall-Native-Addin.cmd` | 卸载原生选项卡，保留模板与桌面程序 |
| `Desktop.cmd` | 通用桌面程序，支持批量公文排版 |

插件使用前先打开并保存本地文档，再启动对应入口。独立的“学研排版”操作窗口会显示当前连接的文档名，提供快速论文排版、论文设置、插入编号/引用、更新目录与引用、导出 PDF。多个 Office 窗口同时打开时，优先连接最近查看的文档窗口；切换独立实例后，点击“连接当前文档”。即使功能区未显示传统工具栏，也能在操作窗口中使用这些功能。正常排版只生成一个新的 DOCX，不自动保存或打开报告；点击“只检查”才生成检查报告，PDF 需主动选择导出。使用面板排版期间请保持面板运行；安装原生入口后无需另开登录自启程序，可随时从功能区重新打开面板。

**4.0.4 已支持原生 COM 功能区自动加载。** 双击 `Install-Native-Addin.cmd` 一次，正常重启 Word/WPS 后会出现“学研排版”选项卡，点击“打开排版面板”即可使用。仅注册当前用户，无需管理员或开启宏；不是 Office 商店加载项。原生入口采用 Rust DLL，排版核心在独立进程中运行。本机 Word 16.0（64 位）与 WPS 12.1.0.28043（32 位）已验证自动加载及按钮启动；其他版本仍需验证。[原生插件说明](addins/native/README.md)

免注册的外部操作窗口、登录后自动连接、桌面版和命令行继续保留。旧 C#/.NET DLL 已停止推荐，原生 Rust 入口替代其自动加载职责。

## 已有功能

| 范围 | 功能 |
| --- | --- |
| 学校模板 | 北大硕士/博士、本科/硕士/博士通用预设；JSON 模板导入导出、读取 DOCX 样稿中的可测样式 |
| 字体与编号 | 独立设置中西文字体、字号、行距；数字/中文/英文章号，多级标题，附录字母或数字，题注分隔符和公式括号，编号字体 |
| 页面结构 | 前置页与正文分节、罗马/阿拉伯页码、页眉、目录、图表目录、脚注与尾注样式 |
| 图表公式 | SEQ 自动编号、REF 交叉引用、三线表、按行数续表、宽表横页、图片尺寸检查 |
| 文献引用 | RIS/BibTeX/JSON 导入；GB/T 7714-2015、APA、IEEE CSL；保留现有 Zotero/EndNote 域 |
| 论文材料 | 论文骨架、封面/声明模板填充、CSV 符号表、章节拆分与合并 |
| 检查与导出 | 只检查、选择修复、独立输出副本、对象完整性检查、HTML/JSON 报告、Word/WPS PDF 导出与预览 |
| 通用模式 | DOCX/TXT/Markdown、公文标题识别、批处理、粘贴文本、可选表格与标点整理；旧 DOC/WPS 按环境转换 |

处理在本机完成，不需要 AI API。规则识别可在工作台中手动修正。排版生成新副本，原件不覆盖；插件中的“更新目录与引用”和“插入编号/引用”直接作用于当前文档，完成后由用户保存。

## 从源码安装

需要 Python 3.10+；桌面界面还需要 Tk。Windows：

```powershell
git clone https://github.com/Studyer-Tang/ThesisCraft.git
cd ThesisCraft
python -m venv .venv
.venv\Scripts\python.exe -m pip install -e ".[desktop,windows]"
.venv\Scripts\python.exe wfp.py --academic
```

安装后也可用 `thesiscraft-desktop` 打开论文工作台。源码内 `start-paper-studio.cmd`、`start-word-plugin.cmd`、`start-wps-plugin.cmd` 会优先使用本地构建的 EXE，否则使用 `.venv`。

macOS/Linux 可安装 `.[desktop]`，CLI 安装 `.` 即可；Office 工具栏和自动 PDF 导出需要 Windows 桌面 Word/WPS。新论文 GUI 在 macOS/Linux 上尚未验收。

## 命令行

以下命令在安装包所在虚拟环境激活后运行，也可用 `.venv\Scripts\python.exe wfp_cli.py` 代替 `thesiscraft`：

```powershell
# 按北大硕士预设生成论文副本
thesiscraft thesis format -i thesis.docx --template pku-master
# 仅检查
thesiscraft thesis check -i thesis.docx --template pku-master
# 调整字体大小并通过 Word 导出 PDF
thesiscraft thesis format -i thesis.docx --set styles.body.size=12 --host word --pdf
# 导出可编辑模板
thesiscraft thesis template --template doctor -o my-template.json
# 通用模式：单文件或目录批处理
thesiscraft format -i report.docx
thesiscraft format -i ./documents -o ./formatted
```

命令行还提供 `wfp` 和 `paper-studio` 别名；`wfp-desktop` 与 `python wfp.py` 打开通用桌面程序。Windows 配置保存在 `%APPDATA%/Study-Tang/ThesisCraft/`，可在界面中导入自己的配置。

## 模板依据与使用边界

- 北大预设参考公开的 **2014 年研究生书写指南及 2024 年博士模板**，不代表 2026 年院系要求或学校认证。北大本科使用通用预设，再按院系要求调整。[依据与细节](docs/THESIS-GUIDE.md)
- 2014 指南引用 GB/T 7714-2005；内置国标 CSL 为 **2015 版**。请按实际要求选用或导入 CSL。
- 封面、签名和声明需按学校要求提供；软件不会生成真实签名。文献元数据须准确完整，Zotero/EndNote 引用继续由原管理器维护。
- 文本框、复杂浮动对象、特殊 MathType 文档需人工核对；纵向合并表格不自动拆分续表。Markdown 为标记清理，不是完整保真渲染。
- PDF 检查不等于 PDF/A 认证。最终分页应在实际提交使用的 Word/WPS 版本中检查。

## 开发与验证

```powershell
python -m pip install -e ".[desktop,windows,dev]"
python wfp_cli.py test
python -m unittest discover -s tests -v
python -m ruff check word_formatter tests
python packaging/build_skill.py
```

Windows 构建：安装 `requirements-build.txt` 后执行 `python packaging/build_release.py windows`，再执行 `python packaging/build_portable.py`。其他平台须在目标系统上构建。

论文功能已在本机 Word 16.0 与 WPS 12.1.0.28043 验证。详细数据见 [4.0 验证记录](docs/VALIDATION-v4.md)。CI 检查三种系统的 Python 核心及实验插件编译，不代替真实 Office 验收。

[贡献指南](CONTRIBUTING.md) · [架构说明](docs/ARCHITECTURE.md) · [更新记录](CHANGELOG.md)

## 开源许可

项目代码采用 [MIT](LICENSE)。[第三方许可文件](THIRD-PARTY-NOTICES.md)。

### 保存位置与界面

4.0.5 的 Word/WPS 面板、论文工作台和通用桌面版均提供“保存位置”。默认与原文件相同，点击“选择文件夹…”可更改，“恢复默认”回到原稿旁；本次窗口内保留选择，下次独立打开恢复默认。插件进入论文设置会带入所选文件夹。批量文件默认分别存放在各自原稿旁。同名结果自动递增，不覆盖原稿；正常排版仍只生成新 DOCX。

界面使用简洁配色，突出主要排版按钮，模板操作收进“模板管理”，较长设置页可滚动，保存位置与操作区固定在底部。结果生成后可直接打开结果文件夹。
