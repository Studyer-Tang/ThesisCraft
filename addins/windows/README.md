> 4.0.4 已提供可自动加载的 [原生 Rust 功能区入口](../native/README.md)。本页保留免注册外部工具栏与旧 .NET 实验代码说明。

# Word / WPS 插件

作者：Study-Tang。版本：4.0.4。

## 推荐：已验证的工具栏插件

在项目根目录双击：

- `start-word-plugin.cmd`：启动 Word 工具栏插件。
- `start-wps-plugin.cmd`：启动 WPS 工具栏插件。

存在 `dist/ThesisCraft.v4.0.4.exe` 时使用打包程序；否则使用项目 `.venv` 中的 Python。启动后出现独立的“学研排版”操作窗口，上方显示连接的文档名称；兼容工具栏也会尽可能出现在 Office 的加载项区域。无需到功能区里寻找按钮。

使用方法：

1. 打开目标文档，按 Ctrl+S 保存修改。
2. 启动插件，核对操作窗口显示的文档名。论文点击 **论文设置 / 检查**；普通文档点击 **通用文档排版**。
3. 完成后自动打开新文档。输出放在原文件旁，名字附带 `_formatted`；遇到已有文件自动增加序号。
4. 在论文工作台调整参数并点击 **设为插件默认**，之后可在操作窗口点击 **快速论文排版**。通用参数仍可从 Desktop.cmd 桌面版调整并保存。
5. 使用操作窗口的 **关闭插件** 卸载本次临时工具栏。插件不会关闭文档，也不会修改 Office 的宏安全策略。

这是通过 Windows COM 提供的**外部工具栏插件**：可手动启动或启用登录后自动连接；只需在使用操作窗口期间保持进程运行。它不是 Office 应用商店 Web Add-in，也不是嵌入文档的宏。启动后不需要网络服务。重复启动同一种宿主会唤起已有操作窗口。

多个独立进程同时打开时，插件优先连接最近查看的可见文档窗口。切换到目标文档，再点击操作窗口的 **连接当前文档**，核对上方文档名后操作。关闭 Word/WPS 后插件保留操作窗口，重新打开文档即可连接；未打开文档时不会额外创建空白 Word/WPS 实例。

已验证环境：本机 Word 16.0 与 WPS `12.1.0.28043`（WPS 的兼容对象模型报告版本 14.0）。测试实际调用 COM 按钮的 Execute，经过点击事件、后台排版和宿主打开结果；确认原文件字节不变、图片保留。其他 Windows/Office/WPS 版本尚需验证。

源码运行：

```powershell
python -m pip install -e ".[desktop,windows]"
python wfp.py --office word
python wfp.py --office wps
```

DOCX 可直接处理。DOC/WPS 旧格式仍依赖安装的 Office 转换能力；插件不接受未保存或仅在线的文档，须先保存到本地。插件内默认关闭“接受修订、自动编号转文本”的 Office 预处理。

## 已停止推荐的 .NET 实验 DLL

`FormatterAddin.cs` 保留为实验代码。直接 COM 激活可以成功，但本机正常启动 Word 时该 .NET 实现触发 CLR 崩溃。此前用自动化宿主的 COMAddIns 集合做检查，也不足以验证正常启动。

原生 Rust 入口已经替代其自动加载职责，见 [原生插件说明](../native/README.md)。不要继续运行旧安装脚本；它默认拒绝安装，仅在隔离开发环境研究时才接受 `-Experimental`。旧实验注册可用本目录的 `uninstall.ps1` 移除。

## 开发验证

```powershell
python addins/windows/toolbar_smoke_test.py
```

测试只建立合成文档，关闭自身打开的文档，保留宿主中其他已打开文档。结果写入 `addins/windows/dist/toolbar-smoke-results.json`。运行前请避免同时启动多份相同宿主，以便诊断实例选择。

微软接口参考：

- https://learn.microsoft.com/en-us/office/vba/api/office.commandbarbutton
- https://learn.microsoft.com/en-us/office/vba/api/office.comaddins
- https://learn.microsoft.com/en-us/office/vba/api/office.iribbonextensibility
