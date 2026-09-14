> 4.0 已增加论文快捷按钮、编号/引用插入、PDF 导出及可选自动连接，见 [论文工作台说明](../../docs/THESIS-GUIDE.md)。

# Word / WPS 插件

作者：Study-Tang。版本：4.0.2。

## 推荐：已验证的工具栏插件

在项目根目录双击：

- `start-word-plugin.cmd`：启动 Word 工具栏插件。
- `start-wps-plugin.cmd`：启动 WPS 工具栏插件。

存在 `dist/ThesisCraft.v4.0.2.exe` 时使用打包程序；否则使用项目 `.venv` 中的 Python。启动后出现独立的“学研排版”操作窗口，上方显示连接的文档名称；兼容工具栏也会尽可能出现在 Office 的加载项区域。无需到功能区里寻找按钮。

使用方法：

1. 打开目标文档，按 Ctrl+S 保存修改。
2. 启动插件，核对操作窗口显示的文档名。论文点击 **论文设置 / 检查**；普通文档点击 **通用文档排版**。
3. 完成后自动打开新文档。输出放在原文件旁，名字附带 `_formatted`；遇到已有文件自动增加序号。
4. 在论文工作台调整参数并点击 **设为插件默认**，之后可在操作窗口点击 **快速论文排版**。通用参数仍可从 Desktop.cmd 桌面版调整并保存。
5. 使用操作窗口的 **关闭插件** 卸载本次临时工具栏。插件不会关闭文档，也不会修改 Office 的宏安全策略。

这是通过 Windows COM 提供的**外部工具栏插件**：需先启动一次并保持进程运行，暂不随 Office 自动启动。它不是 Office 应用商店 Web Add-in，也不是嵌入文档的宏。启动后不需要网络服务。重复启动同一种宿主会唤起已有操作窗口。

多个独立进程同时打开时，插件优先连接最近查看的可见文档窗口。切换到目标文档，再点击操作窗口的 **连接当前文档**，核对上方文档名后操作。关闭 Word/WPS 后插件保留操作窗口，重新打开文档即可连接；未打开文档时不会额外创建空白 Word/WPS 实例。

已验证环境：本机 Word 16.0 与 WPS `12.1.0.28043`（WPS 的兼容对象模型报告版本 14.0）。测试实际调用 COM 按钮的 Execute，经过点击事件、后台排版和宿主打开结果；确认原文件字节不变、图片保留。其他 Windows/Office/WPS 版本尚需验证。

源码运行：

```powershell
python -m pip install -e ".[desktop,windows]"
python wfp.py --office word
python wfp.py --office wps
```

DOCX 可直接处理。DOC/WPS 旧格式仍依赖安装的 Office 转换能力；插件不接受未保存或仅在线的文档，须先保存到本地。插件内默认关闭“接受修订、自动编号转文本”的 Office 预处理。

## 开发中的标准 COM DLL

`FormatterAddin.cs` 是单独的原生 COM 加载项实现，提供 Ribbon 按钮并调用相同的排版核心。`build.ps1` 已能编译 DLL；直接 COM 激活和调用格式化入口也已经通过。

**本机 Word/WPS 的 COMAddIns 集合尚未枚举出该 DLL，自动加载未验收成功。** 不把它作为当前推荐入口，也不声称它已经可随 Office 自动加载。研究代码、安装/卸载脚本及失败的宿主诊断均保留，供继续排查。测试注册项已移除，不要求用户降低 Office 安全设置。

开发命令：

```powershell
./addins/windows/build.ps1
# 以下是实验性注册方式，当前不保证宿主自动加载：
./addins/windows/install.ps1 -PythonPath "$PWD/.venv/Scripts/python.exe" -Target Both
./addins/windows/uninstall.ps1
```

安装脚本只注册当前用户、此项目专属的 CLSID/ProgID；WPS 的兼容注册也会写 Word 的同名加载项发现键，因此 `-Target WPS` 不是严格隔离的仅 WPS 安装。不要将这套实验注册脚本作为商业发布安装器。

## 开发验证

```powershell
python addins/windows/toolbar_smoke_test.py
```

测试只建立合成文档，关闭自身打开的文档，保留宿主中其他已打开文档。结果写入 `addins/windows/dist/toolbar-smoke-results.json`。运行前请避免同时启动多份相同宿主，以便诊断实例选择。

微软接口参考：

- https://learn.microsoft.com/en-us/office/vba/api/office.commandbarbutton
- https://learn.microsoft.com/en-us/office/vba/api/office.comaddins
- https://learn.microsoft.com/en-us/office/vba/api/office.iribbonextensibility
