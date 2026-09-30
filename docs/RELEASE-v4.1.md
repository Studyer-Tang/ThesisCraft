# ThesisCraft 4.1 桌面预览版

Windows 与 macOS 使用同一个论文工作台：选择 DOCX、选择模板、检查或生成排版副本。高级字体、页码、图表、文献设置保留在“高级设置”中。

- Windows：下载 `.exe`，双击运行，无需安装 Python。
- Apple 芯片 Mac：下载 `macOS-arm64.app.zip`。
- Intel Mac：下载 `macOS-x86_64.app.zip`。
- Mac 解压后把 ThesisCraft.app 拖到“应用程序”，再打开。

文档在本机处理，生成新副本。界面显示检查提示和本次修改，可主动打开结果或所在文件夹。默认模板为硕士通用，学校预设须对照当前院系要求。原有默认模板继续保留。

Windows Word/WPS 自动化仍为可选功能；Mac 生成 DOCX 后，在 Word/WPS 中手动更新目录和导出 PDF。此版独立 EXE 不包含原生功能区安装器，旧插件安装包继续保留在历史发布中。

此预览版没有 Apple Developer ID 公证或 Windows Authenticode 发布者证书，首次下载可能被系统拦截。Mac 可在系统“隐私与安全性”中允许本次打开；不需要关闭系统安全保护。请先用论文副本试用并核对最终分页、公式和引用。

每个平台的构建均须通过测试，以及打包程序自身的 DOCX、文献样式、检查报告和 GUI 启动验证，才能上传。自动验证不等于所有 Word/WPS 版本或真实论文都已完成验收。校验文件见同页 SHA256SUMS。
