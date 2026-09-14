# ThesisCraft 原生 Word 与 WPS 入口

作者 Study-Tang，GPT 辅助开发。4.0.4 引入不依赖 CLR 的 Rust COM DLL。

## 使用

1. 解压 Windows 便携包到固定文件夹。
2. 双击 `Install-Native-Addin.cmd`，看到安装完成提示。
3. 保存工作并正常重启 Word/WPS，点击 **学研排版 → 打开排版面板**。
4. 核对面板显示的文档名，然后快速排版或打开论文设置。正常排版只生成新 DOCX。

“模板与图片设置”打开工作台，可调整学校模板、字体、编号、图片尺寸和间距。“关于”显示作者信息。

安装只写当前用户的此项目注册项，提供 32/64 位 DLL，不要求管理员、不启用宏、不改变全局信任策略。卸载入口使用 `Uninstall-Native-Addin.cmd`。移动便携包后重新安装入口即可更新路径。旧 DLL 保留到不再使用时可随旧安装目录处理，更新不会强行关闭 Office。

原生 DLL 只提供功能区和启动面板，排版工作在独立进程完成。原件不覆盖。可继续独立使用 `Word-Plugin.cmd`、`WPS-Plugin.cmd`、桌面版和 CLI。已安装原生入口时通常无需再启用登录后自动连接。

## 验证范围

本机 Windows Word 16.0.20326.20132（64 位）、WPS 12.1.0.28043（32 位）正常启动时，`StudyTang.ThesisCraft` 的 `Connect` 状态自动为 true，均显示功能区，实际点击按钮可打开正确宿主的面板。未通过修改 `Connect` 状态来模拟自动加载。其他版本仍需验证。

Word/WPS 以 COM 自动化模式启动时可能跳过第三方加载项；`DispatchEx` 或手动设 `Connect=true` 不足以验收普通启动。旧 C#/.NET 实验 DLL 在本机普通启动时触发 CLR 崩溃，已停止推荐。

WPS 兼容发现机制同时使用 Word 的发现键，并为本项目写入单项 `AddinsWL` 注册值，因此安装会同时提供 Word/WPS 入口。不会修改其他插件的允许或禁用状态。

## 源码构建

需要 Rust MSVC 工具链、Windows SDK 和 MSVC 链接器：

```powershell
rustup target add i686-pc-windows-msvc
./addins/native/build.ps1
python -c "from word_formatter.native_install import install; install()"
```

`build.ps1` 输出两种架构到 `bin`。也提供供开发调用的 `install.ps1` 与 `uninstall.ps1`；便携包的一键安装由程序自身完成，不依赖 PowerShell 脚本执行策略。安装后需要正常重启宿主。

COM 类标识为 `{F93F581A-74C6-4728-9406-392B6873CA1C}`，ProgID 为 `StudyTang.ThesisCraft`。实现 IUnknown、IDispatch、IDTExtensibility2、IRibbonExtensibility 及类工厂，使用原生 Win32 API，不引入网络服务或驻留后台等待线程。
