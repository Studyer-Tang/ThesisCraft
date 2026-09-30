# 正式签名与稳定版发布

当前 GitHub 托管构建只发布明确的 `-beta.*` 标签，普通版本号不能通过未签名流水线发布。身份审核和费用由账户持有人在 Apple/CA/签名服务官网处理。没有证书时仍可构建预览版，不应删除其未签名声明。

## macOS

加入 Apple Developer Program 后申请 Developer ID Application。将含私钥的证书导入构建机专用钥匙串；用 `xcrun notarytool store-credentials` 保存公证凭据。设置 `APPLE_SIGNING_IDENTITY` 为完整 Developer ID Application 名称，`APPLE_NOTARY_PROFILE` 为钥匙串凭据条目名。

```sh
python packaging/build_release.py macos --signed
```

PyInstaller 对内部二进制及 app 签名并启用 hardened runtime；工具提交 ZIP 公证、对 app stapling、执行 Gatekeeper 验证，再重新归档。DMG 单独签名、公证并验证票据。所有校验和在最终签名、公证后生成。当前应用不申请额外权限，不打包用户凭据；若以后加入 Apple Events 自动化，须另行验证用途说明与最小 entitlement。

## Windows

选择支持所在地区和主体类型的 CA 或申请符合条件的 SignPath Foundation。当前实现使用已安装证书的 Windows 签名提供程序；私钥留在硬件令牌/HSM/云提供程序，不从仓库读取 PFX。设置 `WINDOWS_CERT_SHA1` 为证书指纹，`WINDOWS_TIMESTAMP_URL` 为 CA 提供的 RFC3161 时间戳 URL，`SIGNTOOL` 为 Windows SDK signtool.exe 路径。安装 Inno Setup 6。

```powershell
python packaging/build_release.py windows --signed
```

程序入口、安装包和卸载器均签名，执行 Authenticode 验证。云提供程序须先完成其客户端配置；SignPath/Artifact Signing 的专用 API 集成尚需账户获批后按该服务流程接入，不能仅设置指纹就宣称支持。签名不会保证立刻消除 SmartScreen 提示。

## CI 和稳定版门槛

在已配置硬件/云密钥的受保护构建机执行同一命令。GitHub Actions 可使用受保护 environment、临时钥匙串和加密 secrets 或服务支持的 OIDC；不得把密码放在提交、日志、命令历史中，不向来自 fork 的 PR 提供签名凭据。

稳定发布必须同时具备：两个 Mac 架构公证和 Gatekeeper 验证、Windows 签名及安装卸载验证、源文件保护回归、实际 Word/WPS 文档验收记录、学校来源授权审查。当前没有这些完整证据，不把自动测试成功写成“正式版已认证”。

官方资料：[Apple 注册](https://developer.apple.com/programs/enroll/)、[Developer ID](https://developer.apple.com/developer-id/)、[微软签名选择](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/code-signing-options)、[Artifact Signing 地区限制](https://learn.microsoft.com/en-us/azure/artifact-signing/quickstart)、[SignPath 条件](https://signpath.org/terms.html)。
