# 桌面构建与兼容边界

## 源码运行

需要 Python 3.10+ 和 Tk；构建使用 Python 3.12。Windows 安装 `.[desktop,windows]`，macOS 安装 `.[desktop]`。

```sh
python -m pip install -e '.[desktop,dev]'
python -m word_formatter
python -m word_formatter --general
python -m unittest discover -s tests -v
python wfp_cli.py test
python -m word_formatter --self-check self-check.json
```

默认入口与 `thesiscraft-desktop` 打开简洁论文首页；`--academic` 是保留的旧入口，`--general` / `wfp-desktop` 打开通用批处理。首个位置参数可传入 DOCX 路径。检查不要求安装 Word/WPS。

## 原生构建

```sh
# 在对应系统运行；禁止把其他架构的包改名发布
python packaging/build_release.py macos
python packaging/build_release.py windows
```

脚本创建隔离构建环境，收集 CSL 与 python-docx 等资源，构建后启动真实二进制完成自检，再生成校验和。macOS 在资源收集完成后由 PyInstaller 签名，不再向已签名 bundle 写入资源。Windows 使用 onedir 目录布局，输出 portable.zip 和 Inno Setup 6 安装包；Mac 输出 `.app.zip` 与 DMG。Windows 构建机需安装 Inno Setup 6。

[Desktop builds](../.github/workflows/desktop.yml) 在 Windows x64、Mac arm64 和 Mac Intel 上分别运行，构建结果保存在 Actions artifacts。只有显式 `v*-beta.*` tag 才允许通过未签名流水线发布，所有平台构建成功后才发布；普通分支提交不发布软件。

## 功能矩阵

| 功能 | Windows | macOS |
| --- | --- | --- |
| 本地 DOCX 排版、独立副本、检查报告 | 支持 | 支持 |
| 模板、字体、图表、编号、文献设置 | 支持 | 支持 |
| 使用默认办公软件打开结果 | 支持 | 支持 |
| Word/WPS 自动更新目录与 PDF 导出 | 需要桌面 Office 与 COM | 在办公软件中手动完成 |
| 原生功能区、登录时连接 Office | 另用原生插件安装包 | 不支持 |

Mac 上隐藏 Windows 专属按钮，保留系统默认应用打开 DOCX。模板中的字体名称不自动替换：若本机缺少指定字体，应安装有授权的字体或按学校要求修改模板。不同办公软件的分页仍需人工核对。

自动自检验证应用能够创建、排版、重新打开 DOCX，保护原文件字节，生成 HTML 报告，加载三种 CSL，并初始化和切换 GUI。桌面额外写出完整报告，CLI 默认只生成副本；“只检查”始终保存报告。CI 验证不包括商业 Office 的视觉输出、真实学校验收或发行证书。公开传播前仍需真实论文试用；Developer ID 公证和 Authenticode 签名需要维护者自己的证书。

签名入口见 [SIGNING.md](SIGNING.md)，实际测试与未验证范围见 [COMPATIBILITY.md](COMPATIBILITY.md)。
