# ThesisCraft 架构

作者：Study-Tang。

```
桌面界面 / CLI / Skill / Word-WPS 工具栏
                     ↓
              config → jobs → service
                                ↓
                              engine
                ┌───────────────┼──────────────┐
            conversion     recognition      layout
                           text / ooxml     tables
                                ↓
                             storage
```

- `config.py`：配置默认值、旧字段迁移、类型和范围验证。
- `jobs.py`：收集和去重输入、保留目录结构、输出命名、覆盖策略。
- `service.py`：批处理、取消检查与成功/失败/跳过的结构化结果。
- `engine.py`：排版流程与资源清理，供所有入口复用。
- `conversion.py`：Word/WPS COM 管理与 LibreOffice 转换。
- `text.py`：标点、Markdown 和空行处理。
- `ooxml.py`：保留其他节点的文字修改、带链接的 run 遍历和安全拆分。
- `recognition.py`、`layout.py`、`tables.py`：识别、样式及表格职责。
- `storage.py`：原件路径保护、临时文件完成后发布输出、用户配置位置。
- `gui.py`、`gui_view.py`：界面控制器与控件视图分离。
- `office_toolbar.py`：通过 Windows COM 按钮事件接入 Word/WPS；工作线程处理文件，宿主线程打开结果。
- `office.py` 与 `addins/windows/FormatterAddin.cs`：实验性的标准 COM DLL 接口。

`wfp.py`、`wfp_cli.py`、`wfp_core.py` 等入口共享内部模块。业务代码只维护 `word_formatter` 一份，独立 Skill 由 `packaging/build_skill.py` 生成。

## 默认行为

1. 默认保留原文字颜色和原分页属性；需统一黑色时设置 `force_black_text=true`。
2. 默认不接受修订或把自动编号变为文字；桌面/CLI 可显式启用 `preprocess_office=true`，插件固定关闭此预处理。
3. 默认保护已有输出；自动命名时增加序号，显式指定已有输出时要求 `--overwrite`。无论是否指定覆盖，都拒绝输入与输出同一文件或硬链接。
4. 页码设置保留页脚其他内容；可用 `page_numbers=false` 完全不改页码策略。
5. 未知配置和无效类型会尽早报错，不再静默接受字符串布尔值或错误数值。
6. CLI：0 表示全部成功，1 表示失败，2 表示部分跳过/取消。`--json` 可获取结构化摘要。
7. GUI 默认配置存储在用户配置目录。Windows 路径为 `%APPDATA%/Study-Tang/ThesisCraft/config.json`。旧启动目录的 `default_config.json` 仍可迁移；CLI 保留当前目录 `wfp_config.json` 的优先级。
8. 插件工具栏需要主动启动，当前不自动加载到所有 Office 实例。

## 输出安全

先将 docx 完整写入输出目录中的临时文件。禁止覆盖时通过硬链接发布已完成文件，目标已存在则失败；允许覆盖时使用原子替换。平台或文件系统不支持硬链接时会报错，保留原件和旧输出，不退回到可能截断文件的非原子写法。

## 测试边界

原有 24 项测试与 21 项新增回归测试分别运行。另有真实 Word/WPS 工具栏集成测试。CI 配置覆盖三个系统与 Python 3.10/3.12，Python 核心已通过三系统 CI；真实 Office 验收在 Windows 上执行。

复杂 Word 文档仍有能力边界：未全面重排页眉、脚注、文本框和修订内部内容；Markdown 不做完整语义渲染；真实复杂文档的分页效果仍需要在 Word/WPS 中检查。
