# Physics Lab Report Skill

面向 Codex 等支持 `SKILL.md` 的代理的中文物理实验报告排版技能。

保留实验名称和姓名学号栏的真实下划线；优先使用横向宽表；图、表下方分别编号；使用可编辑 Word 公式，并在交付前检查实际分页。

![空白模板预览](assets/preview.png)

## 安装

将本仓库完整复制到个人技能目录中的 `artifact-template-physics-lab-report` 文件夹：

```powershell
git clone https://github.com/sxdgsy/report-skill.git
Copy-Item -LiteralPath .\report-skill -Destination "$env:USERPROFILE\.codex\skills\artifact-template-physics-lab-report" -Recurse
```

若目标目录已存在，请先比较并合并自己的修改，不要覆盖个人数据。使用其他代理时，按其技能目录约定安装。若应用尚未列出技能，重新加载技能列表或开启新会话。

## 使用

```text
使用 $artifact-template-physics-lab-report，根据我上传的原始数据和课程要求完成 Word 报告。
只包含原始数据、数据分析、课后题三部分。
```

本次课程模板和用户要求优先。姓名、学号、班级、座位号、日期均留空，必须由本次资料提供。不要将“开源技能”误认为可以代造实验数据。

## 脚本

需要 Python 3.10+ 和 `python-docx`。首次使用时在自己的环境安装依赖：

```sh
python -m pip install -r requirements.txt
python scripts/report_style.py --experiment "实验名称" --output build/report.docx
python scripts/audit_layout.py build/report.docx
python -m unittest discover -s tests -v
```

`report_style.py` 提供 `create_report`、`add_figure`、`add_wide_table`；只创建版式，不代填实验内容。辅助函数默认等宽列，复杂长标签可在调用后调整列宽；测点转置由调用者根据实验含义决定。

`audit_layout.py` 检查信息栏下划线、原生图片/表格下方题注及表格宽度。它不是完备的 Word 校验器，对复杂文本框、嵌套表格等对象需要人工复核。结构检查通过不代表排版通过；仍需用 Word 或 LibreOffice 渲染，逐页查看照片、公式、题注和分页。

宋体、黑体等字体不随仓库分发；系统缺少字体时须安装有权使用的字体或明确选择替代字体。Documents 技能可协助渲染；脚本自身不依赖该插件或本机 Office 自动化。

## 公开内容边界

仓库只包含抽象版式规则、脚本和从零生成的空白 DOCX 及预览；不包含个人实验原件、姓名学号、教师签章或设备本地路径。示例或作业如需公开，请由资料所有者另行确认权限并处理个人信息。

许可证：[MIT](LICENSE)。
