# ECU Report Check Plugin

Claude Code 插件，用于自动校验 ECU 软件发布报告中的 DID 件号，以及维护件号对照表。

## 安装

```bash
claude plugin install ecu-report-check@https://github.com/<user>/ecu-report-check-plugin
```

## 功能

### 件号校验

对话中直接说：
- "帮我校验这份报告的件号：E:/path/to/report.xlsx FLSMU"
- `/ecu-report-check E:/path/to/report.xlsx FLSMU`

Claude 会自动运行校验脚本，比对报告中的件号与标准值，汇报 PASS/ISSUES_FOUND 及详情。

### 件号表维护

对话中直接说：
- "FLSMU 的 F1A0 件号改成 S000004325007"
- "新增控制器 ABCPU 的件号信息"
- "查看当前件号表"
- "删除 RLSPU 的 F193 件号"

Claude 会先 dry-run 预览变更，确认后再执行修改。修改前自动备份。

## 支持的控制器

FLSPU, RLSPU, RRSPU, FLSMU, FRSMU（支持新增自定义控制器）

## 支持的 DID

F18E, F193, F180, F104, F102, F105, F1A0, F103, F17F

详见 `skills/ecu-report-check/reference.md`。

## 前置条件

- Python 3.10+
- 依赖包：`pip install openpyxl python-docx`

## 目录结构

```
.claude-plugin/plugin.json       # Plugin manifest
skills/ecu-report-check/
  SKILL.md                       # Skill 定义
  reference.md                   # DID/控制器参考文档
  scripts/
    run_check.py                 # 件号校验脚本
    manage_pn.py                 # 件号表维护脚本
    requirements.txt             # Python 依赖
part_number/part_number.xlsx     # 标准件号数据
```
