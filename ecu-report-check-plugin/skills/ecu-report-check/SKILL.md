---
name: ecu-report-check
description: >
  校验ECU软件发布报告中的DID件号，或维护件号对照表。
  校验：用户提供报告文件和控制器名称，判断件号是否与标准值一致。
  维护：用户以自然语言或其他格式提供件号变更信息，自动更新对照表。
when_to_use: >
  用户提到"校验件号"、"检查报告"、"DID检查"、"ECU报告校验"、
  "part number check"、"更新件号"、"修改件号"、"新增控制器"、
  "删除件号"、"维护件号表"、或提供 .xlsx/.docx 报告文件要求校验时触发。
argument-hint: [文件路径] [控制器名 或 更新内容]
allowed-tools: Bash(python *) Read Glob Write Edit
---

## 功能一：校验件号

当用户提供报告文件要求校验时：

1. 确认 Python 环境可用，必要时安装依赖:
   ```bash
   pip install -r ${CLAUDE_SKILL_DIR}/scripts/requirements.txt
   ```

2. 确定报告文件路径。如果是相对路径，转为绝对路径。

3. 运行校验脚本:
   ```bash
   python "${CLAUDE_SKILL_DIR}/scripts/run_check.py" \
     --files <报告文件路径> \
     --controllers <控制器名> \
     --pn-file "${CLAUDE_SKILL_DIR}/part_number/part_number.xlsx"
   ```

4. 解析 JSON 输出，向用户汇报结果:
   - 如果 status 是 "PASS"：确认全部 OK，给出肯定回复
   - 如果 status 是 "ISSUES_FOUND"：列出 NG / MISSING / UNKNOWN 项的具体文件、Sheet、DID、期望值、实际值

## 功能二：维护件号表

当用户要求更新、新增、删除件号信息时：

1. **解析用户意图**。用户可能说自然语言，如：
   - "FLSMU 的 F1A0 改成 S000004325007" → update FLSMU/F1A0
   - "新增控制器 ABCPU，件号是..." → add controller
   - "把 F193 的件号删掉" → delete DID

2. **将自然语言转为 JSON 操作描述**，写入临时文件:
   ```json
   {
     "action": "update",
     "controller": "FLSMU",
     "dids": {"F1A0": "S000004325007"}
   }
   ```

3. **先 dry-run 预览变更**:
   ```bash
   python "${CLAUDE_SKILL_DIR}/scripts/manage_pn.py" \
     --pn-file "${CLAUDE_SKILL_DIR}/part_number/part_number.xlsx" \
     --input-type json --input-file /tmp/pn_update.json --dry-run
   ```

4. **向用户确认变更内容**，得到确认后再执行:
   ```bash
   python "${CLAUDE_SKILL_DIR}/scripts/manage_pn.py" \
     --pn-file "${CLAUDE_SKILL_DIR}/part_number/part_number.xlsx" \
     --input-type json --input-file /tmp/pn_update.json
   ```

5. **汇报执行结果**：成功/失败，具体变更了哪些件号。

6. 如果用户提供了其他格式的文件（如另一个 Excel 表格），先读取文件内容，
   理解其结构，再转化为标准 JSON 格式调用 manage_pn.py。

## 件号表结构说明

part_number.xlsx 的标准结构：
- C列: DID 代码
- H-L列: 5个控制器的件号（FLSPU, RLSPU, RRSPU, FLSMU, FRSMU）
- 每行一个 DID，每个控制器一列

有效的 DID 代码和含义：
| DID | 含义 |
|-----|------|
| F18E | ECU总成号 |
| F193 | 硬件号 |
| F180 | PBL（初始引导程序）|
| F104 | PBL诊断数据库 |
| F102 | SBL（二级引导程序）|
| F105 | SBL诊断数据库 |
| F1A0 | 应用层软件（SFA1）|
| F103 | APP诊断数据库 |
| F17F | Checksum |

## 结果解读规则

- **OK**: 实际件号与标准件号一致
- **NG**: 件号不匹配，需要修正
- **MISSING**: 报告中未找到该DID对应的件号，可能遗漏
- **UNKNOWN**: 找到件号但无法判断归属DID
