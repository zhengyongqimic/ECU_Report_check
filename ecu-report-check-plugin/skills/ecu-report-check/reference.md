# ECU Report Check — 参考文档

## 控制器列表

| 控制器代码 | 说明 |
|-----------|------|
| FLSPU | Front Left Seat Processing Unit（前左座椅处理单元）|
| RLSPU | Rear Left Seat Processing Unit（后左座椅处理单元）|
| RRSPU | Rear Right Seat Processing Unit（后右座椅处理单元）|
| FLSMU | Front Left Seat Massage Unit（前左座椅按摩单元）|
| FRSMU | Front Right Seat Massage Unit（前右座椅按摩单元）|

## DID 代码列表

| DID | 名称 | 说明 |
|-----|------|------|
| F18E | DU PN / ECU总成号 | 控制器总成件号 |
| F193 | HWBN / 硬件号 | 硬件版本件号 |
| F180 | PBL | Primary Boot Loader 件号 |
| F104 | SWDI-PBL | PBL 诊断数据库件号 |
| F102 | SBL | Secondary Boot Loader 件号 |
| F105 | SWDI-SBL | SBL 诊断数据库件号 |
| F1A0 | SFA1 / 应用层软件 | Application 层软件件号 |
| F103 | SWDI-SFA1 | APP 诊断数据库件号 |
| F17F | Checksum | 校验和（十六进制数据）|

注：SBL（F102/F105）仅部分控制器适用，FLSMU 和 FRSMU 通常不包含。

## part_number.xlsx 布局

```
行1:  [A-G 各种标签/表头] [H] FLSPU [I] RLSPU [J] RRSPU [K] FLSMU [L] FRSMU
行2:  [...] [C] F18E  [H] 件号值  [I] 件号值 ...
行3:  [...] [C] F193  [H] 件号值  [I] 件号值 ...
行4:  [...] [C] F180  [H] 件号值  [I] 件号值 ...
行5:  [...] [C] F104  [H] 件号值  [I] 件号值 ...
行6:  [...] [C] F102  [H] 件号值  [I] 件号值 ...
行7:  [...] [C] F105  [H] 件号值  [I] 件号值 ...
行8:  [...] [C] F1A0  [H] 件号值  [I] 件号值 ...
行9:  [...] [C] F103  [H] 件号值  [I] 件号值 ...
行10: [...] [C] F17F  [H] 件号值  [I] 件号值 ...
```

- C列（列索引3）: DID 代码
- H-L列（列索引8-12）: 各控制器件号
- 件号格式如 `S000004325006`，F17F 为十六进制如 `0102030405060708`
- 无适用件号时填 `-` 或留空
