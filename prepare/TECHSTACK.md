# 技术栈

## 结论

> **运行时零外部依赖：纯 VBA + Excel 自身 + Windows 内置 API。**
> 不需要 Python、不需要 .NET、不需要 COM 注册、不需要管理员权限、不需要安装程序。

`.xlam` 拷进 `%APPDATA%\Microsoft\AddIns\` 就是安装完成。这是本方案选 VBA 的主要理由。

构建期（只在开发机跑一次）可以用 PowerShell 打包；诊断期可以用 PowerShell/Python 一次性分析文件。两者都不进入用户机器。

---

## 1. 逐项需求 → 用什么

| 需求 | 实现 | 依赖 | 风险 |
|---|---|---|---|
| 网络盘复制大文件 | `FileCopy`（或 FSO `CopyFile`） | 无 | 低 |
| SMB 互斥锁 | `Open path For Binary Access Write Lock Read Write` | 无 | **中，需实验** |
| 原子替换文件 | `MoveFileEx(MOVEFILE_REPLACE_EXISTING)` Declare | kernel32 | **中，需实验** |
| 状态文件读写 | 自己写**行式 TSV/INI**，不用 JSON | 无 | 低 |
| 区域哈希 | 手写 FNV-1a（两段 32 位拼 64 位） | 无 | 低 |
| 单元格差异 | `Range.Value2` 读入 Variant 数组 + `Range.Interior.Color`/`Font.Color` | 无 | 低 |
| 跨工作簿搬区域 | `Range.Copy Destination:=` | 无 | 低（实验 1） |
| 超链接清点 | `Range.Hyperlinks.Count` | 无 | 低 |
| 多选 sheet 勾选窗体 | `UserForm` + `ListBox(MultiSelect:=fmMultiSelectMulti, ListStyle:=fmListStyleOption)` | 无 | 低 |
| 带按钮的功能区 | Custom UI XML 嵌入 `.xlam` 包 | 无 | 低（构建期） |
| 毫秒级 sleep | `Declare PtrSafe Sub Sleep Lib "kernel32"` | kernel32 | 低 |
| 文件大小/时间戳 | `FileLen` / `FileDateTime` | 无 | 低（仅 UI 用途） |
| 用户目录 | `Environ("LOCALAPPDATA")` / `Environ("APPDATA")` | 无 | 低 |
| 受信任位置自检 | `WScript.Shell.RegRead` 读 `HKCU\...\Trusted Locations` | wshom.ocx（系统自带） | 低 |
| 审计日志追加 | 原生 `Open For Append` | 无 | 低 |
| Excel 版本判断 | `Application.Version`（"16.0"）+ `Len(Environ("PROCESSOR_ARCHITECTURE"))` | 无 | 低 |

---

## 2. 明确不用的东西（以及为什么）

| 诱惑 | 为什么不用 |
|---|---|
| `CreateObject("System.Security.Cryptography.SHA256Managed")` | 依赖 .NET→COM 注册，部分锁定环境被封，且 32/64 位行为不一致 |
| `JsonConverter`（VBA-JSON，第三方） | 需要导入模块 + `Scripting.Dictionary`；而我们只需要存 5 条 claim → **改成行式 TSV 更简单更安全** |
| Python + `pywin32` | 需要给每台机器装 Python；而所有事 VBA 都能做（本方案已在 DESIGN 中砍掉了需要 Python 的三路合并） |
| openpyxl | 会破坏图表/条件格式/超链接，且要 Python |
| VSTO / C# COM Add-in | 需要 .NET 运行时 + 注册（COM Add-in 要管理员）+ ClickOnce；在受管控的 RDP 上部署是灾难 |
| Office.js / Web Add-in | 需要 HTTPS 托管 + 清单部署，且无法做 SMB 锁；与"数据不能出内网"约束冲突 |
| `CommandBars` 加按钮 | 已弃用、图标丑、部分版本不稳定 → 用 Custom UI XML |
| 第三方 zip 库读 xlsx | 我们用 Excel 本体读写，不需要碰 zip（诊断脚本除外，那是一次性工具） |

---

## 3. 三个需要实测的 Windows/SMB 行为

这三个不是"选型问题"，是"能不能这么干"的问题，各测 10 分钟即可定论。

### 实验 A：VBA `Open ... Lock` 在 SMB 上是否真互斥

```vba
' 机器 1 运行
Open "\\server\share\_sync\test.lock" For Binary Access Write Lock Read Write As #1
' 机器 2 同时运行，预期 Err 70 (Permission denied)
```

- 通过 → 锁方案成立
- 不通过 → 退化为"`CreateFile` + 独占打开"（`Declare` Win32）或"mkdir 原子性"

### 实验 B：`MoveFileEx` 能否覆盖一个正被他人打开的文件

```
机器 1：用 Excel 打开 book.xlsx
机器 2：cmd → move /y new.xlsx book.xlsx
```

- 通过 → 可采用 `PROCESS.md §8.3` 的"写临时 + 重命名换页"，**提交时完全不锁文件，占用窗口 20s → 2s**
- 不通过 → 保守走 `Workbook.Open` + `Save`

### 实验 C：`Range.Copy Destination:=` 跨工作簿是否完整保留超链接 + 填充色

这是**技术前提**。若 `Hyperlinks.Count` 变 0，就必须改成"逐属性写入 + 显式 `Hyperlinks.Add` 重建"，工作量增加约 1 天。

---

## 4. 为什么不存 JSON

原设计的 `owners.json` 改成 `owners.tsv`：

```
# version=1  fileId=book.xlsx
sheet	user	since	ttlMin	mtime	size	baseHash
Sheet3	DOMAIN\zhangsan	2025-11-15T10:23:11	120	15T09:40:02	10485760	a1b2c3d4
Sheet7	DOMAIN\lisi	2025-11-15T10:31:02	120	15T10:20:11	10485760	9f8e7d6c
```

理由：

- VBA 无原生 JSON，任何解析器都是外部代码 + 一个 bug 面
- 我们只有 5 行、7 列、且**全程持锁读改写**，TSV 足够
- 天然 diff/肉眼可读，出问题能手工修
- sheet 名可能含 tab？不可能（Excel 禁止），也不含换行 → TSV 安全

同理，审计日志改用行式文本，不用 JSONL。

---

## 5. 部署（无管理员权限）

| 步骤 | 做法 | 需要管理员？ |
|---|---|---|
| 安装加载项 | 拷 `.xlam` 到 `%APPDATA%\Microsoft\AddIns\` | 否 |
| 注册加载项 | 写 `HKCU\Software\Microsoft\Office\Excel\Addins\ExcelSync\Open=1` | 否 |
| 加受信任位置 | GPO 或用户手工（一次性，最好 IT 统一推） | **可能需要** |
| 宏策略 | 若组织要求"仅允许签名宏" → 需代码签名证书 | **可能需要** |

最后一行是本方案唯一可能的部署阻塞点，建议提前确认：

- 若策略是"禁用所有宏，但受信任位置例外" → 正常
- 若策略是"仅允许数字签名的宏" → 需自签证书 + 推送到 Trusted Publishers，或让 IT 签

**建议提前问 IT 的这一句**：*受信任位置里的 `.xlam`，宏会被拦截吗？*

---

## 6. 构建期工具（只在开发机）

| 用途 | 工具 | 说明 |
|---|---|---|
| 把 Custom UI XML 注入 `.xlam` | PowerShell（`System.IO.Compression`） | 改 `customUI/customUI14.xml` + `[Content_Types].xml` + `_rels/.rels`，一次性脚本 |
| 打包分发 | PowerShell + `.cmd` | 一个 zip，用户解压运行 `install.cmd` |
| 诊断 10MB 构成 | PowerShell（已给过）或 Python | 一次性，不进用户机器 |
| 代码检查 | `Rubberduck` VBE 插件（可选） | 仅开发机，提升代码质量 |

运行时**一个字节的第三方代码都不进用户机器**。

---

## 7. 代码组织

```
ExcelSync.xlam
├─ ThisWorkbook            ' 版本号、加载/卸载
├─ customUI/customUI14.xml ' 功能区定义（构建期注入）
├─ 标准模块
│   ├─ modConfig           路径、注册表、版本、受信任位置自检
│   ├─ modLock             VBA Open Lock 封装 + TTL 接管
│   ├─ modClaim            owners.tsv 读改写（全程持锁）
│   ├─ modWorkspace        本地副本、base 快照、崩溃恢复
│   ├─ modKey              normalizeKey、key 提取、唯一性自检
│   ├─ modHash             FNV-1a 区域哈希
│   ├─ modDiff             base↔local 差异（预览用）
│   ├─ modGuard            base↔server 哈希 + key 级分类（冲突用）
│   ├─ modCommit           提交事务：开→清→粘→校验→存
│   ├─ modAudit            审计日志
│   └─ modWin32            Declare PtrSafe：Sleep、MoveFileEx
└─ 窗体
    ├─ frmPickSheets
    ├─ frmDiffPreview
    ├─ frmConflict
    └─ frmWhoIsEditing
```

约 2,500–3,500 行 VBA。VBA 工程建议设密码防止误改（不防有心人，只防手滑）。

---

## 8. 32/64 位与版本兼容

- 所有 `Declare` 必须 `PtrSafe` + `LongPtr`（否则 64 位 Office 直接编译失败）
- `Application.Version` 判版本；目标范围：Excel 2016/2019/2021/365 桌面版，32 与 64 位都支持
- `Range.Copy` / `Hyperlinks` / `ListBox` 在这些版本里行为一致
- 不支持：Excel 2007（Custom UI XML 用旧格式）、Mac Excel（无 Windows 路径/SMB 锁）

---

## 9. 一句话回答

**纯 VBA `.xlam`，零运行时依赖。** 所有需求都命中 Windows/Office 内置能力，唯一需要先验证的是三件事：SMB 字节锁（实验 A）、`MoveFileEx` 覆盖语义（实验 B）、跨工作簿 Copy 是否保住超链接（实验 C）。这三个实验决定实现细节，但都不改变"纯 VBA"这个选型。