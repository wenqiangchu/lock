# Phase 9：集成与验收

**依赖：Phase 1–8 全部完成。**

## 1. 阶段目的

验证所有模块在目标 Windows/Excel/Samba 环境下按总计划工作，覆盖多人不同列并行、同列互斥、冲突裁决、失败保留和恢复边界，并产出可分发的纯 VBA `.xlam`。

## 2. 实现方式

- 按依赖顺序构建：Phase 1–4 纯逻辑/基础组件，Phase 5–6 锁和 Excel 适配，Phase 7–8 生命周期与 UI，最后用 `tools/Build-Addin.ps1` 打包并按干净用户配置验证加载项。脚本将输出路径规范化，并限制在仓库 `build/` 目录内；路径校验需兼容 Windows PowerShell 5.1 所用的 .NET Framework。
- 构建脚本在隔离 Excel 进程中导入源码和 Office 自带 Forms 引用，生成初始 `.xlam`；生成包后用 XML/ZIP API 注入 `src/MdlUI/ribbon.xml` 并验证 `vbaProject.bin`、Ribbon 关系、`/customUI/customUI14.xml` 的 `application/xml` Content Type 和三个回调。包结构测试必须与 Excel 实际显示 Ribbon 标签的冒烟测试配合；仅结构自检不能替代宿主应用加载验证。Excel `SaveAs` 仅用于首次创建加载项包；共享业务文件运行时保存仍只用 `Save()`。
- 用合成工作簿和隔离共享目录执行端到端测试。至少两台 Windows 客户端同时运行，确保测试数据不包含生产数据或真实 UNC/服务器标识。
- 验证配置可列出多份分散路径的工作簿；不同会话分别选择不同工作簿时，本地工作区、锁键和提交门闩按文件路径隔离；同一会话始终只操作一份工作簿。另验证同一文件的不同 sheet+AMG 列锁可并行、相同锁冲突被拒绝，提交期间通过 Phase 5 的短期提交门闩串行访问对应共享工作簿。
- 覆盖整组签出失败回滚、值/文字字体/填充组件的 Local/Remote 冲突逐项选择、全部裁决前不写共享、取消后本地会话/列锁保留、仅声明列值与格式写回、Save 成功/失败、放弃、锁过期与失联客户端。
- 验收共享路径无临时文件创建；本地文件仅位于 `%LOCALAPPDATA%`；手工配置清单只存 UNC 且拒绝映射盘符；保存只调用 `Save()`；`.xlam` 无运行时第三方依赖。
- 构建/集成脚本放 `tools/`，测试放 `test/`。任何可能长时间运行的 Excel/VBA 自动化必须有硬超时、可追踪 PID，并在测试结束时确认进程退出；不留后台进程和运行产物。

## 3. 涉及函数与模块

模块：全部 `MdlInfra`、`MdlStorage`、`MdlSchema`、`MdlDiff`、`MdlMerge`、`MdlLock`、`MdlSession`、`MdlExcel`、`MdlUI`；构建辅助脚本位于 `tools/`。

当前可运行的本机验证入口：

```text
tools/Run-VbaTests.ps1 [-TestMacro <VBA 宏名>] [-TimeoutSeconds 120]
Private Invoke-TestWorker(OutputPath, KnownPids, MacroName)
tools/Build-Addin.ps1 [-OutputPath <build 内的 .xlam 路径>] [-TimeoutSeconds 120]
Private Build-Workbook(PackagePath, KnownPids)
Private Add-RibbonPart(PackagePath, RibbonPath)
Private Test-AddinPackage(PackagePath)
tools/Probe-ExcelBoundary.ps1 [-TimeoutSeconds 60]
test/tools/test_addin_package.py <package>
test/tools/test_probe_xlsx_structure.py
```

这些脚本不属于加载项运行时 API。当前仓库没有完整签出→提交的 Phase 9 双客户端端到端测试驱动；不可用的真实模板、Samba 共享和交互式 Excel 验收必须继续标记为未验证，不得以建议接口或包结构检查代替。

## 4. 功能范围

做：单元/组件/集成/端到端回归、双客户端并发、打包 `.xlam`、干净环境加载验证、性能与失败行为记录、按结果修订对应计划。

不做：生产数据迁移、部署到用户机器、推送共享分支、引入 M365/第三方运行依赖、扩展到行列增删或实时协同、将准备阶段旧方案直接并入实现。

## 5. 测试方案

- 运行 `tools/Run-VbaTests.ps1 -TimeoutSeconds 120`；内存逻辑与本地组件测试不读取 Excel 工作簿/对象模型，集成测试使用合成工作簿。
- 用唯一输出名运行 `tools/Build-Addin.ps1 -OutputPath build/MDDTLockIssue-<run-id>.xlam -TimeoutSeconds 120`，worker 有 120 秒硬超时，随后最多等待 20 秒让隔离 Excel 退出；超时则终止该进程并将构建判为失败。之后运行 `python -B test/tools/test_addin_package.py build/MDDTLockIssue-<run-id>.xlam` 检查 VBA 工程部件、Ribbon 按钮、回调、关系和内容类型。测试后删除本次生成的包。
- 运行 `tools/Probe-ExcelBoundary.ps1 -TimeoutSeconds 60` 与 `python -B test/tools/test_probe_xlsx_structure.py`；归档样本的匿名结构统计使用 `python -B tools/probe_xlsx_structure.py prepare/测试结果.xlsx prepare/_smoke/schema_probe.xlsx` 复现。
- 双客户端场景：不同列同时签出并分别提交；同一列竞争；批量目标中一个锁冲突；不同列会话同时提交；一个客户端心跳停止后另一客户端检查过期回收。
- 数据正确性场景：无冲突、仅 Local 改、仅 Remote 改、同值改动、同格同组件双方不同改动并逐项选 Local/Remote、Local 改字体而 Remote 改填充时自动组合、局部字符颜色/粗体与多 run 往返、空白/零长度字符串的单元格级字体、文本插入/删除后的字体映射处理、填充变化及冲突、裁决中取消、Remote 在裁决期再次变化、锚点值/顺序变化阻止写回、结构变化阻止写回、仅声明列值与格式写入、非字符串值/公式拒绝、字符串大小写/空格和空白哨兵比较、内部全空行保留物理位置、末行空锚点但声明值存在时拒绝写回、格式延伸到末行不扩大数据行边界。
- 故障场景：配置缺失/损坏、清单中某个文件不存在、UNC 不可达、访问拒绝、Excel 以只读方式打开、损坏/不支持的工作簿、锁目录/owner 损坏、复制失败、写入/Save 失败或结果不确定、进程超时/退出、清理失败。验证只读提示说明可能原因但不把锁冲突说成已确认；未开始签出时不产生锁/本地副本，失败会话保留数据且锁状态符合计划。
- 验收记录不提交真实路径、账号、服务器名或业务内容；测试后确认 Excel/VBA 进程已结束，清理临时工作簿和本地/共享测试目录。
- Phase 2/6 的锚点和数据行边界结论必须区分“合成/隔离样例验证”与“代表性模板实测”；没有代表性模板时，M4 不得标记生产就绪。
- 2026-11-19 本机验证：VBA 聚合测试在 Excel 16.0 build 19127 通过；Excel 合成边界探测和 OpenXML 脱敏结构测试通过；修复后的正式包 `MDDTLockIssue-ribbon-fix-20261119.xlam` 通过 `test/tools/test_addin_package.py` 检查，并在 Office 2019 Version 2508 (Build 19127.20860) 中加载后显示 `MDDT Collaboration` 标签。加载测试中旧包和临时探测包均未启用，临时探测文件已删除。
- 当前未验收：代表性业务模板、至少两个独立 Windows 客户端的 Samba 并发、Ribbon 按钮与窗体完整交互、干净用户环境加载和共享 Save 失败结果。上述项目仍是 Phase 9 验收项；本机验证只证明正式包结构正确且宿主 Excel 显示 Ribbon 标签，不表示端到端流程或生产环境验收完成。

## 6. 风险与未决问题

- **真实 Samba 语义与本地模拟存在差异。** `MkDir` 互斥、心跳时间戳、目录删除、共享 `Save()` 期间的读写竞争必须双客户端实测；未通过时不能以 M4 生产就绪验收。
- **双客户端和代表性业务模板当前可能不可用。** 缺少目标 Samba 共享或经确认的模板时，只能验证本机逻辑、合成 Excel 和包结构；Phase 9 记录为外部环境阻塞，不能把本地测试等同于生产验收。
- **提交互斥实现及 fencing 的最终边界待验证。** 暂按所有共享工作簿写入串行、每次 Save 前重验 lease/fencing；需确认不同列编辑仍可并行且持锁不会长期占用共享文件。
- **Excel/Samba 对“被锁定”与“无写权限”的报告可能相同。** 当前不单独识别 sharing violation；只读打开使用“无法取得写入访问，可能由占用、权限或只读属性造成”的表述。若目标 Excel/Samba 后续提供稳定、可直接读取的锁冲突信号，再评估是否增加分类，不解析模糊描述来猜测。
- 字体 run 和填充格式会增加快照读取、比较、冲突展示与写回工作量；需要记录代表性规模的签出/提交耗时，并验证 Theme/Automatic 颜色及混合格式可稳定往返。任何未能无损往返的属性必须失败关闭。
- M4 还依赖保存异常后的可判定性、孤儿工作区处理规则、宏策略/加载项注入在目标 Excel 上可用；这些结果应记录为明确通过、失败或待确认，不以计划文档代替实测。
