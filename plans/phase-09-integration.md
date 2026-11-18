# Phase 9：集成与验收

**依赖：Phase 1–8 全部完成。**

## 1. 阶段目的

验证所有模块在目标 Windows/Excel/Samba 环境下按总计划工作，覆盖多人不同列并行、同列互斥、冲突裁决、失败保留和恢复边界，并产出可分发的纯 VBA `.xlam`。

## 2. 实现方式

- 按依赖顺序构建：Phase 1–4 纯逻辑/基础组件，Phase 5–6 锁和 Excel 适配，Phase 7–8 生命周期与 UI，最后打包并按干净用户配置验证加载项。
- 用合成工作簿和隔离共享目录执行端到端测试。至少两台 Windows 客户端同时运行，确保测试数据不包含生产数据或真实 UNC/服务器标识。
- 验证不同 sheet+AMG 列锁可并行；不同会话可以编辑同一 sheet 的不同列；提交期间通过 Phase 5 的短期提交门闩串行访问共享工作簿。相同锁冲突必须被拒绝。
- 覆盖整组签出失败回滚、Local/Remote 冲突逐格选择、全部裁决前不写共享、取消后本地会话/列锁保留、仅声明列写回、Save 成功/失败、放弃、锁过期与失联客户端。
- 验收共享路径无临时文件创建；本地文件仅位于 `%LOCALAPPDATA%`；配置只存 UNC；保存只调用 `Save()`；`.xlam` 无运行时第三方依赖。
- 构建/集成脚本放 `tools/`，测试放 `test/`。任何可能长时间运行的 Excel/VBA 自动化必须有硬超时、可追踪 PID，并在测试结束时确认进程退出；不留后台进程和运行产物。

## 3. 涉及函数与模块

模块：全部 `MdlInfra`、`MdlStorage`、`MdlSchema`、`MdlDiff`、`MdlMerge`、`MdlLock`、`MdlSession`、`MdlExcel`、`MdlUI`；构建辅助脚本位于 `tools/`。

建议端到端入口：

```text
Public Function RunCheckoutCommitScenario(testCase As TestScenario, ByRef result As TestResult) As Boolean
Public Function RunTwoClientConcurrencyScenario(testCase As ConcurrencyScenario, ByRef result As TestResult) As Boolean
Public Function VerifyPackage(packagePath As String, ByRef result As PackageCheckResult) As Boolean
Private Sub CleanupTestArtifacts(testWorkspace As String, startedProcessIds As Variant)
```

这些是测试/验收驱动接口，不属于运行时产品 API；Excel/VBA 进程的启动、等待、超时和退出检查均由受控测试驱动管理。

## 4. 功能范围

做：单元/组件/集成/端到端回归、双客户端并发、打包 `.xlam`、干净环境加载验证、性能与失败行为记录、按结果修订对应计划。

不做：生产数据迁移、部署到用户机器、推送共享分支、引入 M365/第三方运行依赖、扩展到行列增删或实时协同、将准备阶段旧方案直接并入实现。

## 5. 测试方案

- 先运行 `test/MdlInfra/`、`MdlSchema/`、`MdlDiff/`、`MdlMerge/` 的无 Excel 测试，再运行 `MdlStorage/`、`MdlLock/` 组件测试，最后运行 `MdlExcel/`、`MdlSession/`、`MdlUI/` 集成与端到端测试。
- 双客户端场景：不同列同时签出并分别提交；同一列竞争；批量目标中一个锁冲突；不同列会话同时提交；一个客户端心跳停止后另一客户端检查过期回收。
- 数据正确性场景：无冲突、仅 Local 改、仅 Remote 改、同值改动、同格双方不同改动并逐项选 Local/Remote、裁决中取消、Remote 在裁决期再次变化、锚点值/顺序变化阻止写回、结构变化阻止写回、只写声明列。
- 故障场景：UNC 不可达、权限不足、锁目录/owner 损坏、复制失败、工作簿打开失败、写入/Save 失败或结果不确定、进程超时/退出、清理失败。每次验证本地数据保留与锁状态符合计划。
- 验收记录不提交真实路径、账号、服务器名或业务内容；测试后确认 Excel/VBA 进程已结束，清理临时工作簿和本地/共享测试目录。

## 6. 风险与未决问题

- **真实 Samba 语义与本地模拟存在差异。** `MkDir` 互斥、心跳时间戳、目录删除、共享 `Save()` 期间的读写竞争必须双客户端实测；未通过时不能以 M4 生产就绪验收。
- **提交互斥实现及 fencing 的最终边界待验证。** 暂按所有共享工作簿写入串行、每次 Save 前重验 lease/fencing；需确认不同列编辑仍可并行且持锁不会长期占用共享文件。
- **多列签出整组失败**按当前保守默认验收；等待用户确认是否支持跳过忙碌列。
- 当前验收只覆盖声明列的单元格值，不测试或宣传格式同步能力。
- M4 还依赖保存异常后的可判定性、孤儿工作区处理规则、宏策略/加载项注入在目标 Excel 上可用；这些结果应记录为明确通过、失败或待确认，不以计划文档代替实测。
