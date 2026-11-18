# Phase 1：基础设施

**依赖：Phase 0 已完成。**

## 1. 阶段目的

建立所有后续模块共用的基础能力：本地日志、统一错误表达、配置读写、UNC 路径校验与路径解析、共享路径延迟探测。完成后，调用方能够得到可记录、可展示、可测试的明确结果；配置始终保存 UNC 路径，不依赖映射盘符。

## 2. 实现方式

- 在 `MdlInfra` 中定义稳定的错误码、日志级别和轻量错误信息结构。公共过程返回成功标志与错误信息，不让底层错误静默消失。
- 配置读写仅接受 UNC 主文件路径；遇到盘符路径时拒绝保存并说明原因。另提供路径解析能力供诊断或迁移输入使用，但解析结果必须经过 UNC 校验后才能成为配置值。
- 延迟探测只对已有共享文件或目录做只读访问计时，不在共享上创建探测文件。探测失败要返回错误状态，不把失败解释成“共享文件不存在”。
- 日志写入用户本机 `%LOCALAPPDATA%` 下的应用目录，避免业务数据和运行文件写入共享；具体日志保留策略在风险项中待确认。
- `MdlInfra` 不依赖其它产品模块，不引用 Excel 对象模型。

## 3. 涉及函数与模块

模块：`src/MdlInfra/`，不新增产品层依赖。

建议过程接口：

```text
Public Sub LogWrite(level As LogLevel, source As String, message As String)
Public Function MakeAppError(code As Long, message As String, context As String) As AppErrorInfo
Public Function ReadConfig(configPath As String, ByRef config As AppConfig, ByRef err As AppErrorInfo) As Boolean
Public Function WriteConfig(configPath As String, config As AppConfig, ByRef err As AppErrorInfo) As Boolean
Public Function ValidateUncPath(path As String, ByRef normalizedPath As String, ByRef err As AppErrorInfo) As Boolean
Public Function ResolvePathToUnc(path As String, ByRef uncPath As String, ByRef err As AppErrorInfo) As Boolean
Public Function ProbeUncLatency(path As String, ByRef elapsedMs As Double, ByRef err As AppErrorInfo) As Boolean
```

`AppConfig` 只包含 UNC 主文件路径及本阶段确有必要的版本字段；不在此阶段加入用户偏好或业务状态。`AppErrorInfo` 至少表达模块来源、稳定错误码和面向上层的说明。

## 4. 功能范围

做：本地日志、错误封装、UNC 校验/解析、配置读写、只读延迟探测、纯逻辑单元测试。

不做：Ribbon/设置窗体、会话状态、共享锁、工作簿读写、共享审计日志、自动寻找或改写业务文件。不得写入映射盘符作为配置，不得为探测在共享上创建临时文件。

## 5. 测试方案

- `test/MdlInfra/`：UNC 有效/无效形式、盘符拒绝写入配置、UNC 解析失败、配置缺失/损坏、日志目录不可写、错误内容可供上层稳定判断。
- 路径和配置测试使用内存或本机临时目录；延迟探测的共享路径测试单独标记环境前提，且只测已有目标。
- 覆盖空字符串、尾部分隔符、大小写、权限不足和 I/O 失败；验证日志失败不会覆盖原始业务错误。
- 不需要 Excel，不启动 Excel/VBA 编译进程。

## 6. 风险与未决问题

- **日志保存位置与保留周期未由总计划确定。** 选项：本机 `%LOCALAPPDATA%` 滚动日志，或只保留短期诊断日志。暂按本机应用目录、限制体积和不记录业务单元格内容处理；需确认保留周期及是否需要用户可见的清理入口。
- **盘符到 UNC 的解析依赖 Windows 网络映射状态。** 暂定解析失败即报错，不猜测服务器路径；配置入口仍只接受 UNC。需在目标 Windows 环境验证解析 API 的可用性和权限行为。
- 延迟探测只能说明某次文件系统访问耗时，不能保证随后的提交速度；界面不得把它当作可用性承诺。
