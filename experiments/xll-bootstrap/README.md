# 原生 XLL 最小验证

当前验证包：`release/xll-native/OfficeBridge-v6.xll`。WPS 已真实加载 v5 后台并可从菜单打开侧栏；现场诊断确认部署接口返回 HTTP 406，但办公桥自身 5.0.0 记录已启用。v6 先核对现有注册、匹配则复用；自动打开与测试选区读取仍待用户重测。下文保留各阶段原始记录。

验证目标：在 Windows WPS 的加载项管理中选择一个 XLL，由它释放内嵌后台、在独立端口启动后台，并通过 WPS 自带部署/调用接口打开可读取临时工作簿选区的验证侧栏。

这是实验包，不替代正式版安装方式。只支持本机匹配的 Windows WPS 位数，尚未完成升级、签名及完整卸载产品流程。不调用 Agent；侧栏只允许读取 `OfficeBridge-XLL-Probe.xlsx`，只读选区，数据只显示在侧栏，后台只接收计数与范围。

只通过 WPS 自带部署接口注册 `OfficeBridgeXllProbe`，不编辑 oem.ini、publish.xml 或其他加载项。加载和信任提示遵循 WPS 自身行为。实际验证结果应单独记录，不能把构建通过视为 WPS 验收。

当前实测结论：**本机未通过单文件安装验收**。WPS 可以选择并加载原生 XLL，但内嵌后台写出时 `WriteFile` 返回错误 5（访问被拒绝），未启动后台、未打开侧栏。独立进程中相同资源释放及 SHA-256 校验通过。不能把此实验包当成已完成的安装方案，详见 [验证记录](RESULTS.md)。

开发者构建（Windows x64、Node.js 24、已安装项目依赖、Zig 0.13）：

```text
npm run package:win
node experiments/xll-bootstrap/build.mjs --zig <zig.exe>
node experiments/xll-bootstrap/verify.mjs
npx tsx experiments/xll-bootstrap/verify-selection.ts
```

产物：`release/xll-probe/OfficeBridgeProbe.xll`，匹配本机 32 位 WPS，内含 x64 后台运行时，用户无需 Node.js/npm。默认端口 43191，与正式服务分离。构建先准备正式发行包，是为了带上相同 Node.js 运行时的完整第三方许可。校验文件和许可随包提供，许可也嵌入 XLL 资源。

手动重现：WPS 中打开临时文件 `OfficeBridge-XLL-Probe.xlsx`，填入少量人工测试数据并框选；进入“工具 → 加载项 → 浏览”，选择 XLL，确认勾选并点击确定。预期成功路径是释放后台、调用 WPS 本机部署/激活接口、自动打开只读验证侧栏。失败阶段会显示明确错误，不改权限、不提升进程、不重试、不修改其他加载项。当前自动化验收只覆盖独立后台和人工宿主替身，不代表真实 WPS 侧栏已通过。

`verify.mjs` 启动独立后台时使用 `--verify-only`：禁用 WPS 注册与激活，结束后停止自己启动的进程。`verify-selection.ts` 使用明确的人工宿主替身，复用真实执行器和读取逻辑，验证 B2:C3 数据、仅回报元信息、拒绝其他文件。

`extract-control.c` 是开发者对照程序，必须单独编译为 x86 控制台程序：读取 XLL 的内嵌资源并校验，写入专用 `XllProbeControl` 目录，绝不预填 XLL 实际使用的 `XllProbe` 目录，也不启动后台。示例：

```text
<zig.exe> cc -target x86-windows-gnu -O2 -municode -I.officebridge/xll-probe experiments/xll-bootstrap/extract-control.c -lwinhttp -ladvapi32 -luser32 -o .officebridge/xll-probe/extract-control.exe
.officebridge/xll-probe/extract-control.exe <OfficeBridgeProbe.xll绝对路径>
```

实验限制：未签名、只生成 x86 宿主版本、需要 x64 Windows；WPS 本机部署服务不可用时明确失败；加载期间固定该 DLL，停用后模块仍可能驻留至退出 WPS。后台正常启动后的退出与卸载、包更新、跨版本兼容性尚未验收，不用于正式分发。

## 原生后台第二阶段验证（v4，待 WPS 实机验收）

用户在 v3 中实测 `create_process` 返回错误 5；调用没有创建成功，提示中的 259 是初始化的 `STILL_ACTIVE`，不能当作真实子进程状态。Windows 错误 5 只说明访问被拒绝，尚未确定由 WPS 自身策略还是其他安全机制返回。

v4 移除生产启动路径中的外部进程创建，不再使用系统 DLL 加载进程。`xlAutoOpen` 在 WPS 内创建专用 HTTP 工作线程并立即返回；监听成功后由另一工作线程调用 WPS 本机加载项部署/激活接口。保留主线程可响应，用户无需启动 exe/npm。当前产物：`release/xll-native/OfficeBridge-v4.xll`。

这仍是原生安装和只读选区验证：只允许 `OfficeBridge-XLL-Probe.xlsx` 的小选区，没有接入完整 Node.js Agent 桥接，也不是独立后台进程。不要把线程模式描述为独立进程隔离或完整 Agent 安装方案。

服务线程持有非永久的 DLL 引用，结束时原子释放引用并退出线程。取消勾选/关闭加载项只发停止标记，不在 WPS 主线程等待；请求带总截止时间、接收超时和停止检查，监听关闭后后台等待正在结束的激活线程。若正在停止时又启用，明确返回繁忙。出现错误提示时，对话框未关闭前对应线程可能继续驻留。

诊断仅含阶段、错误码和进程号，保存在 `%LOCALAPPDATA%\OfficeBridge\XllNativeDiagnostics\startup-<PID>.txt`，不记录表格、路径、凭证、原生请求或对话。未修改任何安全设置、权限或 WPS 安装文件。

开发者运行：

```text
node experiments/xll-bootstrap/build-native.mjs --zig <zig.exe>
node experiments/xll-bootstrap/verify-native.mjs
```

构建还生成仅用于开发验收的 x86 `native-test-host.exe`，它不随 XLL 分发。测试宿主加载实际 XLL，并调用与生产相同的工作线程启动逻辑，显式禁用 WPS 注册和激活。检查同进程 PID、页面/Worker 资源、访问控制、只接收有界元信息及停止；不调用 Agent。此测试不能代替 WPS 实机验收。

用户重测：在 WPS 普通“加载项”中取消旧测试入口，浏览选择 `OfficeBridge-v4.xll`，勾选并确定。使用临时测试表框选 B2:C3；验收自动出现侧栏、显示当前范围和四个测试值，再切换 C2:C3 点击“读取当前选区”。失败反馈新版提示中的阶段和错误码。

## v5：等待 WPS 首次信任确认

`xlAutoOpen`、本机 HTTP 线程和单独激活线程沿用 v4。部署通过 Windows 自带 `WinHttp.WinHttpRequest.5.1` 进程内组件异步发送一次；激活线程每 100 毫秒检查同一个请求的完成状态、停止标志和 120 秒总期限。轮询不是重发。用户关闭/停用加载项或期限到达时中止本次请求；无成功响应不调用侧栏激活，不自动重试。激活本身有独立 30 秒上限。所有 COM 创建、调用、释放和反初始化都在激活线程，不向 XLL 注册 WinHTTP 回调。

只接受 WPS 官方发布器相同的 HTTP 200 且正文为 `OK` 或空；拒绝内嵌 NUL、超长正文、非 200 和重定向，不使用代理或自动登录。只限制接收结果处理长度；系统 COM 组件会缓冲响应，不能称为流式字节上限。取消不能撤销已经发给 WPS 的安装请求，WPS 的信任窗口仍由用户确认/取消；本包在取消后不再请求打开侧栏。

测试使用独立原生宿主和临时本机假服务，不连接 WPS 部署端口、不代点信任、不读取真实文件、不调用 Agent：

```text
node experiments/xll-bootstrap/verify-native.mjs
node experiments/xll-bootstrap/verify-native-relay.mjs
```

假服务覆盖超过 10 秒才批准、空成功正文、拒绝、总期限、等待中停用、迟到批准、HTTP 错误、断开、重定向、内嵌 NUL 与超长正文。测试成功不能代替 WPS 中点击信任后自动打开侧栏和读取选区的验收。

用户重测：关闭旧错误提示，在 WPS 普通“工具 → 加载项”取消旧测试包并确认。打开 `OfficeBridge-XLL-Probe.xlsx`、框选 B2:C3，再浏览选择 `OfficeBridge-v5.xll`、勾选并确定。出现 `OfficeBridgeXllProbe` 和 `http://127.0.0.1:43191/probe/` 的信任窗口时，由用户点击“信任并安装”；预期确认后自动出现侧栏。超时/失败时本轮终止，需要用户重新启用，不会后台重试。

参考 Windows 官方 [WaitForResponse](https://learn.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-waitforresponse)、[Abort](https://learn.microsoft.com/en-us/windows/win32/winhttp/iwinhttprequest-abort)。

## v6：复用已启用的验证加载项

v5 用户截图已证明菜单和侧栏可用；只读现场健康信息显示后台正常、侧栏页面请求 1 次、读取成功 0 次。日志中的真实部署状态是 HTTP 406，而 v5 将任意非 200 一律映射为 12152，因此不能把该值直接解释为响应解析错误，也不能断言实际侧栏激活失败。只读 `/publishlist` 返回办公桥自身记录：在线 et 加载项、`enable_dev`、地址正确、版本 5.0.0；独立单次打开自身侧栏请求返回 200 OK。尚未确定这一次重复部署为什么被 WPS 返回 406，不把“已安装”当作 406 的普遍含义。

v6 在部署前先只读查询 `/publishlist`（5 秒上限），完整解析并严格核对唯一同名记录的类型、在线标志、启用状态、地址和版本。全部匹配时跳过部署，仅发送一次打开请求；没有同名记录时才发起一次原有安装流程。记录禁用、未知、冲突、重复、查询失败或解析不完整均停止，不当作缺失覆盖安装。明确安装拒绝/HTTP 失败/超时之后仍不打开、不重试。提示保留实际 HTTP 状态，并区分安装未确认与打开请求未确认。

本版只变更原生引导，JS 侧栏资源保持不变，故 JS 包版本仍为 5.0.0，XLL 验证协议版本为 6。不能仅改 JS 包版本号迫使用户重新安装。当前启用状态只接受本机及官方 wpsjs 代码明确使用的字符串 `enable_dev`，其他状态需单独验收。

JSON 解析使用固定上游 cJSON v1.7.19，深度限制 32、响应处理预算 64 KiB、拒绝重复字段及解码 NUL，许可内嵌并随包提供；来源与原始哈希见 `vendor/cJSON/SOURCE.md`。不输出、存储或修改其他加载项记录。健康信息中 `registration=2/deployment=3` 表示复用现有记录；`invocation=1` 只表示打开请求返回成功，不能当作侧栏真实可见或选区已读；读取成功仍由侧栏执行器确认。

隔离测试改用临时端口，可在真实 WPS 验证服务运行时执行，不停止、不复用用户的服务。部署假服务覆盖现有记录复用、缺失、冲突、406、打开失败和原有取消/超时等边界。另有 `native-registry-test.c` 的纯合成解析检查，不操作 WPS 或 Agent。

用户重测：关闭旧提示，在 WPS 普通加载项中取消勾选 v5 并确认；保持 JS“办公桥验证”加载项现有安装，不需要卸载。在 `OfficeBridge-XLL-Probe.xlsx` 框选 B2:C3，再浏览加载 `OfficeBridge-v6.xll`。预期复用现有安装并自动打开侧栏，点击“读取当前选区”应显示四个测试值。已打开的旧侧栏会保留上次提示；切换到测试工作簿后再次点击读取，不把旧提示当作新读取结果。当前人工业务表只会显示切换提示，不读取其数据。
