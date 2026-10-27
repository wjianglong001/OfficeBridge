# WorkBuddy 接入

本轮接入 Windows WorkBuddy 桌面安装包自带的 Agent CLI，通过 ACP stdio 创建和复用原生会话。不会将独立 CodeBuddy CLI 冒充 WorkBuddy，也不接管桌面中已有的聊天。一个 OfficeBridge 对话对应一个 WorkBuddy session，新对话更换 session，历史和压缩由 WorkBuddy 管理。

自动查找常见安装目录和 Windows 卸载注册表。`OFFICEBRIDGE_WORKBUDDY_PATH` 可指定 WorkBuddy.exe 或 WorkBuddyAI.exe。使用该安装包的 Electron Node 运行时启动同包 CLI，不要求 exe 用户安装 Node/npm；运行时须支持 `ELECTRON_RUN_AS_NODE`。登录和模型配置交给 WorkBuddy 自己读取，不复制、解析或保存凭证。

每个会话建立仅监听 127.0.0.1 的临时 MCP HTTP 服务，使用独立随机令牌，只提供现有 14 个表格工具（含行列增删、已有表改名、原表筛选及图表）。本轮工具绑定最新快照，空闲、取消和关闭后拒绝工具调用。修改仅生成方案，正常完成才发布；由用户点击应用后交给 WPS 执行。禁用内置工具、其他 MCP、hooks 和 transcript 持久化；不修改用户全局配置。停止通过 ACP session/cancel，确认当前 prompt 结束后继续原会话；无法确认或断连则报会话失效，不自动重建。

修复工具可见性时，必须区分 MCP 握手/列举成功和进入模型可调用列表。会话固定使用 cli 主 Agent，只显式暴露从共享工具清单生成的 14 个 `NoDefer(mcp__officebridge__...)` 工具，并关闭本进程工具延迟加载，不启用文件/命令工具。创建会话时等待 MCP initialized 与 tools/list 完成。每轮统计真实工具成功调用；零调用只可返回文字，完成事件明确注明未读取表格或生成方案，不能凭模型文字制造应用按钮。

桥接按会话记录所属 Agent，创建时选择 Codex 或 WorkBuddy，发送和关闭均按绑定路由。切换 Agent 创建新会话，创建失败保留原对话；不会将历史移交另一个 Agent。侧栏保持聊天主界面，Agent 选择收进右上角菜单。

验证分开报告：假 ACP/MCP 的多轮、工具边界、失败与取消测试；浏览器选择/切换和表格流程；本机真实 initialize/session/new/MCP 预检不发送模型请求。会话创建成功不等于已验证模型登录与推理；真实任务仍受用户 WorkBuddy 账号配置和网络影响。

参考：[官方 ACP 接入](https://www.workbuddy.ai/docs/zh/cli/acp)、[CLI 工具限制和内存会话参数](https://www.codebuddy.cn/docs/cli/cli-reference)、[ACP 会话](https://agentclientprotocol.com/protocol/session-setup)、[MCP HTTP 传输](https://modelcontextprotocol.io/specification/2025-03-26/basic/transports)。
