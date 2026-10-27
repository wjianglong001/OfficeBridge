# 已采用的来源

核验日期：2026-10-20。只采用产品方、工具提供方或仓库作者提供的一手资料。作者的实测声明不等于本机验收；main 分支会变化，真正实施时应固定版本。

[1] 飞书官网：豆包工作的入口、自定义连接器与权限边界。https://www.feishu.cn/content/article/7677519271848610746

[2] 草料开放平台：豆包工作自定义 HTTP/Streamable HTTP 连接器配置。https://cli.im/open-api/agent/open-platform-mcp/doubao-work.html

[3] ChassyChan：Windows Doubao Desktop Gateway README。实验项目，作者称在单台 Windows、豆包 2.31.4 上做过有限实测；明确列出未验证能力。https://github.com/ChassyChan/codex-doubao-desktop-gateway

[4] 同项目 server.mjs：创建任务、指定会话续聊、等待 run、读取消息等接口的实际包装代码。https://raw.githubusercontent.com/ChassyChan/codex-doubao-desktop-gateway/main/src/server.mjs

[5] 同项目 install-windows.ps1：固定上游 revision 并应用补丁。仅阅读，未运行。https://raw.githubusercontent.com/ChassyChan/codex-doubao-desktop-gateway/main/scripts/install-windows.ps1

[6] 同项目 Windows patch：数据目录、进程检测、退出和启动适配。仅阅读，未应用；包含强制结束进程的回退，不应直接照搬。https://raw.githubusercontent.com/ChassyChan/codex-doubao-desktop-gateway/main/patches/doubao-cli-windows.patch

[7] Fullstop000/doubao-cli README：原生会话标识、按 run 等待、停止、MCP 参数和平台限制。上游声明桌面 macOS 支持，不能单独作为 Windows 支持证明。https://github.com/Fullstop000/doubao-cli

[8] 豆包开放平台首页：Skill、Plugin、OpenAPI 的方向和内测状态。仅能支持平台介绍，不足以证明公开的桌面账号会话调用协议。https://open.doubao.com/

[9] MediaClaw 自家接入说明：通过自定义 STDIO 连接器连接本地插件、人工配对的产品案例；不用于证明 OfficeBridge 已兼容。https://mediaclaw.app/features/doubao-integration

[10] Patsnap 自家接入指南：HTTP 连接器、工具发现、本地工作任务选择及其实测边界。明确其截图为 macOS，不能作为 Windows 端到端调用验收。https://open.zhihuiya.com/devportal/integrations/doubao

[11] doubao-cli protocol.mjs：在已认证的客户端渲染进程内提交协议请求的实现。非官方协议，未在本机执行。https://raw.githubusercontent.com/Fullstop000/doubao-cli/main/src/protocol.mjs

## 本地代码证据

- `packages/agents/src/table-mcp.ts`：随机端口与令牌；tools/call 仅允许绑定的当前任务；解绑或取消后拒绝调用。
- `packages/agents/src/workbuddy.ts`：原生会话与工具网关的创建、逐轮 bind/clear、正常完成后发布方案。
- `packages/agents/src/index.ts`：当前 AgentAdapter 接口，尚无豆包实现。
- `AGENTS.md`、`docs/table-operations.md`：历史由 Agent 管理、修改方案与宿主应用分开、禁止默认消耗模型额度。

## 未采用的线索

搜索结果中的新闻、聚合站和非官方使用手册仅用于找到上述来源，没有承担核心结论。官方 OpenAPI 详细会话接口、Windows 2.32.6 内部 MCP 的兼容性均未获可确认的证据。
