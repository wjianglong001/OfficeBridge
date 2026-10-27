# 架构与初始化范围

## 目标

用户在 WPS 选择单元格、行、列、区域、工作表或工作簿，直接输入指令，由所选本地 Agent 分析或生成修改方案，再由宿主工具应用到当前工作簿。
OfficeBridge 负责宿主上下文、通信与操作执行；Agent 负责理解、规划和模型调用。

这是混合型系统：桥接和表格操作使用确定性协议，开放式推理委托给外部 Agent。
初始化先验证通信和上下文契约，不引入自建 Agent 循环、数据库或云端后端。

## 分层

```text
WPS 加载项 <--> 本地桥接程序 <--> Agent 适配器 <--> 用户本地 Agent
                    ^                                 |
                    +--------- 表格工具调用 -----------+
```

- `apps/wps-addin`：工具栏、侧栏、范围元信息与完整快照、发送/取消、方案预览、WPS 写入/回读与最近一次恢复。
- `apps/bridge`：HTTP 健康检查、Agent 状态、加载项资源和官方模板生成的安装页。
- `packages/protocol`：选区、工作簿范围、宿主身份、修改方案及边界校验；同一份校验逻辑构建为浏览器脚本。
- `packages/agents`：Codex App Server / WorkBuddy ACP 适配器、stdio RPC、私有表格 MCP，以及读取/统计/检查/生成方案的共用工具。
- Codex 使用动态工具，WorkBuddy 使用私有 MCP HTTP 接口调用相同的表格工具。接口只在本机临时端口监听，每个会话独立令牌，只允许访问正在处理的当前快照。

## 当前 API

默认监听 `http://127.0.0.1:4318`，环境变量 `OFFICEBRIDGE_PORT` 可以修改端口。

| 路径                              | 方法 | 用途                                              |
| --------------------------------- | ---- | ------------------------------------------------- |
| `/api/v1/health`                  | GET  | 进程健康和协议版本                                |
| `/api/v1/agents`                  | GET  | Codex / WorkBuddy 接口可用性及其他 Agent 计划状态 |
| `/wps/taskpane.html`              | GET  | 加载项面板开发预览                                |
| `/wps/ribbon.xml`、`/wps/main.js` | GET  | WPS 加载项开发资源                                |

另外提供 `/wps/install.html` 安装页面和 `/wps/index.html` 加载引导页，均为 GET 请求。

新增 `/api/v1/session`（GET）提供随机内存令牌；`/api/v1/analyze`（POST）接收分析要求和完整选区，返回 NDJSON 事件。
新版 `/api/v1/run`（POST）接收 `{ conversationId, prompt, workbook, operationFeedback? }`，提供多范围工具并返回待应用方案。`/api/v1/conversations`（POST）接收 `{ id?, agentId? }`，选择 Codex 或 WorkBuddy 创建/替换原生会话；`/api/v1/conversations/close`（POST）接收 `{ id }` 释放所属 Agent 的会话。服务按会话记录所属适配器，发送及关闭均按绑定路由，不接受请求自行更换运行中的 Agent。插件不管理或补传聊天历史。旧分析端点保持 Codex 一次性只读兼容。所有 POST 端点共用同源/令牌校验；任务和创建会话共用并发限制。表格写入由 WPS 侧栏独立核对后执行。
服务拒绝其他 Origin、非本机 Host，写请求另外要求同源 Origin 和随机令牌；其余路由只允许 GET。会话只保留在内存，关闭时清理，断连报错而不自动更换。
安装页的 CSP 单独允许官方模板的内联脚本及对 WPS 本机中继端口 58890 的访问，其他页面与 API 不放宽。
开发工具依赖 `wpsjs@2.2.3`，只在生成安装资源时使用；桥接运行时不加载它。

## 下一阶段

1. WPS 安装、侧栏与原版选区读取已由用户实机确认；新增范围/写入/恢复需要实机验收。
2. 当前使用不可变范围快照和阶段性方案；方案绑定侧栏 connectionId、snapshotId、工作簿完整名称及目标原值，应用后回读确认。
3. 首批核心操作完成后，再推进大表、长会话管理、跨工作簿与原生图表。
4. 核心体验稳定后复用契约接入其他 Agent/MCP，逐一核实调用方式。

范围、方案和失败恢复的完整约定见 [表格操作规格](table-operations.md)。

当前 Codex 会话、数据限制和测试边界见 [Codex 接入说明](codex-analysis.md)。

后续实机验收覆盖：切换工作簿、选区变化、超大区域、Agent 取消/超时、表格写入失败，以及恢复行为。
保存脱敏错误和耗时用于定位问题；不记录原始单元格数据和 Agent 凭证。

## 参考

- [WPS 加载项开发说明](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-addin-development/wps-addin-development-instructions)
- [WPS 任务窗格](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/TaskPane/task-pane-overview)
- [Codex App Server](https://learn.chatgpt.com/docs/app-server)
- [Claude Code 程序化调用](https://code.claude.com/docs/en/headless)
