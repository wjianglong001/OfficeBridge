# Windows 分发与 WPS 原生入口

## 交付范围

当前版本 0.12.8 提供 WPS 原生入口包、独立 exe 和 npm 三个入口。独立 `OfficeBridge.exe` 内置 Node.js、桥接代码和 WPS 静态资源，无需安装 Node.js、npm 或下载源码。开发者继续使用 `npm run dev`，也可 `npm run build` 后 `npm start`。

这些入口使用同一份服务和端口（默认 4318）。exe 默认打开 WPS 安装页面，窗口提示保持运行；关闭窗口或 Ctrl+C 停止服务。加载项安装仍由用户在页面点击完成，不自动修改 WPS 配置。此版是带控制台的便携程序，不是安装向导或托盘程序。

重复启动时，只有健康检查确认该端口已运行兼容的 OfficeBridge 服务，才复用现有服务并退出新进程；其他端口冲突清楚报错，不停止已有进程，不自动换端口，避免已安装加载项地址失效。

## WPS 原生入口包

产物为 `release/wps-native/OfficeBridge-WPS-0.12.8.zip`，构建目录为 `release/wps-native/0.12.8/OfficeBridge/`：

```text
OfficeBridge/
  OfficeBridge.xll
  OfficeBridgeService.exe
  THIRD-PARTY-NOTICES.txt
  使用说明.txt
```

用户先完整解压，再在 WPS 表格的 **工具 → 加载项 → 浏览** 选择 `OfficeBridge.xll`，无需单独打开 exe 或脚本；同目录后台必须保留。如 WPS 提示信任并安装，由用户确认。该原生包匹配本次验证的 Windows x64 与 32 位 WPS 表格。

XLL 在 WPS 内运行本机引导线程，监听 `127.0.0.1:43192`。完整后台单独运行于 `127.0.0.1:4318`：启动前核验同目录固定文件的 SHA-256，用普通 `CreateProcessW` 创建 `OfficeBridgeService.exe`，不经 `OAAssist.ShellExecute`、命令解释器或其他进程代理。参数固定，不从表格或 Agent 回复接收可执行路径。健康检查确认服务身份和协议后才打开完整对话侧栏，表格仍由 WPS 执行器操作。

兼容后台已运行时可复用；端口被其他服务占用、文件缺失/不匹配或系统拒绝创建进程时，显示真实阶段和错误码，不自动重试。停止等待不代表已经启动的独立后台会退出，关闭 WPS 也不保证退出；使用期间保留整个版本目录。

原生包后台以 Windows GUI 子系统运行，内嵌与独立 exe 相同的完整桥接功能和资源。构建只改变副本的 PE 子系统字段，并以修改后的文件计算哈希；不覆盖原独立 exe。每个版本输出到独立目录，不覆盖当前加载的旧 XLL 或运行中的后台。新版本目录隔离不等于会话或服务自动升级，已有服务是否复用仍按健康检查判断。

## 打包约定

使用 Node.js 官方 Single Executable Application（SEA）、esbuild 和 postject。在 Windows x64、Node.js 24 LTS 上运行 `npm ci`、`npm run package:win`，生成 `release/OfficeBridge.exe` 和 SHA-256 校验文件。只打包列出的加载项资源，不能包含工作区、用户数据、Agent 配置或凭证。

旧服务仍运行时可执行 `npm run package:win -- --versioned`，输出到 `release/<package.json 版本>/`，不覆盖旧 exe；使用 `npm run test:exe -- --versioned` 验证对应版本。执行器脚本及生成的同源 Worker 必须内嵌，离开源码目录仍可使用；打包不会停止或替换旧进程。

exe 的 Agent 临时工作目录放在 `%LOCALAPPDATA%\OfficeBridge\agent-workspace`，不依赖 exe 所在目录或启动目录可写；npm 默认继续使用项目 `.officebridge/agent-workspace`。可用 `OFFICEBRIDGE_DATA_DIR` 覆盖。WPS、Codex 等外部应用仍需用户自行安装和登录。

支持 `--no-open` 禁止自动打开浏览器、`--open` 显式打开浏览器、`--port <端口>` 覆盖环境变量端口，以及 `--help`。端口更改后需在新地址重新安装加载项。

0.12.2 起豆包连接准备集成于 WPS 侧栏，不再分发 `启用豆包连接.cmd`。用户选择豆包时自动检查并在未运行时启动；需要重启正在运行的豆包时由侧栏确认，再自动重启和连接。单文件 exe 与 npm 入口共用同一实现，不依赖额外脚本或修改客户端配置。详细行为见 [豆包适配](doubao.md)。

生成完整原生入口包：

```powershell
npm run package:wps
npm run test:wps-package
```

`package:wps` 先生成完整 exe，再调用 `scripts/package-wps-native.mjs`。直接调用该脚本要求已有 `release/OfficeBridge.exe`；编译器可用 `--zig <zig.exe>` 指定，默认查找项目内 Zig 0.13 工具目录。ZIP 只打包上述四个文件，不扫入输出目录中的其他文件；后台及 cJSON 的许可证随包提供并内嵌 XLL。构建不加载 WPS、重启已有服务或发送模型请求。

## 验收

原生最小验证的侧栏和真实选区读取已由用户确认，历史过程见 [验证记录](../experiments/xll-bootstrap/RESULTS.md)。这不代表完整后台、Agent 和实际写入已经通过。

- 独立 exe / GUI 后台检查覆盖项目外中文路径、无外部 Node.js、内嵌资源和 API 认证；不请求真实模型。
- 原生 API 的独立宿主和本机替身检查与 WPS 实机检查分开报告。
- 当前排查中直接启动的 4318 后台属于开发诊断，不能算作 WPS 自动启动成功。复用已运行的兼容服务同样不能代替冷启动验收。
- 完整包仍需实际从 WPS 加载、确认安装、启动后台、打开完整侧栏，以及使用真实 Agent 预览/应用的验收。
- 打包和结构检查不等于实际处理结果验证。本构建未配置代码签名；系统拒绝时保留错误，不承诺签名或换入口即可解决。

参考：[Node.js SEA](https://nodejs.org/download/release/v24.13.1/docs/api/single-executable-applications.html)。
