# Native Doubao protocol subset

Source: https://github.com/Fullstop000/doubao-cli

Pinned revision: `b790ea2d127922902f7dc461b4169c28a2438f26`.

MIT License, Copyright (c) 2026 Fullstop000. See LICENSE.

Unmodified upstream modules: config.mjs, permissions.mjs. OfficeBridge patches protocol.mjs to use end-event brief text only when no streamed answer exists, preventing truncated summaries from overwriting complete tool requests. app-modules.mjs and protocol.mjs recognize two observed desktop chunk registries for the upstream sandbox reconciliation path. turns.mjs uses authenticated HTTP for task cancellation. OfficeBridge's app.mjs replaces installation/profile discovery with explicit in-memory context. Declaration files are OfficeBridge additions. Formatting of upstream modules is preserved.

Only native send, exact-turn reply reading, and cancellation/status functions are invoked. CLI startup, updates, authentication, on-disk receipt storage, local connectors, filesystem sandbox provisioning, and profile history enumeration are not invoked. The driver fixes cloud task context and stages bounded table operations in OfficeBridge. Reply reads and native task cancellation inspect only the adapter-owned conversation and its child task state; messages are never replayed into prompts.

OfficeBridge 0.12.10 cloud sessions no longer require a webpack registry for connection, send, exact-turn reading or cancellation. Task termination shares the authenticated HTTP path used by the upstream web implementation. Runtime identity reads the native runtimeStorage key, following upstream revision `62dd0ac1a0d1435f3978ce78c33b0202c3e6c8bf`, but never initializes persistent storage during probes. Registry-based sandbox reconciliation remains outside the OfficeBridge cloud path.

The protocol is private and compatibility is not guaranteed. The Windows gateway's limited test results do not establish compatibility of this adapter or this cloud-task configuration. Revalidate on the user's client before claiming live support.
