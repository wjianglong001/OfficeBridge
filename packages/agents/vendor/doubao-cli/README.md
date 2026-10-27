# Native Doubao protocol subset

Source: https://github.com/Fullstop000/doubao-cli

Pinned revision: `b790ea2d127922902f7dc461b4169c28a2438f26`.

MIT License, Copyright (c) 2026 Fullstop000. See LICENSE.

Unmodified upstream modules: turns.mjs, config.mjs, app-modules.mjs, permissions.mjs. OfficeBridge patches protocol.mjs to use end-event brief text only when no streamed answer exists, preventing truncated summaries from overwriting complete tool requests. OfficeBridge's app.mjs replaces installation/profile discovery with explicit in-memory context. Declaration files are OfficeBridge additions. Formatting of upstream modules is preserved.

Only native send, exact-turn reply reading, and cancellation/status functions are invoked. CLI startup, updates, authentication, on-disk receipt storage, local connectors, filesystem sandbox provisioning, and profile history enumeration are not invoked. The driver fixes cloud task context and stages bounded table operations in OfficeBridge. Reply reads and native task cancellation inspect only the adapter-owned conversation and its child task state; messages are never replayed into prompts.

The protocol is private and compatibility is not guaranteed. The Windows gateway's limited test results do not establish compatibility of this adapter or this cloud-task configuration. Revalidate on the user's client before claiming live support.
