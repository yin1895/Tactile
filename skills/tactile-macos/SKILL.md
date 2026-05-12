---
name: tactile
description: "Operate and diagnose local macOS apps through the bundled MacosUseSDK CLI. Use when Codex/Tactile needs to list or open apps, inspect Accessibility trees, run OCR, click/type/scroll, send keyboard shortcuts, execute observe-plan-act workflows, or debug UI automation permissions and app-specific flows for Lark/Feishu/WeChat/Apple Music."
---

# Tactile

Use this skill to inspect and operate local macOS applications through Accessibility, input events, screenshots, and OCR.

## Requirements

- Run on macOS with Accessibility and Screen Recording permissions granted to the host app or terminal.
- Have Swift/Xcode command line tools available for building the bundled `scripts/MacosUseSDK` tools on first use.
- Have `uv` on `PATH`, or set `TACTILE_UV` to a uv binary path.
- For LLM-driven workflows, set `TACTILE_OPENAI_API_KEY` and optionally `TACTILE_OPENAI_BASE_URL`.

## Primary CLI

Use the skill-local CLI from this directory:

```bash
bin/tactile-macos <command> ...
```

`bin/tactile-macos` creates and reuses a skill-local uv environment synchronized from `pyproject.toml`/`uv.lock`, so dependencies such as `openai` are installed from declared dependency files instead of ad hoc task-time installs.

## App Resolution Rule

When the user asks to operate an app, list matching installed apps once and choose the best installed identifier or path from that result. Do not directly operate a raw user-facing name, because display names often differ from bundle IDs, aliases, localized names, or `.app` paths.

Prefer a compact matched lookup instead of printing the full app catalog and filtering it in a second shell command. `--match` accepts a regex or literal fallback, `--compact` merges matching running processes into installed `.app` records, and `--best` prints the preferred single target. Use the returned `identifier` as the target for `open`, `observe`, or `workflow`.

For multi-app tasks, list once with a combined match expression and resolve every requested app against the same compact discovered list.

```bash
bin/tactile-macos list-apps --match '飞书|Feishu|Lark' --compact --best
bin/tactile-macos list-apps --match '微信|WeChat|飞书|Feishu|Lark' --compact
bin/tactile-macos open Calculator --json
```

## Feishu/Lark Fast Path

Feishu/Lark common operations are an explicit exception to the generic multi-step workflow rule below. For routine Feishu/Lark tasks, MUST use the dedicated fast commands first, not the generic LLM `workflow`. Only fall back to `workflow` when a required fast command does not exist, a fast command returns failure, or the task is outside the supported fast-command surface.

For routine Feishu/Lark fast commands, do not spend a turn resolving the app with `list-apps`, reading command help, or inspecting source code. The fast commands already default to the Feishu/Lark bundle and have built-in fallback targets. Resolve the app manually only when a fast command fails to open Feishu/Lark or the user explicitly asks for a non-default installation.

The fast commands avoid multi-step planner calls and use the actual exposed Feishu controls observed on macOS: `消息`, `知识问答`, `日历`, `多维表格`, `云文档`, `视频会议`, `飞行社`, `工作台`, `应用中心`, `通讯录`, `更多`, `搜索（⌘＋K）`, `创建`, and organization account buttons.

Use these commands for common Feishu operations:

```bash
bin/tactile-macos feishu-list-buttons
bin/tactile-macos feishu-open-section 消息
bin/tactile-macos feishu-open-section 日历
bin/tactile-macos feishu-search "飞书汇报" --open
bin/tactile-macos feishu-open-app "飞书汇报"
bin/tactile-macos feishu-open-chat --chat "张三"
bin/tactile-macos feishu-send-message --chat "张三" --message "收到" --send
bin/tactile-macos feishu-send-message --org "LGAI" --chat "张三" --message "收到" --send
bin/tactile-macos feishu-switch-org --name "个人用户"
bin/tactile-macos feishu-create-doc --org "个人用户" --title "健身营养学综述" --body "..." --copy-url
bin/tactile-macos feishu-create-doc --org "个人用户" --title "健身营养学综述" --body "..." --send-to "石展鹏" --send
bin/tactile-macos feishu-open-url --url "lark://..."
```

Defaults are speed-oriented: no LLM planning, no full-window OCR, no clipboard restore, no preflight help/source reads, and no post-action verification unless `--verify` is passed. For explicit user requests to send a routine message, send directly with `--send`; use `--draft-only --verify` only when the user asks to draft, preview, or handle unusually sensitive content.

Intent routing examples:

- If the user says "打开飞书", use `feishu-list-buttons`, `feishu-open-section`, or direct app open depending on the requested end state.
- If the user says "切换组织到 LGAI", use `feishu-switch-org --name "LGAI"`.
- If the user says "给石展鹏发消息", use `feishu-send-message --chat "石展鹏" --message "..." --send`.
- If the user combines organization switching and messaging, prefer one command: `feishu-send-message --org "LGAI" --chat "石展鹏" --message "..." --send`; do not use `workflow` or split draft/send unless the user asks.
- If the user asks to create a Feishu cloud document, use `feishu-create-doc` first. Feishu opens the newly created document in the default browser, so this fast command uses Feishu to trigger creation, then pastes title/body and optionally shares the copied browser URL from the browser foreground.
- If the user asks for 飞书汇报, 审批, 工作台应用, or a named app, use `feishu-open-app "<query>"` or `feishu-search "<query>" --open`.

## Core Workflow

1. Resolve the target app identifier/path with one compact matched `list-apps` call.
   If the match is ambiguous, for example personal WeChat and WeCom both match "微信", inspect the compact list and choose the intended bundle/path explicitly instead of relying on `--best`.
2. Before starting the concrete app operation flow, check `references/app-guides/` for a matching app guide. If one exists, read the relevant guide first and follow its app-specific constraints; do this before both `workflow` runs and manual `observe`/`ax`/`input` sequences.
3. For multi-step tasks, start with the end-to-end `workflow` for observation and bounded execution so one process owns observe-plan-act state and artifacts. This is the default path for AX-rich apps such as Slack, browsers, and other Electron/WebView apps. Feishu/Lark routine tasks are excluded from this default and must use the `feishu-*` fast commands first.
4. Use manual `observe`, `traverse`, `ax`, or `input` commands instead of `workflow` only for single-step diagnostics, when a reliable `ax_path` is already known, when the workflow is unavailable or blocked, or for the final verified submit action after a high-risk draft has been checked.
5. For high-risk external actions such as sending messages, posting comments, liking/reacting, payments, destructive edits, and account changes, split the task into locate, draft/open controls, verify, submit. Use a bounded workflow to navigate and draft or reveal controls; then manually verify the active target, visible context, and draft/action state from the latest observation before submitting.
6. Use the observation priority `AX > OCR > visual planner`. Local OCR runs as a text fallback even for AX-rich apps; attach screenshots to the visual planner only when AX and OCR are insufficient. Treat coordinate input as the last resort, not a default navigation method.
7. Send keyboard combos as one string, for example `cmd+f`.
8. Re-observe after actions that should change UI state, using an existing pid with no activation when inspecting transient popups. When a new popup, modal, child window, or secondary app window appears, identify the active/top `AXWindow` by title and frame before the next click or text entry.
9. If permissions are blocked, tell the user which macOS permission is missing and where to grant it.

## CLI Examples

```bash
artifact_dir=$(bin/tactile-macos artifact-dir)
bin/tactile-macos workflow "message a contact" -- --target WeChat --execute --debug-observation
bin/tactile-macos observe "open search" --target WeChat --output "$artifact_dir/wechat_observation.json"
bin/tactile-macos traverse 12345 --summary --output "$artifact_dir/elements.json"
bin/tactile-macos ax axpress 12345 'app.windows[0].children[3]'
bin/tactile-macos input keypress cmd+f
bin/tactile-macos ocr --pid 12345 --format tsv
bin/tactile-macos plan-log "$artifact_dir/<run>.json"
```

To debug AX geometry visually, add `--debug-ax-grid` to commands that know a target PID. This launches the same red overlay grid used by the Swift traversal highlight tests without polluting command stdout:

```bash
bin/tactile-macos traverse 12345 --debug-ax-grid --debug-ax-grid-duration 2
bin/tactile-macos ax --debug-ax-grid axactivate 12345 'app.windows[0].children[3]'
bin/tactile-macos workflow "inspect Lark" -- --target /Applications/Lark.app --debug-ax-grid --debug-ax-grid-duration 1.5
```

For raw input commands, pass `--debug-ax-grid-pid <pid>` because coordinate input itself does not identify an app process. The same mode can be enabled with environment variables: `TACTILE_DEBUG_AX_GRID=1` and optional `TACTILE_DEBUG_AX_GRID_DURATION=2`.

For the end-to-end observe-plan-act workflow, pass workflow-specific flags after `--`:

```bash
bin/tactile-macos workflow "calculate 2+3" -- --target Calculator --execute --mode auto
```

`--mode auto` chooses AX-rich mode for apps like Feishu/Lark/Slack/browser apps and AX-poor mode for apps like WeChat. Both modes collect AX and local OCR; AX-poor additionally uses app/profile region hints for sparse Accessibility. For unknown apps, `--capability-selection auto` asks the LLM once after the first traversal to choose `ax-rich` or `ax-poor` from the visible AX summary and an app-window screenshot.

When the app is expected to be AX-rich and the user asks for a multi-step app task, run a bounded workflow first instead of manually chaining `observe`/`ax`/`input` commands.

Use `--visual-planning on` only when AX/OCR text is insufficient and visual-only state matters, such as selected rows, unlabeled icon buttons, popovers, badges, or canvas-like content.

For long or irreversible tasks, bound the run first:

```bash
bin/tactile-macos workflow "switch org and draft a message" -- --target /Applications/Lark.app --execute --mode auto --visual-planning off --max-steps 6 --max-actions-per-step 1 --plan-output "$artifact_dir/lark-run.json"
```

Keep message sends, payments, destructive edits, and account changes behind an explicit verification step: confirm the active target and draft body from the latest observation before pressing Enter or clicking a submit/send control.

## Workflow Bailout Rules

- If `workflow` repeats the same ineffective action, such as scrolling a picker several times without approaching the target option, stop that run and switch to manual `observe`/`ocr`/`ax`/`input`. Continuing a looping workflow is riskier than taking over with fresh observations.
- Treat picker-backed controls as stateful popups. After every attempted picker change, close or commit the picker, then re-observe the closed field value before moving to the next field.
- When a workflow or visual planner gives coordinate clicks in a picker, verify the same control and target text in a fresh OCR/AX observation before reusing coordinates. Detached popups can expose different window indexes and unusual screen coordinates.
- Use `pgrep -af "tactile|MacosUse|workflow"` and `ps` only to diagnose an apparently stuck workflow. Kill only the specific workflow PIDs you started for the current task.

## Multi-Window And Popup Handling

- Treat detached popups, profile cards, media viewers, menus, and modal child windows as separate targets. Re-run `traverse`, `observe`, or OCR after they appear and use the matching window frame or region for subsequent actions.
- If a workflow cycles windows, activates an element, or reports a visual-planner coordinate, verify the foreground title/content changed as expected before acting on that result.
- Do not mix coordinates from the main app window with coordinates from a child window or popup. If the popup moves, closes, scrolls, or is partially covered, discard prior coordinates and re-observe.
- When a control is visual-only, prefer revealing a labeled menu or hover state, then OCR/observe the revealed controls before clicking. If the control remains unlabeled, use one fresh coordinate click and stop to re-plan if the expected state is not visible.

## Coordinate And Screenshot Safety

- Coordinates sent to `input click` are macOS top-left screen points. They are not screenshot pixels and not coordinates from images rendered in the chat UI.
- Never infer click coordinates from a displayed screenshot, a copied screenshot, or `step-*-visual.png`. Retina captures may be 2x pixels and planner images may be resized again.
- For coordinate clicks, use only a fresh `observe` element center, OCR JSON/TSV `screenCenter`/`screenFrame`, or, as a last resort, `visual_observation.coordinate_space.screenshot_region` from the same workflow step.
- Do not use plain coordinate guesses quoted in a planner summary unless they can be tied back to the latest observation's coordinate space and active window. Re-observe and compute the click point from the current `screenCenter`/window frame instead.
- Before a coordinate click, verify the target text/control exists in the latest observation and click the center of that target, not an estimated edge or row position.
- After one coordinate click that does not change the expected state, stop and re-plan through AX, keyboard shortcuts, OCR, or app-specific navigation. Do not repeat nearby coordinates.
- Save screenshots in the run artifact directory with unique names. Avoid reusing generic `/tmp` filenames when reporting or inspecting screenshots because stale images can look plausible.
- When multiple windows or popovers exist, identify the active `AXWindow` and use the matching pid/window frame or an explicit OCR `--region`. Do not mix screenshots or coordinates across windows.

## Text Input And Clipboard

- For text boxes, search boxes, and compose fields, focus the real AX/OCR-identified input control and prefer `bin/tactile-macos input writetext "text"` first. It simulates incremental real text input events and is usually better than AX value writes for live search, compose, and WebView/Electron fields.
- If `writetext` does not make the expected text visible or does not trigger app behavior, then use UTF-8 clipboard paste (`pbcopy` plus `input keypress cmd+v`) as the next fallback. Use direct AX value writes only when real input events and paste are unsuitable.
- In non-ASCII text paths, force UTF-8 for macOS clipboard diagnostics and helpers: `LC_ALL=en_US.UTF-8 pbcopy` and `LC_ALL=en_US.UTF-8 pbpaste`. A shell with `LC_ALL=C` can make `pbcopy` silently produce an empty clipboard for Chinese text.
- After `writetext` or paste into a compose/search field, re-observe and require the expected text to be visible before submitting or finishing.
- If the current observation already shows the intended text in the target text input, do not run another `writetext`; submit, finish, or clear and replace only if the visible text is wrong.

## App Guides

App-specific operation guidance lives under `references/app-guides/`.

- If the target app has a guide in this directory, read that guide before starting the concrete operation flow. This is mandatory for both automated `workflow` execution and manual diagnosis/control.
- Read `references/app-guides/Lark.md` before operating Feishu/Lark organization switching, contact search, or message sending.
- Read `references/app-guides/WeChat.md` before operating WeChat search, message sending, profile cards, Moments, likes, or comments.
- Read `references/app-guides/AppleMusic.md` before operating Apple Music search, album navigation, playlist selection, queue actions, or playback.
- Read `references/app-guides/TencentMeeting.md` before operating Tencent Meeting meeting scheduling, date/time selection, duration selection, or invite sharing.
- Read `references/app-guides/Jianying.md` before operating Jianying/CapCut video imports, timeline assembly, speed changes, transitions, or export.

The end-to-end workflow loads matching guides automatically and injects their planner guidance into each step, but the operator must still read the matching guide first so the initial strategy and safety checks follow app-specific guidance.

## Artifacts

`artifact-dir` resolves to a `macos-app-workflow` subdirectory under the current session directory when one is available. Workflow run logs, step observations, OCR screenshots, and redirected `/tmp` outputs are written there.

OCR JSON keeps the raw screenshot-relative pixel `frame`; when the source is `--pid` or `--region`, it also includes `screenFrame` and `screenCenter` in top-left screen coordinates suitable for coordinate input. TSV output uses `screenFrame` when available.

## Safety

- Keep actions scoped to the app and task the user requested.
- Do not type or submit sensitive data unless the user explicitly supplied it and asked for that action.
- Avoid destructive UI actions unless the user clearly requests them.
