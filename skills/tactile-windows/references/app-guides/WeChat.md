# WeChat

## Profile

| name | uia_view | workflow_bias | visual_planning | fixed_strategy |
| --- | --- | --- | --- | --- |
| wechat | sparse-uia-with-profile-regions | targeted-ocr | true | true |

## Match Terms

- WeChat
- 微信
- Weixin
- XinWeChat
- com.tencent.xinWeChat

## Planner Guidance

- Open WeChat first and keep the returned `hwnd` and `frame`. Skip broad app listing unless the target is ambiguous or open fails. If discovery shows a minimized/offscreen auxiliary window such as a media viewer, do not chase it after `open` recovers a visible main chat window.
- For routine one-recipient messaging, use the fast path only after the current WeChat session has a verified search-focus path, or when the intended chat is already open and the flow can skip searching: `.\bin\tactile-windows.cmd wechat-send-message --chat "<name>" --message "<text>" --require-title-match --no-draft-ocr`. It bundles search, open, compose, existing-draft replacement, one targeted recipient check, and send into one command, uses direct Win32 input by default, then focuses the compose area and presses Enter to send.
- For low-risk short messages, once the main window is visible and the search-focus path is known, run the helper directly instead of doing `dry-run`, broad `observe`, raw UIA scans, or compose OCR. Read the result JSON for `status: success`, `message_sent: true`, and `title_verification.matched: true`; ignore mojibake in PowerShell-rendered Chinese fields unless the saved UTF-8 JSON value is actually wrong.
- Do not use `wechat-send-message --draft-only` as a preflight and then manually press Enter to send. The helper still searches by pasting the contact name and pressing Enter; if the search box did not receive focus, that Enter can send the contact name to the currently open chat. If the currently open chat is already the requested recipient, the later title check can pass and hide the mistake.
- Before searching, check whether the open chat title already matches the intended recipient. If it does, skip the search token entirely, focus the compose area, replace any existing draft, type only the requested message, verify the draft when needed, and submit once. Never type the recipient name into WeChat unless the latest UIA observation or targeted probe confirms the search field has focus.
- For high-value details such as meeting links, meeting numbers, addresses, or codes, pass a single-line message or split the notification into short sequential messages. Multiline command arguments can be truncated or interpreted differently by the shell/chat input path; visual-check the sent bubble when preservation matters.
- Do not preflight with a separate dry-run or character-code probe for normal user-supplied Chinese text. PowerShell console output can render UTF-8 JSON as mojibake while the saved JSON/string value remains correct.
- Use `--draft-only` when the user asks to prepare but not send or when a high-value message needs human-visible review before submission. Use `--dry-run` to inspect computed coordinates only for debugging, `--sdk-input` if direct Win32 input is blocked by focus or policy, `--send-method button` only when Enter is known not to send in that WeChat setup, and `--keep-existing-draft` only when appending to an existing draft is intentional.
- Run `elements --target 微信` once. If it only returns sparse `Window`/`Pane` content, continue with frame-relative profile regions and targeted OCR instead of repeatedly scanning RawView.
- Convert local profile-region points to screen coordinates from the latest `frame` immediately before each click or OCR crop.
- Search/open a contact by clicking the search field, pressing `ctrl+a`, pasting the full contact name, waiting briefly, and pressing `enter` only after the current observation/probe confirms the search field is focused. In sparse WeChat UIA, a targeted probe of the profile search region that changes from the placeholder text to blank/cleared state after clicking the known search center is acceptable focus evidence for routine helper sends.
- Verify the opened chat with targeted title OCR before typing the message body. The helper normalizes OCR spaces, so `石 展 鹏` can match `石展鹏`; only skip `--require-title-match` when OCR text is mojibake or unreadable in the current environment.
- For Chinese contact names and messages, prefer `pastetext`; Unicode streaming may fail in WeChat.
- Before Send, refresh or trust only the latest `frame`, focus the compose area, and press Enter. Do not rely on the bottom-right Send button when the window is partially offscreen, scaled, or overlapped. Skip compose OCR for routine short messages because it is slower and often misses dark-theme drafts; use compose/recent-message OCR only for links, codes, addresses, or other must-preserve details.
- Treat OCR and screenshots as supporting verification, not the primary recovery path. If focus or draft state is uncertain after a failed search/open step, re-establish state through UIA/profile-region probing or stop; do not press Enter based on visual confidence alone.
- For Moments, treat the timeline/detail popup as a separate window. Re-observe the popup `hwnd` and frame before likes, comments, or visual-only controls.

## Read-And-Reply Notes

- Avoid jumping to visual inspection. Verify the active chat title first, then OCR the newest message region with frame-relative crops.
- If the latest-message OCR text is mojibake or semantically incomplete, retry a tighter crop around the latest bubble and re-read the saved JSON as UTF-8 before opening the capture image.
- Use image inspection only as the final fallback when exact message content is required to write the reply and text extraction is still ambiguous.
- When OCR JSON and PowerShell output disagree on Chinese text, suspect console/code-page mojibake first; when explicit UTF-8 readback still shows mojibake-like OCR text, treat only that dense OCR extraction as unreliable.
- Current WeChat 4.x does not expose chat message text through ControlView or RawView UIA in verified runs; both views may only return the top-level window/pane. Do not keep scanning UIA for message bubbles after this sparse state is confirmed.
- Local message stores live under `Documents\WeChat Files\<wxid>\Msg`, but message databases such as `ChatMsg.db`, `MicroMsg.db`, `FTSMsg.db`, and `MultiSearchChatMsg.db` may be encrypted or non-standard even though they use `.db` names. Treat direct SQLite reads as unavailable unless a read-only decrypted adapter is explicitly implemented and authorized.
- For exact text extraction without OCR, the safer candidate is an in-app copy path: locate the target text bubble, invoke WeChat's Copy action, read the clipboard, then restore the previous clipboard. This is not fully non-visual because locating the bubble still needs UI state, but the extracted text is exact once copied.

## Default Profile Regions

The following are local coordinates relative to the current WeChat window frame. Convert to screen points with `screenX = frame.x + localX` and `screenY = frame.y + localY`.

| id | description | local_x | local_y | width | height |
| --- | --- | --- | --- | --- | --- |
| wechat_search_center | Left-column search box center | 125 | 55 | 1 | 1 |
| wechat_title_ocr | Chat title OCR rect | 235 | 35 | 260 | 40 |
| wechat_left_results | Search/results OCR rect | 70 | 40 | 165 | 360 |
| wechat_compose_center | Compose area center | max(260,width*0.45) | height-85 | 1 | 1 |
| wechat_send_center | Send button center | width-60 | height-42 | 1 | 1 |
| wechat_draft_ocr | Compose/draft OCR rect | 220 | height-280 | width-250 | 240 |
| wechat_recent_sent_ocr | Recent sent-message OCR rect | 235 | max(90,height-590) | width-275 | 350 |

## Pitfalls

- Do not use screenshot/image inspection as a routine step for WeChat messaging. It is acceptable only for content-reading tasks after targeted OCR/probe and UTF-8 readback fail to produce trustworthy text.
- Do not type the message body into the global search box.
- Do not type the recipient/search token into the compose box. This can happen when the WeChat search click misses; if it happens, do not continue to send the actual message.
- Do not treat a successful title match as proof that the search step was safe. When the target chat was already open, a failed search can still send the contact name and the title match will still pass.
- Do not combine helper draft mode with a separate manual send action. Use either a single helper send or a manual locate-compose-submit sequence with one final Enter.
- Do not assume the first chat list row or top search result is the intended recipient unless title OCR confirms the opened chat.
- Do not reuse absolute send coordinates after restore, focus change, window move, resize, or monitor/DPI change.
- Do not let `list-apps` output for an offscreen WeChat child window override a fresh `open` result with a visible main window frame.
- Do not trust command success alone when the message contains links or codes. Confirm the recent sent-message region contains the preserved details, and send a corrective follow-up immediately if a line was dropped.
- Do not continue clicking Moments controls after one failed attempt. Re-observe or change strategy to avoid unliking or duplicate comments.
