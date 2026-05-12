# WindowsUseSDK

Windows counterpart to `MacosUseSDK`. It keeps the same operating model:

1. Use an OS-native semantic tree first.
2. Traverse the current state before every decision.
3. Prefer direct element actions.
4. Fall back to keyboard/mouse input and OCR only when the semantic tree is sparse.

The native entry point is `WindowsUseSDK.ps1`. It uses Windows UI Automation through
`.NET` and Win32 input APIs available on standard Windows installations.

## Tools

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 list-apps --query Calculator
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 open Calculator
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 traverse --hwnd 123456 --visible-only --no-activate
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 traverse --target 飞书 --view raw --visible-only --no-activate
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 elements --hwnd 123456 --limit 80
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 elements --target 飞书 --view raw --limit 120
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 probe --target 飞书 --x 138 --y 35
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 observe Calculator --visible-only
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 uia click 123456 "uia:123456:root.children[0]"
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 uia click 123456 "uia:123456:raw:root.children[0]"
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 uia activate 123456 "uia:123456:root.children[0]"
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 input --hwnd 123456 keypress ctrl+f
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 input --hwnd 123456 streamtext "hello"
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 ocr --hwnd 123456 --uia-path "uia:123456:root.children[0]"
powershell -NoProfile -ExecutionPolicy Bypass -File .\WindowsUseSDK.ps1 ocr --hwnd 123456 --rect 100,200,360,80
```

`traverse` returns elements with stable `uiaPath`/`uia_path` values like
`uia:<hwnd>:root.children[3].children[1]`. The workflow treats these as the
Windows UI Automation paths.

`elements` is the preferred first inspection command for app automation. It
filters the UIA tree into actionable controls with `uia_path`, screen frame,
center point, and action hints. Use full `traverse` only when the compact list is
not enough.

`traverse`, `elements`, and `observe` accept `--view control|raw|content`.
ControlView is the default and matches normal Windows UIA behavior. RawView is
important for Electron/Chromium apps such as Feishu/Lark, Slack, Teams, and
embedded browser surfaces where ControlView exposes only top-level panes. Raw
paths include the view in the `uia_path`, for example
`uia:<hwnd>:raw:root.children[3]`, and can still be used by `uia` and `ocr`.

`probe` moves the pointer to a screen point, asks UIA for the element under the
pointer with `AutomationElement.FromPoint`, resolves an `uia_path` when possible,
and OCRs a tight local crop around that element or point. Use it for hover
tooltips and sparse Electron/Chromium UI where a whole tree traversal is poor
but the candidate center is known.

`list-apps`/`open` enumerate real visible top-level windows, not only
`Process.MainWindowHandle`. When apps show login, QR, splash, or chat windows in
the same process, the SDK ranks candidates by title match, foreground status,
main-window status, and window size while penalizing login-like titles.
Chinese/English aliases are normalized for common apps such as WeChat/Weixin,
WXWork, Feishu/Lark, Calculator, and Notepad.

When WeChat only exposes a small QR login window but the real chat window is
hidden in the tray, `open 微信` tries WeChat's Ctrl+Alt+W global restore shortcut
once, re-ranks windows, and returns the larger non-login `hwnd` with
`recovered_from_authentication_window: true` when recovery succeeds. Only treat
`authentication_required: true` as final after this recovery path fails.

For click-like actions, `uia click` resolves the UIA element and clicks the
center of its bounding rectangle. `select` still uses UIA selection semantics,
but list-like controls also get a center-click fallback because many modern apps
report successful selection without navigating.

For text, `writetext`, `streamtext`, and `typetext` send Unicode characters with
Win32 `SendInput`, so normal typing is visible and does not borrow the clipboard.
Use `pastetext`/`clipboardtext` when bulk paste is explicitly preferred.

OCR supports targeted capture by full window, absolute screen rect, or UIA
element `uia_path`. Targeted OCR returns recognized text plus `screen_frame` and
`center` coordinates, so OCR remains a coordinate-aware fallback instead of a
vision-first screenshot workflow.

Feishu/Lark notes: start with ControlView, then probe RawView before OCR. Use
`Ctrl+K` for search when no edit field is exposed. For organization switching,
use the bottom-left organization dock, not the top-left avatar/profile card.
The top-left text can be the user display name and may stay unchanged across
organizations. Verify a switch by opening the top-left profile card after
clicking a bottom dock org icon and matching targeted OCR against the profile
subtitle/team text. Never click join/create/login account actions unless the
user explicitly asks.
When UIA remains sparse, `elements --target 飞书` returns app-profile
`VirtualRegion` hints with `frame` and `center` values for these same areas, so
the fallback remains region-aware instead of full-screen visual guessing.
Probe a region center before clicking when the target is ambiguous.

## Workflow

`workflows/llm_app_workflow.py` mirrors the macOS workflow. It opens a target app,
traverses UIA, asks the LLM for one action, executes at most one action, then
re-traverses before planning the next step.

```powershell
python .\workflows\llm_app_workflow.py "open Calculator and calculate 2+3" --target Calculator
python .\workflows\llm_app_workflow.py "open Calculator and calculate 2+3" --target Calculator --execute
python .\workflows\llm_app_workflow.py "open 飞书 and inspect the workspace switcher" --target 飞书 --uia-view auto --debug-observation
```
