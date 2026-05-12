$ErrorActionPreference = "Stop"
try {
    $utf8 = [System.Text.UTF8Encoding]::new($false)
    [Console]::InputEncoding = $utf8
    [Console]::OutputEncoding = $utf8
    $OutputEncoding = $utf8
} catch {}

$Command = if ($args.Count -gt 0) { [string]$args[0] } else { "" }
$Rest = if ($args.Count -gt 1) { [string[]]$args[1..($args.Count - 1)] } else { @() }

try { Add-Type -AssemblyName UIAutomationClient } catch {}
try { Add-Type -AssemblyName UIAutomationTypes } catch {}
try { Add-Type -AssemblyName WindowsBase } catch {}
try { Add-Type -AssemblyName System.Windows.Forms } catch {}
try { Add-Type -AssemblyName System.Drawing } catch {}
try { Add-Type -AssemblyName System.Runtime.WindowsRuntime } catch {}

$nativeSource = @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class LuopanNativeMethods {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();

    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    public static extern IntPtr WindowFromPoint(POINT Point);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT {
        public uint type;
        public INPUTUNION U;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct INPUTUNION {
        [FieldOffset(0)]
        public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct KEYBDINPUT {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    private const uint INPUT_KEYBOARD = 1;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_UNICODE = 0x0004;

    public static string GetWindowTextRaw(IntPtr hWnd) {
        int length = GetWindowTextLength(hWnd);
        StringBuilder builder = new StringBuilder(Math.Max(length + 1, 256));
        GetWindowText(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    public static string GetClassNameRaw(IntPtr hWnd) {
        StringBuilder builder = new StringBuilder(256);
        GetClassName(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    public static int GetWindowProcessId(IntPtr hWnd) {
        uint pid;
        GetWindowThreadProcessId(hWnd, out pid);
        return (int)pid;
    }

    public static bool SendUnicodeText(string text, int delayMs) {
        if (text == null) {
            return true;
        }
        foreach (char ch in text) {
            INPUT[] inputs = new INPUT[2];
            inputs[0].type = INPUT_KEYBOARD;
            inputs[0].U.ki.wVk = 0;
            inputs[0].U.ki.wScan = ch;
            inputs[0].U.ki.dwFlags = KEYEVENTF_UNICODE;

            inputs[1].type = INPUT_KEYBOARD;
            inputs[1].U.ki.wVk = 0;
            inputs[1].U.ki.wScan = ch;
            inputs[1].U.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;

            if (SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT))) == 0) {
                return false;
            }
            if (delayMs > 0) {
                Thread.Sleep(delayMs);
            }
        }
        return true;
    }
}
"@
try { Add-Type -TypeDefinition $nativeSource -Language CSharp } catch {}
try { [void][LuopanNativeMethods]::SetProcessDPIAware() } catch {}

$script:MaxTraversalDepth = 60
$script:MaxTraversalElements = 2500
$script:MaxChildrenPerElement = 220
$script:MaxTraversalSeconds = 8.0
$script:ActionableRoles = @(
    "Button", "Calendar", "CheckBox", "ComboBox", "DataGrid", "DataItem",
    "Document", "Edit", "Hyperlink", "List", "ListItem", "Menu", "MenuBar",
    "MenuItem", "RadioButton", "ScrollBar", "Slider", "Spinner", "SplitButton",
    "Tab", "TabItem", "Table", "Tree", "TreeItem", "Window"
)

function Write-Json {
    param([Parameter(Mandatory = $true)]$Value)
    $Value | ConvertTo-Json -Depth 40
}

function Fail {
    param([string]$Message, [int]$Code = 1)
    [Console]::Error.WriteLine($Message)
    exit $Code
}

function Get-OptionValue {
    param([string[]]$Items, [string]$Name, [string]$Default = $null)
    for ($i = 0; $i -lt $Items.Count; $i++) {
        if ($Items[$i] -eq $Name -and ($i + 1) -lt $Items.Count) {
            return $Items[$i + 1]
        }
    }
    return $Default
}

function Has-Flag {
    param([string[]]$Items, [string]$Name)
    return [Array]::IndexOf($Items, $Name) -ge 0
}

function Positional-Args {
    param([string[]]$Items)
    $resultItems = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $arg = $Items[$i]
        if ($arg.StartsWith("--")) {
            if (($i + 1) -lt $Items.Count -and -not $Items[$i + 1].StartsWith("--")) {
                $i++
            }
            continue
        }
        $resultItems.Add($arg)
    }
    return $resultItems.ToArray()
}

function Normalize-Text {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return (($Value.ToString()) -replace "\s+", " ").Trim()
}

function Query-Aliases {
    param([string]$Identifier)
    $raw = Normalize-Text $Identifier
    $lower = $raw.ToLowerInvariant()
    $aliases = New-Object System.Collections.Generic.List[string]
    if ($raw) { $aliases.Add($raw) }
    if ($lower -in @("wechat", "weixin") -or $raw -eq "$([char]0x5fae)$([char]0x4fe1)") {
        $aliases.Add("WeChat")
        $aliases.Add("Weixin")
        $aliases.Add("Weixin.exe")
        $aliases.Add("$([char]0x5fae)$([char]0x4fe1)")
    }
    if ($lower -in @("wxwork", "wecom") -or $raw -eq "$([char]0x4f01)$([char]0x4e1a)$([char]0x5fae)$([char]0x4fe1)") {
        $aliases.Add("WXWork")
        $aliases.Add("WXWork.exe")
        $aliases.Add("WeCom")
        $aliases.Add("$([char]0x4f01)$([char]0x4e1a)$([char]0x5fae)$([char]0x4fe1)")
    }
    if ($lower -eq "feishu" -or $lower -eq "lark" -or $raw -eq "$([char]0x98de)$([char]0x4e66)") {
        $aliases.Add("Feishu")
        $aliases.Add("Lark")
        $aliases.Add("$([char]0x98de)$([char]0x4e66)")
        $aliases.Add("com.electron.Feishu")
    }
    return @($aliases | Where-Object { $_ } | Select-Object -Unique)
}

function Haystack-Matches-Any {
    param([string]$Haystack, [string[]]$Needles)
    $lowerHaystack = $Haystack.ToLowerInvariant()
    foreach ($needle in @($Needles)) {
        $normalizedNeedle = (Normalize-Text $needle).ToLowerInvariant()
        if ($normalizedNeedle -and $lowerHaystack.Contains($normalizedNeedle)) { return $true }
    }
    return $false
}

function Try-Value {
    param([scriptblock]$Block, [object]$Default = $null)
    try { return & $Block } catch { return $Default }
}

function Window-Rect {
    param([IntPtr]$Hwnd)
    $rect = New-Object LuopanNativeMethods+RECT
    if (-not [LuopanNativeMethods]::GetWindowRect($Hwnd, [ref]$rect)) {
        return $null
    }
    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top
    if ($width -le 0 -or $height -le 0) {
        return $null
    }
    [pscustomobject]@{
        x = $rect.Left
        y = $rect.Top
        width = $width
        height = $height
    }
}

function Activate-Window {
    param([IntPtr]$Hwnd)
    if ($Hwnd -eq [IntPtr]::Zero) { return $false }
    [void][LuopanNativeMethods]::ShowWindowAsync($Hwnd, 9)
    Start-Sleep -Milliseconds 80
    return [LuopanNativeMethods]::SetForegroundWindow($Hwnd)
}

function Window-Area {
    param($Frame)
    if ($null -eq $Frame) { return 0 }
    return [double]$Frame.width * [double]$Frame.height
}

function Top-Level-Window-Handles {
    $handles = New-Object System.Collections.Generic.List[IntPtr]
    $callback = [LuopanNativeMethods+EnumWindowsProc]{
        param([IntPtr]$Hwnd, [IntPtr]$LParam)
        if ([LuopanNativeMethods]::IsWindow($Hwnd) -and [LuopanNativeMethods]::IsWindowVisible($Hwnd)) {
            [void]$handles.Add($Hwnd)
        }
        return $true
    }
    [void][LuopanNativeMethods]::EnumWindows($callback, [IntPtr]::Zero)
    return $handles
}

function Running-Windows {
    param([string]$Query = "")
    $queryAliases = @(Query-Aliases $Query)
    $items = New-Object System.Collections.Generic.List[object]
    $foreground = [Int64][LuopanNativeMethods]::GetForegroundWindow()
    foreach ($handle in Top-Level-Window-Handles) {
        $hwnd = [Int64]$handle
        $processId = [LuopanNativeMethods]::GetWindowProcessId($handle)
        if ($processId -le 0) { continue }
        $proc = $null
        try { $proc = Get-Process -Id $processId -ErrorAction Stop } catch { continue }
        $frame = Window-Rect $handle
        if ($null -eq $frame) { continue }
        $title = Normalize-Text ([LuopanNativeMethods]::GetWindowTextRaw($handle))
        $className = Normalize-Text ([LuopanNativeMethods]::GetClassNameRaw($handle))
        $name = Normalize-Text $proc.ProcessName
        $path = Try-Value { $proc.Path }
        $exeName = Try-Value { [IO.Path]::GetFileName($proc.Path) }
        if (-not $title -and (Window-Area $frame) -lt 25000) { continue }
        if ($queryAliases.Count -gt 0) {
            $haystack = ("$name $title $exeName $path $className").ToLowerInvariant()
            if (-not (Haystack-Matches-Any $haystack $queryAliases)) { continue }
        }
        $items.Add([pscustomobject]@{
            name = $name
            title = $title
            pid = $processId
            hwnd = $hwnd
            exe_name = $exeName
            path = $path
            class_name = $className
            is_foreground = ($hwnd -eq $foreground)
            is_main_window = ($hwnd -eq [Int64]$proc.MainWindowHandle)
            source = "running_window"
            frame = $frame
            area = [math]::Round((Window-Area $frame), 0)
        })
    }
    return @($items | Sort-Object @{ Expression = { if ($_.is_foreground) { 0 } else { 1 } } }, @{ Expression = { -1 * [double]$_.area } }, name, title)
}

function Start-Apps {
    param([string]$Query = "")
    $queryAliases = @(Query-Aliases $Query)
    $items = New-Object System.Collections.Generic.List[object]
    try {
        foreach ($app in Get-StartApps) {
            $name = Normalize-Text $app.Name
            $appId = Normalize-Text $app.AppID
            if ($queryAliases.Count -gt 0) {
                $haystack = "$name $appId".ToLowerInvariant()
                if (-not (Haystack-Matches-Any $haystack $queryAliases)) { continue }
            }
            $items.Add([pscustomobject]@{
                name = $name
                app_id = $appId
                source = "start_apps"
            })
        }
    } catch {}
    return $items
}

function Window-Match-Score {
    param($Window, [string]$Identifier)
    $queries = @(Query-Aliases $Identifier)
    if ($queries.Count -eq 0) { return -1 }
    $name = (Normalize-Text $Window.name).ToLowerInvariant()
    $title = (Normalize-Text $Window.title).ToLowerInvariant()
    $exeName = (Normalize-Text $Window.exe_name).ToLowerInvariant()
    $path = (Normalize-Text $Window.path).ToLowerInvariant()
    $className = (Normalize-Text $Window.class_name).ToLowerInvariant()
    $score = -1
    foreach ($queryRaw in $queries) {
        $query = (Normalize-Text $queryRaw).ToLowerInvariant()
        if (-not $query) { continue }
        if ($title -eq $query) { $score = [Math]::Max($score, 6000) }
        if ($name -eq $query -or $exeName -eq $query -or $exeName -eq "$query.exe") { $score = [Math]::Max($score, 5600) }
        if ($title.Contains($query)) { $score = [Math]::Max($score, 4200 + $query.Length) }
        if ($name.Contains($query) -or $exeName.Contains($query) -or $className.Contains($query)) { $score = [Math]::Max($score, 3600 + $query.Length) }
        if ($path.Contains($query)) { $score = [Math]::Max($score, 3000 + $query.Length) }
    }
    if ($score -lt 0) { return -1 }

    $area = [double](Try-Value { $Window.area } 0)
    $score += [Math]::Min(450, [Math]::Round($area / 10000))
    if ($Window.is_foreground) { $score += 350 }
    if ($Window.is_main_window) { $score += 120 }
    if ($title) { $score += 60 }
    if ($area -ge 240000) { $score += 700 }

    $loginLike = @(
        "login", "log in", "sign in", "signin", "auth", "authenticate",
        "qr", "qrcode", "scan",
        "$([char]0x626b)$([char]0x7801)",
        "$([char]0x767b)$([char]0x5f55)",
        "$([char]0x767b)$([char]0x5165)",
        "$([char]0x4e8c)$([char]0x7ef4)$([char]0x7801)",
        "$([char]0x9a8c)$([char]0x8bc1)",
        "$([char]0x5b89)$([char]0x5168)$([char]0x9a8c)$([char]0x8bc1)"
    )
    $windowText = "$title $name $className"
    foreach ($keyword in $loginLike) {
        if ($windowText.Contains($keyword)) {
            $score -= 2200
            break
        }
    }
    if ($className.Contains("loginwindow") -or $className.Contains("auth") -or $className.Contains("qrcode")) {
        $score -= 2800
    }
    if ($area -gt 0 -and $area -lt 160000) { $score -= 180 }
    return $score
}

function Is-Authentication-Window {
    param($Window)
    if ($null -eq $Window) { return $false }
    $className = (Normalize-Text $Window.class_name).ToLowerInvariant()
    $title = (Normalize-Text $Window.title).ToLowerInvariant()
    $area = [double](Try-Value { $Window.area } 0)
    if ($className.Contains("loginwindow") -or $className.Contains("qrcode")) { return $true }
    if ($title.Contains("login") -or $title.Contains("sign in") -or $title.Contains("$([char]0x767b)$([char]0x5f55)") -or $title.Contains("$([char]0x626b)$([char]0x7801)")) { return $true }
    return ($area -gt 0 -and $area -lt 160000 -and ($className.Contains("weixin") -or $Window.exe_name -eq "Weixin.exe"))
}

function Is-WeChat-Identifier {
    param([string]$Identifier)
    foreach ($alias in @(Query-Aliases $Identifier)) {
        $lower = (Normalize-Text $alias).ToLowerInvariant()
        if ($lower -in @("wechat", "weixin", "weixin.exe") -or $alias -eq "$([char]0x5fae)$([char]0x4fe1)") {
            return $true
        }
    }
    return $false
}

function Try-Recover-From-Authentication-Window {
    param([string]$Identifier, $CurrentWindow)
    if (-not (Is-Authentication-Window $CurrentWindow)) { return $null }
    if (-not (Is-WeChat-Identifier $Identifier)) { return $null }

    # WeChat can keep the real chat window hidden in the tray while a small QR
    # login window remains visible. Ctrl+Alt+W is WeChat's standard global
    # show/hide shortcut; use it once, then re-rank windows.
    try {
        [System.Windows.Forms.SendKeys]::SendWait("^%w")
        Start-Sleep -Milliseconds 900
    } catch {}

    $candidate = Find-Window $Identifier
    if ($candidate -and -not (Is-Authentication-Window $candidate)) {
        $candidate | Add-Member -NotePropertyName recovered_from_authentication_window -NotePropertyValue $true -Force
        $candidate | Add-Member -NotePropertyName recovery_method -NotePropertyValue "wechat_global_hotkey_ctrl_alt_w" -Force
        return $candidate
    }
    return $null
}

function Find-Window {
    param([string]$Identifier)
    $query = Normalize-Text $Identifier
    if ($query.Length -eq 0) { return $null }
    $ranked = foreach ($window in @(Running-Windows)) {
        $score = Window-Match-Score $window $query
        if ($score -ge 0) {
            $window | Add-Member -NotePropertyName match_score -NotePropertyValue $score -Force
            $window | Add-Member -NotePropertyName authentication_window -NotePropertyValue (Is-Authentication-Window $window) -Force
            $window
        }
    }
    return @($ranked | Sort-Object @{ Expression = { -1 * [double]$_.match_score } }, @{ Expression = { -1 * [double]$_.area } }, name, title | Select-Object -First 1)
}

function Open-App {
    param([string]$Identifier, [bool]$NoActivate = $false)
    $existing = Find-Window $Identifier
    if ($existing) {
        $recovered = Try-Recover-From-Authentication-Window $Identifier $existing
        if ($recovered) { $existing = $recovered }
        if (-not $NoActivate) { [void](Activate-Window ([IntPtr]$existing.hwnd)) }
        return [pscustomobject]@{
            identifier = $Identifier
            mode = "existing_window"
            pid = $existing.pid
            hwnd = $existing.hwnd
            title = $existing.title
            exe_name = $existing.exe_name
            class_name = $existing.class_name
            match_score = $existing.match_score
            authentication_window = $existing.authentication_window
            authentication_required = [bool]$existing.authentication_window
            recovered_from_authentication_window = [bool]$existing.recovered_from_authentication_window
            recovery_method = $existing.recovery_method
            frame = $existing.frame
        }
    }

    $startApp = @(Start-Apps $Identifier) | Sort-Object @{ Expression = { if ($_.name -eq $Identifier) { 0 } else { 1 } } }, name | Select-Object -First 1
    if ($startApp) {
        $existingByLocalizedName = Find-Window $startApp.name
        if ($existingByLocalizedName) {
            $recovered = Try-Recover-From-Authentication-Window $startApp.name $existingByLocalizedName
            if ($recovered) { $existingByLocalizedName = $recovered }
            if (-not $NoActivate) { [void](Activate-Window ([IntPtr]$existingByLocalizedName.hwnd)) }
            return [pscustomobject]@{
                identifier = $Identifier
                resolved_identifier = $startApp.name
                mode = "existing_window_start_app_alias"
                pid = $existingByLocalizedName.pid
                hwnd = $existingByLocalizedName.hwnd
                title = $existingByLocalizedName.title
                exe_name = $existingByLocalizedName.exe_name
                class_name = $existingByLocalizedName.class_name
                match_score = $existingByLocalizedName.match_score
                authentication_window = $existingByLocalizedName.authentication_window
                authentication_required = [bool]$existingByLocalizedName.authentication_window
                recovered_from_authentication_window = [bool]$existingByLocalizedName.recovered_from_authentication_window
                recovery_method = $existingByLocalizedName.recovery_method
                frame = $existingByLocalizedName.frame
            }
        }
    }

    $started = $null
    if ($startApp) {
        Start-Process "explorer.exe" "shell:AppsFolder\$($startApp.app_id)"
    } elseif (Test-Path -LiteralPath $Identifier) {
        $started = Start-Process -FilePath $Identifier -PassThru
    } else {
        $started = Start-Process $Identifier -PassThru
    }

    $deadline = (Get-Date).AddSeconds(12)
    do {
        Start-Sleep -Milliseconds 250
        if ($started -and $started.Id) {
            try {
                $proc = Get-Process -Id $started.Id -ErrorAction Stop
                if ($proc.MainWindowHandle -ne 0) {
                    $hwnd = [Int64]$proc.MainWindowHandle
                    if (-not $NoActivate) { [void](Activate-Window ([IntPtr]$hwnd)) }
                    return [pscustomobject]@{
                        identifier = $Identifier
                        mode = "started_process"
                        pid = $proc.Id
                        hwnd = $hwnd
                        title = Normalize-Text $proc.MainWindowTitle
                        exe_name = Try-Value { [IO.Path]::GetFileName($proc.Path) }
                        class_name = $null
                        match_score = $null
                        authentication_window = $false
                        authentication_required = $false
                        frame = Window-Rect ([IntPtr]$hwnd)
                    }
                }
            } catch {}
        }
        $window = Find-Window $Identifier
        if (-not $window -and $startApp) {
            $window = Find-Window $startApp.name
        }
        if ($window) {
            $recovered = Try-Recover-From-Authentication-Window $Identifier $window
            if ($recovered) { $window = $recovered }
            if (-not $NoActivate) { [void](Activate-Window ([IntPtr]$window.hwnd)) }
            return [pscustomobject]@{
                identifier = $Identifier
                resolved_identifier = if ($startApp) { $startApp.name } else { $Identifier }
                mode = if ($startApp) { "start_app" } else { "shell_start" }
                pid = $window.pid
                hwnd = $window.hwnd
                title = $window.title
                exe_name = $window.exe_name
                class_name = $window.class_name
                match_score = $window.match_score
                authentication_window = $window.authentication_window
                authentication_required = [bool]$window.authentication_window
                recovered_from_authentication_window = [bool]$window.recovered_from_authentication_window
                recovery_method = $window.recovery_method
                frame = $window.frame
            }
        }
    } while ((Get-Date) -lt $deadline)

    throw "could not find a top-level window after opening '$Identifier'"
}

function Resolve-Hwnd {
    param([string]$Hwnd, [string]$ProcessId, [string]$Target)
    if ($Hwnd) { return [Int64]$Hwnd }
    if ($ProcessId) {
        $proc = Get-Process -Id ([int]$ProcessId) -ErrorAction Stop
        if ($proc.MainWindowHandle -eq 0) { throw "process $ProcessId has no MainWindowHandle" }
        return [Int64]$proc.MainWindowHandle
    }
    if ($Target) {
        $window = Find-Window $Target
        if ($window) { return [Int64]$window.hwnd }
        return [Int64](Open-App $Target $true).hwnd
    }
    throw "expected --hwnd, --pid, or --target"
}

function Get-Element-Text {
    param($Element)
    $parts = New-Object System.Collections.Generic.List[string]
    try {
        $name = Normalize-Text $Element.Current.Name
        if ($name) { $parts.Add($name) }
    } catch {}
    try {
        $help = Normalize-Text $Element.Current.HelpText
        if ($help) { $parts.Add($help) }
    } catch {}
    try {
        $pattern = $null
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
            $value = Normalize-Text ([System.Windows.Automation.ValuePattern]$pattern).Current.Value
            if ($value) { $parts.Add($value) }
        }
    } catch {}
    $seen = @{}
    $unique = foreach ($part in $parts) {
        if (-not $seen.ContainsKey($part)) {
            $seen[$part] = $true
            $part
        }
    }
    return (Normalize-Text ($unique -join " "))
}

function Element-Role {
    param($Element)
    try {
        $role = $Element.Current.ControlType.ProgrammaticName
        if ($role.StartsWith("ControlType.")) { return $role.Substring(12) }
        return $role
    } catch {
        return "Unknown"
    }
}

function Runtime-Key {
    param($Element)
    try {
        return [string]::Join(".", $Element.GetRuntimeId())
    } catch {
        try {
            return "$($Element.Current.NativeWindowHandle):$($Element.Current.AutomationId):$($Element.Current.Name)"
        } catch {
            return [Guid]::NewGuid().ToString()
        }
    }
}

function Element-Frame {
    param($Element)
    try {
        $rect = $Element.Current.BoundingRectangle
        if ($rect.IsEmpty -or $rect.Width -le 0 -or $rect.Height -le 0) { return $null }
        return [pscustomobject]@{
            x = [math]::Round($rect.X, 2)
            y = [math]::Round($rect.Y, 2)
            width = [math]::Round($rect.Width, 2)
            height = [math]::Round($rect.Height, 2)
        }
    } catch {
        return $null
    }
}

function Element-Patterns {
    param($Element)
    try {
        return @($Element.GetSupportedPatterns() | ForEach-Object {
            if ($_.ProgrammaticName.StartsWith("Pattern.")) { $_.ProgrammaticName.Substring(8) } else { $_.ProgrammaticName }
        })
    } catch {
        return @()
    }
}

function Pattern-Contains {
    param([object[]]$Patterns, [string]$Name)
    foreach ($pattern in @($Patterns)) {
        if (($pattern.ToString()).Contains($Name)) { return $true }
    }
    return $false
}

function Normalize-UiaView {
    param([string]$View)
    $normalized = (Normalize-Text $View).ToLowerInvariant()
    if (-not $normalized) { return "control" }
    switch ($normalized) {
        "control" { return "control" }
        "raw" { return "raw" }
        "content" { return "content" }
        default { throw "unsupported UIA view '$View'; expected control, raw, or content" }
    }
}

function Get-UiaTreeWalker {
    param([string]$View)
    switch (Normalize-UiaView $View) {
        "raw" { return [System.Windows.Automation.TreeWalker]::RawViewWalker }
        "content" { return [System.Windows.Automation.TreeWalker]::ContentViewWalker }
        default { return [System.Windows.Automation.TreeWalker]::ControlViewWalker }
    }
}

function Root-UiaPath {
    param([Int64]$Hwnd, [string]$View)
    $viewName = Normalize-UiaView $View
    if ($viewName -eq "control") { return "uia:${Hwnd}:root" }
    return "uia:${Hwnd}:${viewName}:root"
}

function Traverse-Window {
    param(
        [Int64]$Hwnd,
        [bool]$VisibleOnly = $true,
        [bool]$NoActivate = $true,
        [string]$View = "control"
    )

    if (-not $NoActivate) { [void](Activate-Window ([IntPtr]$Hwnd)) }
    $root = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$Hwnd)
    if ($null -eq $root) { throw "UIAutomation root not found for hwnd $Hwnd" }

    $viewName = Normalize-UiaView $View
    $walker = Get-UiaTreeWalker $viewName
    $queue = New-Object System.Collections.ArrayList
    [void]$queue.Add([pscustomobject]@{ element = $root; depth = 0; path = (Root-UiaPath $Hwnd $viewName) })
    $readIndex = 0
    $visited = @{}
    $elements = New-Object System.Collections.Generic.List[object]
    $started = Get-Date
    $stats = [ordered]@{
        count = 0
        visited_count = 0
        excluded_count = 0
        excluded_no_text = 0
        excluded_offscreen = 0
        visible_elements_count = 0
        with_text_count = 0
        without_text_count = 0
        truncated = $false
        role_counts = @{}
    }

    while ($readIndex -lt $queue.Count) {
        if ($elements.Count -ge $script:MaxTraversalElements -or ((Get-Date) - $started).TotalSeconds -gt $script:MaxTraversalSeconds) {
            $stats.truncated = $true
            break
        }
        $item = $queue[$readIndex]
        $readIndex++
        if ($item.depth -gt $script:MaxTraversalDepth) { continue }

        $element = $item.element
        $key = Runtime-Key $element
        if ($visited.ContainsKey($key)) { continue }
        $visited[$key] = $true
        $stats.visited_count++

        $role = Element-Role $element
        if (-not $stats.role_counts.ContainsKey($role)) { $stats.role_counts[$role] = 0 }
        $stats.role_counts[$role]++

        $text = Get-Element-Text $element
        $frame = Element-Frame $element
        $isOffscreen = $false
        try { $isOffscreen = [bool]$element.Current.IsOffscreen } catch {}
        $isVisible = ($null -ne $frame) -and (-not $isOffscreen)
        if ($isVisible) { $stats.visible_elements_count++ }

        $hasText = $text.Length -gt 0
        $isActionable = $script:ActionableRoles -contains $role
        $shouldCollect = ($hasText -or $isActionable) -and ((-not $VisibleOnly) -or $isVisible)

        if ($shouldCollect) {
            $patterns = @(Element-Patterns $element)
            $nativeHwnd = Try-Value { [Int64]$element.Current.NativeWindowHandle } 0
            $elements.Add([pscustomobject]@{
                role = $role
                text = if ($hasText) { $text } else { $null }
                x = if ($frame) { $frame.x } else { $null }
                y = if ($frame) { $frame.y } else { $null }
                width = if ($frame) { $frame.width } else { $null }
                height = if ($frame) { $frame.height } else { $null }
                uiaPath = $item.path
                uia_path = $item.path
                automation_id = Try-Value { $element.Current.AutomationId }
                class_name = Try-Value { $element.Current.ClassName }
                native_window_handle = $nativeHwnd
                is_enabled = Try-Value { [bool]$element.Current.IsEnabled }
                is_offscreen = $isOffscreen
                patterns = $patterns
            })
            if ($hasText) { $stats.with_text_count++ } else { $stats.without_text_count++ }
        } else {
            $stats.excluded_count++
            if (-not $hasText) { $stats.excluded_no_text++ }
            if ($VisibleOnly -and -not $isVisible) { $stats.excluded_offscreen++ }
        }

        $child = $null
        try { $child = $walker.GetFirstChild($element) } catch {}
        $childIndex = 0
        while ($null -ne $child -and $childIndex -lt $script:MaxChildrenPerElement) {
            [void]$queue.Add([pscustomobject]@{
                element = $child
                depth = $item.depth + 1
                path = "$($item.path).children[$childIndex]"
            })
            try { $child = $walker.GetNextSibling($child) } catch { $child = $null }
            $childIndex++
        }
    }

    $sorted = @($elements | Sort-Object @{ Expression = { if ($null -ne $_.y) { $_.y } else { [double]::MaxValue } } }, @{ Expression = { if ($null -ne $_.x) { $_.x } else { [double]::MaxValue } } })
    $stats.count = $sorted.Count
    $process = $null
    try {
        $rootPid = $root.Current.ProcessId
        $process = Get-Process -Id $rootPid -ErrorAction Stop
    } catch {}
    $rootClassName = Normalize-Text (Try-Value { $root.Current.ClassName } "")
    $rootFrameworkId = Normalize-Text (Try-Value { $root.Current.FrameworkId } "")
    $accessibilityHint = $null
    if ($rootClassName -eq "Chrome_WidgetWin_1" -and $stats.count -le 5) {
        $accessibilityHint = [pscustomobject]@{
            suspected_chromium_accessibility_disabled = $true
            reason = "UIA exposed only a sparse Chromium legacy window tree"
            suggested_launch_flag = "--force-renderer-accessibility"
            safe_next_step = "Use targeted OCR for the current run, or restart the Electron/Chromium app with the suggested flag before relying on UIA semantics."
        }
    }

    return [pscustomobject]@{
        app_name = if ($process) { $process.ProcessName } else { "Window($Hwnd)" }
        hwnd = $Hwnd
        pid = if ($process) { $process.Id } else { $null }
        title = if ($process) { Normalize-Text $process.MainWindowTitle } else { $null }
        class_name = $rootClassName
        framework_id = $rootFrameworkId
        view = $viewName
        accessibility_hint = $accessibilityHint
        elements = $sorted
        stats = $stats
        processing_time_seconds = "{0:N2}" -f ((Get-Date) - $started).TotalSeconds
    }
}

function Element-Action-Hints {
    param([string]$Role, [object[]]$Patterns)
    $hints = New-Object System.Collections.Generic.List[string]
    if ((Pattern-Contains $Patterns "Invoke") -or $Role -in @("Button", "Hyperlink", "MenuItem", "SplitButton")) {
        $hints.Add("click")
        $hints.Add("press")
    }
    if ((Pattern-Contains $Patterns "SelectionItem") -or $Role -in @("ListItem", "DataItem", "TreeItem", "TabItem")) {
        $hints.Add("click")
        $hints.Add("select")
    }
    if ((Pattern-Contains $Patterns "Value") -or $Role -in @("Edit", "ComboBox", "Document")) {
        $hints.Add("focus")
        $hints.Add("streamtext")
        $hints.Add("pastetext")
        if (Pattern-Contains $Patterns "Value") { $hints.Add("set_value") }
    }
    if ((Pattern-Contains $Patterns "Toggle") -or $Role -in @("CheckBox", "RadioButton")) {
        $hints.Add("click")
    }
    if (Pattern-Contains $Patterns "Scroll") {
        $hints.Add("scroll")
    }
    return @($hints | Select-Object -Unique)
}

function Get-Interactive-Elements {
    param(
        [Int64]$Hwnd,
        [bool]$VisibleOnly = $true,
        [bool]$NoActivate = $true,
        [int]$Limit = 160,
        [string]$Query = "",
        [string]$View = "control"
    )
    $tree = Traverse-Window $Hwnd $VisibleOnly $NoActivate $View
    $queryLower = (Normalize-Text $Query).ToLowerInvariant()
    $items = New-Object System.Collections.Generic.List[object]
    $index = 0
    foreach ($element in @($tree.elements)) {
        $role = [string]$element.role
        $patterns = @($element.patterns)
        $text = Normalize-Text $element.text
        $hints = @(Element-Action-Hints $role $patterns)
        if ($hints.Count -eq 0) { continue }
        if ($queryLower.Length -gt 0) {
            $haystack = ("$role $text $($element.automation_id) $($element.class_name)").ToLowerInvariant()
            if (-not $haystack.Contains($queryLower)) { continue }
        }
        $center = $null
        if ($null -ne $element.x -and $null -ne $element.y -and $null -ne $element.width -and $null -ne $element.height) {
            $center = [pscustomobject]@{
                x = [math]::Round(([double]$element.x + [double]$element.width / 2), 2)
                y = [math]::Round(([double]$element.y + [double]$element.height / 2), 2)
            }
        }
        $items.Add([pscustomobject]@{
            id = "i$index"
            role = $role
            text = if ($text) { $text } else { $null }
            action_hints = $hints
            uia_path = $element.uia_path
            frame = [pscustomobject]@{
                x = $element.x
                y = $element.y
                width = $element.width
                height = $element.height
            }
            center = $center
            automation_id = $element.automation_id
            class_name = $element.class_name
            patterns = $patterns
        })
        $index++
        if ($items.Count -ge $Limit) { break }
    }
    $windowRect = Window-Rect ([IntPtr]$Hwnd)
    if ($windowRect -and $items.Count -lt $Limit) {
        $appText = ("$($tree.app_name) $($tree.title)").ToLowerInvariant()
        $virtualSpecs = @()
        if ($appText.Contains("feishu") -or $appText.Contains("lark") -or $appText.Contains("$([char]0x98de)$([char]0x4e66)")) {
            $x = [double]$windowRect.x
            $y = [double]$windowRect.y
            $width = [double]$windowRect.width
            $height = [double]$windowRect.height
            $dockCenterY = $y + $height - 52
            $dockSlotSize = 56
            $virtualSpecs = @(
                [pscustomobject]@{
                text = "Feishu/Lark profile/avatar button; use only to open the profile card for verification, not as the organization switcher"
                frame = [pscustomobject]@{ x = $x + 16; y = $y + 22; width = 82; height = 64 }
                },
                [pscustomobject]@{
                text = "Feishu/Lark bottom organization dock; probe/click visible org icons here and verify by OCRing the profile card"
                frame = [pscustomobject]@{ x = $x + 20; y = $y + $height - 104; width = [Math]::Min(340, [Math]::Max(220, $width * 0.20)); height = 92 }
                }
            )
            for ($slot = 0; $slot -lt 4; $slot++) {
                $centerX = $x + 56 + (60 * $slot)
                $slotText = if ($slot -eq 3) {
                    "Feishu/Lark bottom organization dock more button; open only to inspect visible existing orgs, never join/create/login"
                } else {
                    "Feishu/Lark bottom organization dock slot $($slot + 1); current org is usually first after switching, verify via profile card OCR"
                }
                $virtualSpecs += [pscustomobject]@{
                    text = $slotText
                    frame = [pscustomobject]@{ x = $centerX - ($dockSlotSize / 2); y = $dockCenterY - ($dockSlotSize / 2); width = $dockSlotSize; height = $dockSlotSize }
                }
            }
            $virtualSpecs += [pscustomobject]@{
                text = "Feishu/Lark global search candidate; Ctrl+K is usually preferred when UIA exposes no search edit"
                frame = [pscustomobject]@{ x = $x + 70; y = $y + 36; width = [Math]::Min(360, [Math]::Max(220, $width * 0.34)); height = 46 }
            }
            $virtualSpecs += [pscustomobject]@{
                text = "Feishu/Lark first search/chat result candidate; re-observe or targeted-OCR this row before opening"
                frame = [pscustomobject]@{ x = $x + 64; y = $y + 88; width = [Math]::Min(420, [Math]::Max(260, $width * 0.38)); height = 78 }
            }
            $virtualSpecs += [pscustomobject]@{
                text = "Feishu/Lark compose input candidate; use only after title or placeholder confirms the recipient"
                frame = [pscustomobject]@{ x = $x + $width * 0.28; y = $y + $height - 132; width = $width * 0.68; height = 96 }
            }
        }
        foreach ($spec in @($virtualSpecs)) {
            if ($items.Count -ge $Limit) { break }
            if ($queryLower.Length -gt 0) {
                $haystack = ("virtualregion $($spec.text)").ToLowerInvariant()
                if (-not $haystack.Contains($queryLower)) { continue }
            }
            $frame = $spec.frame
            $center = [pscustomobject]@{
                x = [math]::Round(([double]$frame.x + [double]$frame.width / 2), 2)
                y = [math]::Round(([double]$frame.y + [double]$frame.height / 2), 2)
            }
            $items.Add([pscustomobject]@{
                id = "v$index"
                role = "VirtualRegion"
                text = $spec.text
                action_hints = @("click", "ocr")
                uia_path = $null
                frame = [pscustomobject]@{
                    x = [math]::Round([double]$frame.x, 2)
                    y = [math]::Round([double]$frame.y, 2)
                    width = [math]::Round([double]$frame.width, 2)
                    height = [math]::Round([double]$frame.height, 2)
                }
                center = $center
                automation_id = $null
                class_name = "app_profile"
                patterns = @()
                direct_uia = $false
                source = "feishu_app_profile"
            })
            $index++
        }
    }
    return [pscustomobject]@{
        hwnd = $Hwnd
        app_name = $tree.app_name
        pid = $tree.pid
        title = $tree.title
        view = $tree.view
        accessibility_hint = $tree.accessibility_hint
        count = $items.Count
        elements = $items
        source_stats = $tree.stats
        status = "success"
    }
}

function Resolve-UiaPath {
    param([Int64]$Hwnd, [string]$UiaPath)
    $root = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$Hwnd)
    if ($null -eq $root) { throw "UIAutomation root not found for hwnd $Hwnd" }
    $path = $UiaPath
    $view = "control"
    if ($path.StartsWith("uia:")) {
        $parts = $path -split ":", 4
        if ($parts.Count -eq 3) {
            $path = $parts[2]
        } elseif ($parts.Count -eq 4) {
            $view = Normalize-UiaView $parts[2]
            $path = $parts[3]
        }
    }
    if (-not $path.StartsWith("root")) { throw "UIA path must start with root: $UiaPath" }

    $current = $root
    $segments = @($path.Split(".") | Select-Object -Skip 1)
    $walker = Get-UiaTreeWalker $view
    foreach ($segment in $segments) {
        if ($segment -notmatch '^children\[(\d+)\]$') { throw "unsupported UIA path segment '$segment' in $UiaPath" }
        $target = [int]$Matches[1]
        $child = $walker.GetFirstChild($current)
        $index = 0
        while ($null -ne $child -and $index -lt $target) {
            $child = $walker.GetNextSibling($child)
            $index++
        }
        if ($null -eq $child) { throw "UIA child index $target is unavailable in $UiaPath" }
        $current = $child
    }
    return $current
}

function Try-Pattern {
    param($Element, $Pattern)
    $instance = $null
    if ($Element.TryGetCurrentPattern($Pattern, [ref]$instance)) { return $instance }
    return $null
}

function Click-Element-Center {
    param($Element)
    $frame = Element-Frame $Element
    if ($null -eq $frame) { throw "element has no frame for coordinate fallback" }
    $x = [int][math]::Round($frame.x + $frame.width / 2)
    $y = [int][math]::Round($frame.y + $frame.height / 2)
    Send-Mouse -Action "click" -InputArgs @("$x", "$y")
    return [pscustomobject]@{
        x = $x
        y = $y
        frame = $frame
    }
}

function Element-Snapshot {
    param($Element, [string]$UiaPath = $null)
    $frame = Element-Frame $Element
    $center = $null
    if ($frame) {
        $center = [pscustomobject]@{
            x = [math]::Round([double]$frame.x + [double]$frame.width / 2, 2)
            y = [math]::Round([double]$frame.y + [double]$frame.height / 2, 2)
        }
    }
    $patterns = @(Element-Patterns $Element)
    return [pscustomobject]@{
        role = (Element-Role $Element)
        text = (Get-Element-Text $Element)
        frame = $frame
        center = $center
        uia_path = $UiaPath
        automation_id = Try-Value { $Element.Current.AutomationId }
        class_name = Try-Value { $Element.Current.ClassName }
        framework_id = Try-Value { $Element.Current.FrameworkId }
        native_window_handle = Try-Value { [Int64]$Element.Current.NativeWindowHandle } 0
        process_id = Try-Value { [int]$Element.Current.ProcessId } 0
        is_enabled = Try-Value { [bool]$Element.Current.IsEnabled }
        is_offscreen = Try-Value { [bool]$Element.Current.IsOffscreen }
        patterns = $patterns
        action_hints = @(Element-Action-Hints (Element-Role $Element) $patterns)
    }
}

function Find-UiaPathForElement {
    param([Int64]$Hwnd, $TargetElement, [string]$View = "control")
    $root = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$Hwnd)
    if ($null -eq $root -or $null -eq $TargetElement) { return $null }
    $targetKey = Runtime-Key $TargetElement
    if ((Runtime-Key $root) -eq $targetKey) { return (Root-UiaPath $Hwnd $View) }

    $walker = Get-UiaTreeWalker $View
    $queue = New-Object System.Collections.ArrayList
    [void]$queue.Add([pscustomobject]@{ element = $root; depth = 0; path = (Root-UiaPath $Hwnd $View) })
    $readIndex = 0
    $visited = @{}
    $started = Get-Date
    while ($readIndex -lt $queue.Count) {
        if ($queue.Count -ge $script:MaxTraversalElements -or ((Get-Date) - $started).TotalSeconds -gt $script:MaxTraversalSeconds) {
            break
        }
        $item = $queue[$readIndex]
        $readIndex++
        if ($item.depth -gt $script:MaxTraversalDepth) { continue }
        $element = $item.element
        $key = Runtime-Key $element
        if ($visited.ContainsKey($key)) { continue }
        $visited[$key] = $true
        if ($key -eq $targetKey) { return $item.path }

        $child = $null
        try { $child = $walker.GetFirstChild($element) } catch {}
        $childIndex = 0
        while ($null -ne $child -and $childIndex -lt $script:MaxChildrenPerElement) {
            [void]$queue.Add([pscustomobject]@{
                element = $child
                depth = $item.depth + 1
                path = "$($item.path).children[$childIndex]"
            })
            try { $child = $walker.GetNextSibling($child) } catch { $child = $null }
            $childIndex++
        }
    }
    return $null
}

function Hwnd-From-Point {
    param([int]$X, [int]$Y)
    $point = New-Object LuopanNativeMethods+POINT
    $point.X = $X
    $point.Y = $Y
    return [Int64][LuopanNativeMethods]::WindowFromPoint($point)
}

function Expand-Rect-Around-Point {
    param($Rect, [double]$PointX, [double]$PointY, [int]$Padding = 16, [int]$MinWidth = 220, [int]$MinHeight = 80, [int]$MaxElementWidth = 640, [int]$MaxElementHeight = 260)
    $useElementRect = $false
    if ($Rect) {
        $useElementRect = ([double]$Rect.width -le $MaxElementWidth -and [double]$Rect.height -le $MaxElementHeight)
    }
    if ($useElementRect) {
        $x = [double]$Rect.x - $Padding
        $y = [double]$Rect.y - $Padding
        $width = [double]$Rect.width + 2 * $Padding
        $height = [double]$Rect.height + 2 * $Padding
        if ($width -lt $MinWidth) {
            $delta = ($MinWidth - $width) / 2
            $x -= $delta
            $width = $MinWidth
        }
        if ($height -lt $MinHeight) {
            $delta = ($MinHeight - $height) / 2
            $y -= $delta
            $height = $MinHeight
        }
    } else {
        $width = $MinWidth
        $height = $MinHeight
        $x = $PointX - $width / 2
        $y = $PointY - $height / 2
    }
    return [pscustomobject]@{
        x = [math]::Round([Math]::Max(0, $x), 2)
        y = [math]::Round([Math]::Max(0, $y), 2)
        width = [math]::Round([Math]::Max(1, $width), 2)
        height = [math]::Round([Math]::Max(1, $height), 2)
    }
}

function Format-Rect {
    param($Rect)
    return ("{0},{1},{2},{3}" -f $Rect.x, $Rect.y, $Rect.width, $Rect.height)
}

function Probe-Point {
    param(
        [string]$Hwnd,
        [string]$ProcessId,
        [string]$Target,
        [double]$X,
        [double]$Y,
        [string]$View = "control",
        [int]$HoverMs = 350,
        [int]$Padding = 16,
        [bool]$DoOcr = $true
    )
    $resolvedHwnd = $null
    if ($Hwnd -or $ProcessId -or $Target) {
        $resolvedHwnd = Resolve-Hwnd $Hwnd $ProcessId $Target
    } else {
        $resolvedHwnd = Hwnd-From-Point ([int][math]::Round($X)) ([int][math]::Round($Y))
    }
    [void][LuopanNativeMethods]::SetCursorPos(([int][math]::Round($X)), ([int][math]::Round($Y)))
    if ($HoverMs -gt 0) { Start-Sleep -Milliseconds $HoverMs }

    $point = [System.Windows.Point]::new($X, $Y)
    $element = [System.Windows.Automation.AutomationElement]::FromPoint($point)
    $viewName = Normalize-UiaView $View
    $uiaPath = $null
    if ($resolvedHwnd -and $element) {
        $uiaPath = Find-UiaPathForElement ([Int64]$resolvedHwnd) $element $viewName
        if (-not $uiaPath -and $viewName -ne "raw") {
            $rawPath = Find-UiaPathForElement ([Int64]$resolvedHwnd) $element "raw"
            if ($rawPath) { $uiaPath = $rawPath }
        }
    }
    $snapshot = if ($element) { Element-Snapshot $element $uiaPath } else { $null }
    $ocr = $null
    if ($DoOcr) {
        $snapshotFrame = if ($snapshot) { $snapshot.frame } else { $null }
        $region = Expand-Rect-Around-Point $snapshotFrame $X $Y $Padding
        $ocr = Invoke-Ocr $null ([string]$resolvedHwnd) $null $null (Format-Rect $region) $null
    }
    return [pscustomobject]@{
        hwnd = $resolvedHwnd
        point = [pscustomobject]@{ x = [math]::Round($X, 2); y = [math]::Round($Y, 2) }
        view = $viewName
        hover_ms = $HoverMs
        element = $snapshot
        ocr = $ocr
        status = "success"
    }
}

function Invoke-Uia {
    param([string]$Action, [Int64]$Hwnd, [string]$UiaPath, [object]$Value = $null)
    [void](Activate-Window ([IntPtr]$Hwnd))
    $element = Resolve-UiaPath $Hwnd $UiaPath
    $role = Element-Role $element
    $mode = $null
    $click = $null

    switch ($Action) {
        "click" {
            $click = Click-Element-Center $element
            $mode = "coordinate_click"
        }
        "focus" {
            $element.SetFocus()
            $mode = "focus"
        }
        "press" {
            $invoke = Try-Pattern $element ([System.Windows.Automation.InvokePattern]::Pattern)
            if ($invoke) {
                ([System.Windows.Automation.InvokePattern]$invoke).Invoke()
                $mode = "invoke"
            } else {
                $click = Click-Element-Center $element
                $mode = "coordinate_click_fallback"
            }
        }
        "select" {
            $selection = Try-Pattern $element ([System.Windows.Automation.SelectionItemPattern]::Pattern)
            if ($selection) {
                ([System.Windows.Automation.SelectionItemPattern]$selection).Select()
                $mode = "selection_item"
                if ($role -in @("ListItem", "DataItem", "TreeItem", "TabItem")) {
                    Start-Sleep -Milliseconds 80
                    $click = Click-Element-Center $element
                    $mode = "selection_item_then_coordinate_click"
                }
            } else {
                $toggle = Try-Pattern $element ([System.Windows.Automation.TogglePattern]::Pattern)
                if ($toggle) {
                    ([System.Windows.Automation.TogglePattern]$toggle).Toggle()
                    $mode = "toggle"
                } else {
                    $click = Click-Element-Center $element
                    $mode = "coordinate_click_fallback"
                }
            }
        }
        "set_value" {
            $valuePattern = Try-Pattern $element ([System.Windows.Automation.ValuePattern]::Pattern)
            if (-not $valuePattern) { throw "element does not support ValuePattern: $UiaPath" }
            ([System.Windows.Automation.ValuePattern]$valuePattern).SetValue([string]$Value)
            $mode = "value_pattern"
        }
        default {
            $valuePattern = Try-Pattern $element ([System.Windows.Automation.ValuePattern]::Pattern)
            if ($valuePattern -and $role -in @("Edit", "ComboBox")) {
                $element.SetFocus()
                $mode = "focus_text"
            } elseif ($role -in @("ListItem", "DataItem", "TreeItem", "TabItem", "Hyperlink")) {
                $click = Click-Element-Center $element
                $mode = "coordinate_click_for_activation"
            } else {
                $invoke = Try-Pattern $element ([System.Windows.Automation.InvokePattern]::Pattern)
                if ($invoke) {
                    ([System.Windows.Automation.InvokePattern]$invoke).Invoke()
                    $mode = "invoke"
                } else {
                    $selection = Try-Pattern $element ([System.Windows.Automation.SelectionItemPattern]::Pattern)
                    if ($selection) {
                        ([System.Windows.Automation.SelectionItemPattern]$selection).Select()
                        $mode = "selection_item"
                        if ($role -in @("ListItem", "DataItem", "TreeItem", "TabItem")) {
                            Start-Sleep -Milliseconds 80
                            $click = Click-Element-Center $element
                            $mode = "selection_item_then_coordinate_click"
                        }
                    } else {
                        try {
                            $element.SetFocus()
                            $mode = "focus_fallback"
                        } catch {
                            $click = Click-Element-Center $element
                            $mode = "coordinate_click_fallback"
                        }
                    }
                }
            }
        }
    }

    return [pscustomobject]@{
        action = $Action
        hwnd = $Hwnd
        uia_path = $UiaPath
        role = $role
        mode = $mode
        click = $click
        status = "success"
    }
}

function Convert-KeyCombo {
    param([string]$Combo)
    $parts = @($Combo.Split("+") | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ })
    if ($parts.Count -eq 0) { throw "empty key combo" }
    $key = $parts[-1]
    $prefix = ""
    foreach ($part in $parts[0..([Math]::Max(0, $parts.Count - 2))]) {
        switch ($part) {
            "ctrl" { $prefix += "^" }
            "control" { $prefix += "^" }
            "alt" { $prefix += "%" }
            "option" { $prefix += "%" }
            "shift" { $prefix += "+" }
            default {}
        }
    }
    $special = @{
        "enter" = "{ENTER}"
        "return" = "{ENTER}"
        "esc" = "{ESC}"
        "escape" = "{ESC}"
        "tab" = "{TAB}"
        "backspace" = "{BACKSPACE}"
        "delete" = "{DELETE}"
        "left" = "{LEFT}"
        "right" = "{RIGHT}"
        "up" = "{UP}"
        "down" = "{DOWN}"
        "home" = "{HOME}"
        "end" = "{END}"
        "pageup" = "{PGUP}"
        "pagedown" = "{PGDN}"
        "space" = " "
    }
    if ($special.ContainsKey($key)) {
        return "$prefix$($special[$key])"
    }
    if ($key -match '^f(\d{1,2})$') {
        return "$prefix{F$($Matches[1])}"
    }
    return "$prefix$key"
}

function Send-Mouse {
    param([string]$Action, [string[]]$InputArgs)
    $MOUSEEVENTF_MOVE = 0x0001
    $MOUSEEVENTF_LEFTDOWN = 0x0002
    $MOUSEEVENTF_LEFTUP = 0x0004
    $MOUSEEVENTF_RIGHTDOWN = 0x0008
    $MOUSEEVENTF_RIGHTUP = 0x0010
    $MOUSEEVENTF_WHEEL = 0x0800

    if ($Action -eq "scroll") {
        if ($InputArgs.Count -lt 3) { throw "scroll requires x y deltaY [deltaX]" }
        $x = [int][double]$InputArgs[0]
        $y = [int][double]$InputArgs[1]
        $delta = [int]$InputArgs[2] * -120
        [void][LuopanNativeMethods]::SetCursorPos($x, $y)
        [LuopanNativeMethods]::mouse_event($MOUSEEVENTF_WHEEL, 0, 0, $delta, [UIntPtr]::Zero)
        return
    }

    if ($InputArgs.Count -lt 2) { throw "$Action requires x y" }
    $mx = [int][double]$InputArgs[0]
    $my = [int][double]$InputArgs[1]
    [void][LuopanNativeMethods]::SetCursorPos($mx, $my)
    Start-Sleep -Milliseconds 40
    switch ($Action) {
        "mousemove" {
            [LuopanNativeMethods]::mouse_event($MOUSEEVENTF_MOVE, 0, 0, 0, [UIntPtr]::Zero)
        }
        "click" {
            [LuopanNativeMethods]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
            Start-Sleep -Milliseconds 40
            [LuopanNativeMethods]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, [UIntPtr]::Zero)
        }
        "doubleclick" {
            1..2 | ForEach-Object {
                [LuopanNativeMethods]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
                Start-Sleep -Milliseconds 35
                [LuopanNativeMethods]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, [UIntPtr]::Zero)
                Start-Sleep -Milliseconds 65
            }
        }
        "rightclick" {
            [LuopanNativeMethods]::mouse_event($MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, [UIntPtr]::Zero)
            Start-Sleep -Milliseconds 40
            [LuopanNativeMethods]::mouse_event($MOUSEEVENTF_RIGHTUP, 0, 0, 0, [UIntPtr]::Zero)
        }
    }
}

function Send-ClipboardText {
    param([string]$Text)
    $oldClipboard = $null
    $hadClipboard = $false
    try {
        $oldClipboard = Get-Clipboard -Raw
        $hadClipboard = $true
    } catch {}
    Set-Clipboard -Value $Text
    Start-Sleep -Milliseconds 70
    [System.Windows.Forms.SendKeys]::SendWait("^v")
    Start-Sleep -Milliseconds 150
    if ($hadClipboard) {
        try { Set-Clipboard -Value $oldClipboard } catch {}
    }
}

function Send-StreamText {
    param([string]$Text, [int]$DelayMs = 18)
    if (-not [LuopanNativeMethods]::SendUnicodeText($Text, $DelayMs)) {
        throw "SendInput unicode text failed"
    }
}

function Send-InputAction {
    param([string]$Action, [string[]]$InputArgs, [string]$Hwnd = $null)
    if ($Hwnd) { [void](Activate-Window ([IntPtr][Int64]$Hwnd)) }
    $mode = $Action
    switch ($Action) {
        "keypress" {
            if ($InputArgs.Count -lt 1) { throw "keypress requires a combo string" }
            [System.Windows.Forms.SendKeys]::SendWait((Convert-KeyCombo $InputArgs[0]))
            $mode = "sendkeys"
        }
        "writetext" {
            if ($InputArgs.Count -lt 1) { throw "writetext requires text" }
            $delay = if ($InputArgs.Count -ge 2) { [int]$InputArgs[1] } else { 18 }
            Send-StreamText $InputArgs[0] $delay
            $mode = "unicode_stream"
        }
        "streamtext" {
            if ($InputArgs.Count -lt 1) { throw "streamtext requires text" }
            $delay = if ($InputArgs.Count -ge 2) { [int]$InputArgs[1] } else { 18 }
            Send-StreamText $InputArgs[0] $delay
            $mode = "unicode_stream"
        }
        "typetext" {
            if ($InputArgs.Count -lt 1) { throw "typetext requires text" }
            $delay = if ($InputArgs.Count -ge 2) { [int]$InputArgs[1] } else { 18 }
            Send-StreamText $InputArgs[0] $delay
            $mode = "unicode_stream"
        }
        "pastetext" {
            if ($InputArgs.Count -lt 1) { throw "pastetext requires text" }
            Send-ClipboardText $InputArgs[0]
            $mode = "clipboard_paste"
        }
        "clipboardtext" {
            if ($InputArgs.Count -lt 1) { throw "clipboardtext requires text" }
            Send-ClipboardText $InputArgs[0]
            $mode = "clipboard_paste"
        }
        "click" { Send-Mouse -Action "click" -InputArgs $InputArgs }
        "doubleclick" { Send-Mouse -Action "doubleclick" -InputArgs $InputArgs }
        "rightclick" { Send-Mouse -Action "rightclick" -InputArgs $InputArgs }
        "mousemove" { Send-Mouse -Action "mousemove" -InputArgs $InputArgs }
        "scroll" { Send-Mouse -Action "scroll" -InputArgs $InputArgs }
        default { throw "unsupported input action: $Action" }
    }
    return [pscustomobject]@{ action = $Action; args = $InputArgs; hwnd = $Hwnd; mode = $mode; status = "success" }
}

function Send-InputSequence {
    param([string]$Json, [string]$Hwnd = $null)
    if ([string]::IsNullOrWhiteSpace($Json)) {
        throw "input-sequence requires --json"
    }
    $items = ConvertFrom-Json -InputObject $Json
    if ($null -eq $items) {
        throw "input-sequence JSON must contain at least one action"
    }
    if ($items -isnot [System.Array]) {
        $items = @($items)
    }
    $results = New-Object System.Collections.Generic.List[object]
    foreach ($item in $items) {
        $action = if ($item.action) { [string]$item.action } elseif ($item.type) { [string]$item.type } else { "" }
        if ([string]::IsNullOrWhiteSpace($action)) {
            throw "input-sequence action is missing action/type"
        }
        if ($action -eq "sleep" -or $action -eq "wait") {
            $ms = if ($item.ms) { [int]$item.ms } elseif ($item.wait_ms) { [int]$item.wait_ms } else { 100 }
            Start-Sleep -Milliseconds $ms
            $results.Add([pscustomobject]@{ action = $action; ms = $ms; status = "success" })
            continue
        }

        $inputArgs = New-Object System.Collections.Generic.List[string]
        if ($item.args) {
            foreach ($arg in @($item.args)) {
                $inputArgs.Add([string]$arg)
            }
        } elseif ($item.key) {
            $inputArgs.Add([string]$item.key)
        } elseif ($null -ne $item.text) {
            $inputArgs.Add([string]$item.text)
        } elseif ($null -ne $item.x -and $null -ne $item.y) {
            $inputArgs.Add([string]$item.x)
            $inputArgs.Add([string]$item.y)
            if ($null -ne $item.deltaY) { $inputArgs.Add([string]$item.deltaY) }
            if ($null -ne $item.deltaX) { $inputArgs.Add([string]$item.deltaX) }
        }
        $results.Add((Send-InputAction -Action $action -InputArgs $inputArgs.ToArray() -Hwnd $Hwnd))
    }
    return [pscustomobject]@{ action = "input-sequence"; count = $results.Count; hwnd = $Hwnd; results = $results; status = "success" }
}

function Save-Screenshot {
    param([string]$Path, [int]$X, [int]$Y, [int]$Width, [int]$Height)
    $Width = [Math]::Max(1, $Width)
    $Height = [Math]::Max(1, $Height)
    $bitmap = [System.Drawing.Bitmap]::new([int]$Width, [int]$Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen($X, $Y, 0, 0, $bitmap.Size)
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Rect-Object {
    param([object]$Rect)
    if ($null -eq $Rect) { return $null }
    return [pscustomobject]@{
        x = [math]::Round([double]$Rect.x, 2)
        y = [math]::Round([double]$Rect.y, 2)
        width = [math]::Round([double]$Rect.width, 2)
        height = [math]::Round([double]$Rect.height, 2)
    }
}

function Save-WindowScreenshot {
    param([string]$Path, [Int64]$Hwnd)
    $rect = Window-Rect ([IntPtr]$Hwnd)
    if ($null -eq $rect) { throw "cannot capture hwnd $Hwnd because it has no valid window rect" }
    $width = [Math]::Max(1, [int][math]::Round([double]$rect.width))
    $height = [Math]::Max(1, [int][math]::Round([double]$rect.height))
    $bitmap = [System.Drawing.Bitmap]::new($width, $height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $hdc = $graphics.GetHdc()
        try {
            $ok = [LuopanNativeMethods]::PrintWindow([IntPtr]$Hwnd, $hdc, 2)
        } finally {
            $graphics.ReleaseHdc($hdc)
        }
        if (-not $ok) {
            $graphics.CopyFromScreen([int]$rect.x, [int]$rect.y, 0, 0, $bitmap.Size)
        }
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
    return (Rect-Object $rect)
}

function Save-WindowRegionScreenshot {
    param([string]$Path, [Int64]$Hwnd, [object]$ScreenRect)
    $windowRect = Window-Rect ([IntPtr]$Hwnd)
    if ($null -eq $windowRect) { throw "cannot capture hwnd $Hwnd because it has no valid window rect" }
    $screenRect = Rect-Object $ScreenRect
    $fullPath = Join-Path ([IO.Path]::GetTempPath()) ("luopan-windows-ocr-full-{0}-{1}.png" -f ([System.Diagnostics.Process]::GetCurrentProcess().Id), ([DateTimeOffset]::Now.ToUnixTimeMilliseconds()))
    $capturedWindowRect = Save-WindowScreenshot $fullPath $Hwnd
    $sourceBitmap = $null
    try {
        $sourceBitmap = [System.Drawing.Bitmap]::FromFile($fullPath)
        $sourceX = [int][math]::Round([double]$screenRect.x - [double]$capturedWindowRect.x)
        $sourceY = [int][math]::Round([double]$screenRect.y - [double]$capturedWindowRect.y)
        $sourceWidth = [int][math]::Round([double]$screenRect.width)
        $sourceHeight = [int][math]::Round([double]$screenRect.height)
        $sourceX = [Math]::Max(0, [Math]::Min($sourceX, $sourceBitmap.Width - 1))
        $sourceY = [Math]::Max(0, [Math]::Min($sourceY, $sourceBitmap.Height - 1))
        $sourceWidth = [Math]::Max(1, [Math]::Min($sourceWidth, $sourceBitmap.Width - $sourceX))
        $sourceHeight = [Math]::Max(1, [Math]::Min($sourceHeight, $sourceBitmap.Height - $sourceY))
        $scale = if ($sourceWidth -lt 360 -or $sourceHeight -lt 120) { 3 } elseif ($sourceWidth -lt 720 -or $sourceHeight -lt 240) { 2 } else { 1 }
        $targetBitmap = [System.Drawing.Bitmap]::new($sourceWidth * $scale, $sourceHeight * $scale)
        $graphics = [System.Drawing.Graphics]::FromImage($targetBitmap)
        try {
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.DrawImage(
                $sourceBitmap,
                [System.Drawing.Rectangle]::new(0, 0, $sourceWidth * $scale, $sourceHeight * $scale),
                [System.Drawing.Rectangle]::new($sourceX, $sourceY, $sourceWidth, $sourceHeight),
                [System.Drawing.GraphicsUnit]::Pixel
            )
            $targetBitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $graphics.Dispose()
            $targetBitmap.Dispose()
        }
    } finally {
        if ($sourceBitmap) { $sourceBitmap.Dispose() }
        try { Remove-Item -LiteralPath $fullPath -Force } catch {}
    }
    return [pscustomobject]@{ region = $screenRect; scale = $scale }
}

function Parse-Rect {
    param([string]$Rect)
    if (-not $Rect) { return $null }
    $parts = @($Rect -split "[, ]+" | Where-Object { $_ })
    if ($parts.Count -ne 4) { throw "--rect requires x,y,width,height" }
    $x = [double]$parts[0]
    $y = [double]$parts[1]
    $width = [double]$parts[2]
    $height = [double]$parts[3]
    if ($width -le 0 -or $height -le 0) { throw "--rect width and height must be positive" }
    return [pscustomobject]@{
        x = [math]::Round($x, 2)
        y = [math]::Round($y, 2)
        width = [math]::Round($width, 2)
        height = [math]::Round($height, 2)
    }
}

function Await-WinRt {
    param($AsyncOperation, [Type]$ResultType = $null)
    $methods = @([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1
    })

    if ($ResultType) {
        $method = $methods | Where-Object {
            $_.IsGenericMethodDefinition -and
            $_.GetGenericArguments().Count -eq 1 -and
            $_.ToString().Contains("Windows.Foundation.IAsyncOperation")
        } | Select-Object -First 1
        if (-not $method) { throw "could not find WinRT AsTask<TResult> helper" }
        $task = $method.MakeGenericMethod($ResultType).Invoke($null, @($AsyncOperation))
    } else {
        $method = $methods | Where-Object {
            -not $_.IsGenericMethodDefinition -and
            $_.GetParameters()[0].ParameterType.FullName -eq "Windows.Foundation.IAsyncAction"
        } | Select-Object -First 1
        if (-not $method) { throw "could not find WinRT AsTask action helper" }
        $task = $method.Invoke($null, @($AsyncOperation))
    }

    $task.Wait()
    if ($ResultType) { return $task.Result }
    return $null
}

function Invoke-Ocr {
    param(
        [string]$ImagePath,
        [string]$Hwnd,
        [string]$ProcessId,
        [string]$Identifier,
        [string]$RectText,
        [string]$UiaPath
    )
    $source = "image_path"
    $capture = $null
    $ocrScale = 1.0
    if (-not $ImagePath) {
        $resolvedHwnd = Resolve-Hwnd $Hwnd $ProcessId $Identifier
        [void](Activate-Window ([IntPtr]$resolvedHwnd))
        Start-Sleep -Milliseconds 250
        if ($UiaPath) {
            $element = Resolve-UiaPath $resolvedHwnd $UiaPath
            $regionRect = Element-Frame $element
            $source = "uia_element_screenshot"
        } elseif ($RectText) {
            $regionRect = Parse-Rect $RectText
            $source = "region_screenshot"
        } else {
            $regionRect = Window-Rect ([IntPtr]$resolvedHwnd)
            $source = "window_screenshot"
        }
        if ($null -eq $regionRect) { throw "cannot capture OCR region because it has no valid rect" }
        $currentProcessId = [System.Diagnostics.Process]::GetCurrentProcess().Id
        $ImagePath = Join-Path ([IO.Path]::GetTempPath()) ("luopan-windows-ocr-{0}-{1}.png" -f $currentProcessId, ([DateTimeOffset]::Now.ToUnixTimeMilliseconds()))
        if ($source -eq "window_screenshot") {
            $capturedRegion = Save-WindowScreenshot $ImagePath $resolvedHwnd
            $ocrScale = 1.0
        } else {
            $regionCapture = Save-WindowRegionScreenshot $ImagePath $resolvedHwnd $regionRect
            $capturedRegion = Rect-Object $regionCapture.region
            $ocrScale = [double]$regionCapture.scale
        }
        $capture = [pscustomobject]@{ hwnd = $resolvedHwnd; region = (Rect-Object $capturedRegion); scale = $ocrScale; uia_path = $UiaPath; screenshot = $ImagePath }
    }

    try {
        [void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
        [void][Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
        [void][Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        [void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
        [void][Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
        [void][Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
        [void][Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType = WindowsRuntime]
        $file = Await-WinRt ([Windows.Storage.StorageFile]::GetFileFromPathAsync((Resolve-Path -LiteralPath $ImagePath).Path)) ([Windows.Storage.StorageFile])
        $stream = Await-WinRt ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
        $decoder = Await-WinRt ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Await-WinRt ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
        if ($null -eq $engine) { throw "Windows OCR engine is unavailable for current user profile languages" }
        $result = Await-WinRt ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
        $lines = @()
        foreach ($line in $result.Lines) {
            $text = Normalize-Text $line.Text
            if (-not $text) { continue }
            $boxes = @($line.Words | ForEach-Object { $_.BoundingRect })
            if ($boxes.Count -gt 0) {
                $rawMinX = ($boxes | Measure-Object X -Minimum).Minimum
                $rawMinY = ($boxes | Measure-Object Y -Minimum).Minimum
                $rawMaxX = ($boxes | ForEach-Object { $_.X + $_.Width } | Measure-Object -Maximum).Maximum
                $rawMaxY = ($boxes | ForEach-Object { $_.Y + $_.Height } | Measure-Object -Maximum).Maximum
                $minX = [double]$rawMinX / $ocrScale
                $minY = [double]$rawMinY / $ocrScale
                $maxX = [double]$rawMaxX / $ocrScale
                $maxY = [double]$rawMaxY / $ocrScale
                $frame = [pscustomobject]@{ x = [math]::Round($minX, 2); y = [math]::Round($minY, 2); width = [math]::Round($maxX - $minX, 2); height = [math]::Round($maxY - $minY, 2) }
                if ($capture -and $capture.region) {
                    $screenFrame = [pscustomobject]@{
                        x = [math]::Round([double]$capture.region.x + [double]$minX, 2)
                        y = [math]::Round([double]$capture.region.y + [double]$minY, 2)
                        width = [math]::Round([double]($maxX - $minX), 2)
                        height = [math]::Round([double]($maxY - $minY), 2)
                    }
                    $center = [pscustomobject]@{
                        x = [math]::Round([double]$screenFrame.x + [double]$screenFrame.width / 2, 2)
                        y = [math]::Round([double]$screenFrame.y + [double]$screenFrame.height / 2, 2)
                    }
                } else {
                    $screenFrame = $null
                    $center = $null
                }
            } else {
                $frame = $null
                $screenFrame = $null
                $center = $null
            }
            $lines += [pscustomobject]@{ text = $text; frame = $frame; screen_frame = $screenFrame; center = $center }
        }
        return [pscustomobject]@{
            image_path = $ImagePath
            source = $source
            capture = $capture
            lines = $lines
            text = ($lines | ForEach-Object { $_.text }) -join [Environment]::NewLine
            status = "success"
        }
    } catch {
        return [pscustomobject]@{
            image_path = $ImagePath
            source = $source
            capture = $capture
            lines = @()
            text = ""
            status = "unavailable"
            error = $_.Exception.Message
        }
    }
}

try {
    switch ($Command) {
        "list-apps" {
            $query = Get-OptionValue $Rest "--query" ""
            $limit = [int](Get-OptionValue $Rest "--limit" "100")
            $apps = @()
            $apps += @(Running-Windows $query)
            $apps += @(Start-Apps $query)
            Write-Json ([pscustomobject]@{ applications = @($apps | Select-Object -First $limit); count = [Math]::Min($apps.Count, $limit) })
        }
        "open" {
            $pos = @(Positional-Args $Rest)
            if ($pos.Count -lt 1) { throw "open requires an identifier" }
            Write-Json (Open-App $pos[0] (Has-Flag $Rest "--no-activate"))
        }
        "traverse" {
            $hwnd = Resolve-Hwnd (Get-OptionValue $Rest "--hwnd") (Get-OptionValue $Rest "--pid") (Get-OptionValue $Rest "--target")
            Write-Json (Traverse-Window $hwnd (Has-Flag $Rest "--visible-only") (Has-Flag $Rest "--no-activate") (Get-OptionValue $Rest "--view" "control"))
        }
        "elements" {
            $hwnd = Resolve-Hwnd (Get-OptionValue $Rest "--hwnd") (Get-OptionValue $Rest "--pid") (Get-OptionValue $Rest "--target")
            $limit = [int](Get-OptionValue $Rest "--limit" "160")
            $query = Get-OptionValue $Rest "--query" ""
            Write-Json (Get-Interactive-Elements $hwnd (Has-Flag $Rest "--visible-only") (Has-Flag $Rest "--no-activate") $limit $query (Get-OptionValue $Rest "--view" "control"))
        }
        "probe" {
            $pos = @(Positional-Args $Rest)
            $xText = Get-OptionValue $Rest "--x"
            $yText = Get-OptionValue $Rest "--y"
            if (-not $xText -and $pos.Count -ge 1) { $xText = $pos[0] }
            if (-not $yText -and $pos.Count -ge 2) { $yText = $pos[1] }
            if (-not $xText -or -not $yText) { throw "probe requires --x and --y or positional x y" }
            Write-Json (Probe-Point `
                (Get-OptionValue $Rest "--hwnd") `
                (Get-OptionValue $Rest "--pid") `
                (Get-OptionValue $Rest "--target") `
                ([double]$xText) `
                ([double]$yText) `
                (Get-OptionValue $Rest "--view" "control") `
                ([int](Get-OptionValue $Rest "--hover-ms" "350")) `
                ([int](Get-OptionValue $Rest "--padding" "16")) `
                (-not (Has-Flag $Rest "--no-ocr")))
        }
        "observe" {
            $pos = @(Positional-Args $Rest)
            $identifier = Get-OptionValue $Rest "--target"
            if (-not $identifier -and $pos.Count -gt 0) { $identifier = $pos[0] }
            if (-not $identifier) { throw "observe requires an identifier or --target" }
            $open = Open-App $identifier (Has-Flag $Rest "--no-activate")
            $tree = Traverse-Window ([Int64]$open.hwnd) (Has-Flag $Rest "--visible-only") $true (Get-OptionValue $Rest "--view" "control")
            Write-Json ([pscustomobject]@{ target = $open; tree = $tree })
        }
        "uia" {
            $pos = @(Positional-Args $Rest)
            if ($pos.Count -lt 3) { throw "uia requires action hwnd uia_path [value]" }
            $action = $pos[0]
            if ($action -in @("uiaactivate", "activate")) { $action = "activate" }
            elseif ($action -in @("uiaclick", "click")) { $action = "click" }
            elseif ($action -in @("uiapress", "press")) { $action = "press" }
            elseif ($action -in @("uiafocus", "focus")) { $action = "focus" }
            elseif ($action -in @("uiaselect", "select")) { $action = "select" }
            elseif ($action -in @("uiasetvalue", "set_value")) { $action = "set_value" }
            $value = if ($pos.Count -ge 4) { $pos[3] } else { $null }
            Write-Json (Invoke-Uia $action ([Int64]$pos[1]) $pos[2] $value)
        }
        "input" {
            $hwnd = Get-OptionValue $Rest "--hwnd"
            $pos = @(Positional-Args $Rest)
            if ($pos.Count -lt 1) { throw "input requires action [args...]" }
            Write-Json (Send-InputAction -Action $pos[0] -InputArgs @($pos | Select-Object -Skip 1) -Hwnd $hwnd)
        }
        "input-sequence" {
            $hwnd = Get-OptionValue $Rest "--hwnd"
            $json = Get-OptionValue $Rest "--json"
            Write-Json (Send-InputSequence -Json $json -Hwnd $hwnd)
        }
        "ocr" {
            $identifier = Get-OptionValue $Rest "--identifier"
            if (-not $identifier) { $identifier = Get-OptionValue $Rest "--target" }
            Write-Json (Invoke-Ocr (Get-OptionValue $Rest "--image") (Get-OptionValue $Rest "--hwnd") (Get-OptionValue $Rest "--pid") $identifier (Get-OptionValue $Rest "--rect") (Get-OptionValue $Rest "--uia-path"))
        }
        default {
            Fail "usage: WindowsUseSDK.ps1 <list-apps|open|traverse|elements|probe|observe|uia|input|input-sequence|ocr> [args]" 2
        }
    }
} catch {
    Write-Json ([pscustomobject]@{ status = "error"; error = $_.Exception.Message; command = $Command })
    exit 1
}
