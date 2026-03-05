# WeChat Desktop Message Sender
# Usage: powershell -STA -File send_message.ps1 -ContactChars "20320,22909" -MessageChars "20320,30495,22909,30475"
# Char arrays are Unicode code points for Chinese text

param(
    [string]$ContactChars,
    [string]$MessageChars
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WeChatWin32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    public static extern void SetCursorPos(int x, int y);
    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left; public int Top; public int Right; public int Bottom;
    }

    public static void ClickAt(int x, int y) {
        SetCursorPos(x, y);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
}
"@

function Set-ClipboardSafe($text) {
    for ($i = 0; $i -lt 5; $i++) {
        try {
            [System.Windows.Forms.Clipboard]::Clear()
            Start-Sleep -Milliseconds 100
            [System.Windows.Forms.Clipboard]::SetText($text)
            return $true
        } catch {
            Start-Sleep -Milliseconds 300
        }
    }
    return $false
}

function ConvertTo-StringFromCharCodes($charCodes) {
    $codes = $charCodes -split "," | ForEach-Object { [int]$_.Trim() }
    return [string]::new([char[]]$codes)
}

# Parse parameters
$contactName = ConvertTo-StringFromCharCodes $ContactChars
$message = ConvertTo-StringFromCharCodes $MessageChars

Write-Output "Contact: $contactName"
Write-Output "Message: $message"

# Find WeChat process
$procs = Get-Process -Name Weixin -ErrorAction SilentlyContinue
$hwnd = [IntPtr]::Zero
foreach ($p in $procs) {
    if ($p.MainWindowHandle -ne [IntPtr]::Zero) {
        $hwnd = $p.MainWindowHandle
        break
    }
}

if ($hwnd -eq [IntPtr]::Zero) {
    Write-Error "WeChat (Weixin) is not running or has no visible window."
    exit 1
}

# Stage 1: Activate WeChat window
[WeChatWin32]::ShowWindow($hwnd, 9)
Start-Sleep -Milliseconds 300
[WeChatWin32]::SetForegroundWindow($hwnd)
Start-Sleep -Milliseconds 1000

# Get window dimensions
$rect = New-Object WeChatWin32+RECT
[WeChatWin32]::GetWindowRect($hwnd, [ref]$rect)
$winW = $rect.Right - $rect.Left
$winH = $rect.Bottom - $rect.Top
Write-Output "Window: ${winW}x${winH} at ($($rect.Left),$($rect.Top))"

# Stage 2: Search and select contact
[System.Windows.Forms.SendKeys]::SendWait("^f")
Start-Sleep -Milliseconds 600

if (-not (Set-ClipboardSafe $contactName)) {
    Write-Error "Failed to set clipboard for contact name"
    exit 1
}
Start-Sleep -Milliseconds 300
[System.Windows.Forms.SendKeys]::SendWait("^v")
Start-Sleep -Milliseconds 2000

[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Write-Output "Contact selected, waiting for chat to load..."
Start-Sleep -Milliseconds 2500

# Stage 3: Click input area, paste message, send
$clickX = $rect.Left + [int]($winW * 0.65)
$clickY = $rect.Bottom - [int]($winH * 0.12)
Write-Output "Clicking input area at: ($clickX, $clickY)"
[WeChatWin32]::ClickAt($clickX, $clickY)
Start-Sleep -Milliseconds 1000

[WeChatWin32]::SetForegroundWindow($hwnd)
Start-Sleep -Milliseconds 500

if (-not (Set-ClipboardSafe $message)) {
    Write-Error "Failed to set clipboard for message"
    exit 1
}
Start-Sleep -Milliseconds 500
[System.Windows.Forms.SendKeys]::SendWait("^v")
Write-Output "Message pasted, sending..."
Start-Sleep -Milliseconds 800

[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Milliseconds 500

Write-Output "Message sent successfully!"
