---
name: wechat-sender
description: |
  Send messages to WeChat contacts via desktop automation on Windows. Use this skill whenever the user asks to "send a WeChat message", "message someone on WeChat", "打开微信发消息", "给某某发微信", "微信发送", or any request involving sending text to a WeChat contact. Requires WeChat desktop client to be running on Windows.
---

# WeChat Desktop Message Sender

Automate sending messages through the WeChat (微信) desktop client on Windows using PowerShell + Win32 API automation.

## Prerequisites

- Windows OS with PowerShell
- WeChat desktop client (`Weixin.exe`) running and logged in
- The target contact must exist in the user's WeChat (by nickname, remark name, or WeChat ID)

## How It Works

The skill uses a three-stage automation approach:

### Stage 1: Activate WeChat Window
- Find the `Weixin` process and get its main window handle
- Bring the window to foreground using `ShowWindow` + `SetForegroundWindow` (Win32 API)
- Wait for the window to fully activate

### Stage 2: Search and Select Contact
- Send `Ctrl+F` to open WeChat's built-in search
- Copy the contact name to clipboard and paste with `Ctrl+V`
- Wait for search results to load (2 seconds recommended)
- Press `Enter` to select the first matching contact and open the chat

### Stage 3: Click Input Area, Paste Message, and Send
- **Critical step**: Use Win32 `SetCursorPos` + `mouse_event` to physically click on the message input area
  - Position: approximately 65% from left, 88% from top of the WeChat window
  - This ensures the input box is focused — `SendKeys` alone does NOT reliably focus the input area
- Re-assert foreground window to prevent focus loss
- Use clipboard (`Clipboard.Clear()` + `Clipboard.SetText()`) with retry logic to avoid clipboard lock errors
- Paste the message with `Ctrl+V`
- Press `Enter` to send

## Key Technical Insights

### Why SendKeys alone fails for message sending
After selecting a contact from search, WeChat's message input area does NOT automatically receive keyboard focus. You MUST use mouse click automation (Win32 `SetCursorPos` + `mouse_event`) to click on the input area before pasting text.

### Clipboard reliability
The Windows clipboard can be locked by other processes (especially WeChat itself after a paste operation). Always:
1. Call `Clipboard.Clear()` before `SetText()`
2. Wrap clipboard operations in a retry loop (up to 5 attempts with 300ms delay)
3. Add a 100ms pause between `Clear()` and `SetText()`

### Unicode text handling
For Chinese characters, use Unicode char arrays to avoid encoding issues in PowerShell:
```powershell
# Instead of literal strings that may break in different encodings:
$text = [string]::new([char[]]@(20320, 22909))  # "你好"
```
To get Unicode code points for any text, use Python:
```python
print([ord(c) for c in "你好"])  # [20320, 22909]
```

### Window positioning
Calculate click coordinates relative to the WeChat window rect:
```powershell
$clickX = $rect.Left + [int]($winWidth * 0.65)   # Horizontal center-right of chat area
$clickY = $rect.Bottom - [int]($winHeight * 0.12) # Near bottom where input box sits
```

### Timing
Recommended delays between operations:
| Operation | Delay |
|-----------|-------|
| After `ShowWindow` / `SetForegroundWindow` | 300ms + 1000ms |
| After `Ctrl+F` (open search) | 600ms |
| After pasting contact name | 2000ms (search needs time) |
| After `Enter` to select contact | 2500ms (chat needs to load) |
| After clicking input area | 1000ms |
| After pasting message | 800ms |

## Implementation Template

Use the helper script at `~/.claude/skills/wechat-sender/scripts/send_message.ps1`:

```bash
# Step 1: Generate Unicode char arrays for contact name and message using Python
python -c "print([ord(c) for c in '联系人名字'])"
python -c "print([ord(c) for c in '消息内容'])"

# Step 2: Update the script with the char arrays and run
powershell -STA -File ~/.claude/skills/wechat-sender/scripts/send_message.ps1
```

Or dynamically generate and run:

```bash
# Generate the PowerShell script with the correct Unicode values
python ~/.claude/skills/wechat-sender/scripts/generate_and_send.py "联系人名字" "消息内容"
```

## Workflow for Claude

When a user asks to send a WeChat message:

1. **Extract** the contact name and message from the user's request
2. **Convert** both strings to Unicode char arrays using Python
3. **Generate** a PowerShell script using the template with the correct char arrays
4. **Execute** the script with `powershell -STA -File <script>`
5. **Verify** success — offer to take a screenshot if the user wants confirmation

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "WeChat not found" | Ensure `Weixin.exe` is running. Check with `Get-Process -Name Weixin` |
| Search finds wrong contact | Use the exact remark name or nickname. Try a more specific search term |
| Message not sent (input area not focused) | Adjust click coordinates — the input area position varies by window size |
| Clipboard error | The retry loop should handle this. If persistent, close other clipboard-heavy apps |
| Window not coming to foreground | WeChat may be minimized to tray. Click the tray icon manually first |
