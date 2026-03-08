\![Visitors](https://visitor-badge.laobi.icu/badge?page_id=24kchengYe.wechat-sender-skill)
# WeChat Desktop Message Sender

A Claude Code skill that automates sending messages through the WeChat (微信) desktop client on Windows.

## How It Works

This skill uses **PowerShell + Win32 API** to automate the WeChat desktop client through three stages:

1. **Activate** — Find the WeChat process, bring its window to the foreground
2. **Search** — Open WeChat's search (`Ctrl+F`), paste the contact name, press Enter
3. **Send** — Click the message input area (mouse automation), paste the message, press Enter

### Key Insight

After selecting a contact from search, WeChat's message input area does **NOT** automatically receive keyboard focus. The critical fix is using Win32 `SetCursorPos` + `mouse_event` to physically click on the input area before pasting text.

## Prerequisites

- Windows OS with PowerShell
- WeChat desktop client (`Weixin.exe`) running and logged in
- Target contact must exist in your WeChat contacts

## Installation

```bash
# Clone to your Claude Code skills directory
git clone https://github.com/24kchengYe/wechat-sender-skill ~/.claude/skills/wechat-sender
```

## Usage

### Via Claude Code (Recommended)

Just tell Claude Code in natural language:

- "Send a WeChat message to 张三 saying 你好"
- "给张三发微信说你好"
- "打开微信给某某发消息"

Claude Code will automatically trigger this skill and handle the automation.

### Via Command Line

```bash
# Using the Python wrapper (handles Unicode conversion automatically)
python ~/.claude/skills/wechat-sender/scripts/generate_and_send.py "联系人名字" "消息内容"

# Using PowerShell directly (requires Unicode char codes)
# First get char codes: python -c "print(','.join(str(ord(c)) for c in '你好'))"
powershell -STA -File ~/.claude/skills/wechat-sender/scripts/send_message.ps1 -ContactChars "20320,22909" -MessageChars "20320,30495,22909,30475"
```

## Technical Details

### Unicode Handling

Chinese characters are passed as Unicode code point arrays to avoid PowerShell encoding issues:

```python
# Convert text to char codes
print([ord(c) for c in "你好"])  # [20320, 22909]
```

### Clipboard Safety

The Windows clipboard can be locked by other processes. The script uses a retry loop with `Clipboard.Clear()` before each `SetText()` call.

### Window Positioning

Click coordinates are calculated relative to the WeChat window:
- **X**: 65% from left edge (center of chat area)
- **Y**: 88% from top (where the input box sits)

## Troubleshooting

| Problem | Solution |
|---------|----------|
| WeChat not found | Ensure `Weixin.exe` is running |
| Wrong contact selected | Use the exact remark name or nickname |
| Message not sent | Input area may not be focused — adjust click coordinates |
| Clipboard error | Close other clipboard-heavy apps, the retry loop should handle most cases |

## License

MIT
