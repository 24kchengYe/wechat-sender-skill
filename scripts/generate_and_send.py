"""
Generate and execute WeChat message sending script.
Usage: python generate_and_send.py "联系人名字" "消息内容"
"""

import sys
import subprocess
import os

def to_char_codes(text):
    """Convert string to comma-separated Unicode code points."""
    return ",".join(str(ord(c)) for c in text)

def main():
    if len(sys.argv) < 3:
        print("Usage: python generate_and_send.py <contact_name> <message>")
        print("Example: python generate_and_send.py '张三' '你好'")
        sys.exit(1)

    contact = sys.argv[1]
    message = sys.argv[2]

    print(f"Contact: {contact}")
    print(f"Message: {message}")
    print(f"Contact chars: {to_char_codes(contact)}")
    print(f"Message chars: {to_char_codes(message)}")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    ps_script = os.path.join(script_dir, "send_message.ps1")

    cmd = [
        "powershell", "-STA", "-File", ps_script,
        "-ContactChars", to_char_codes(contact),
        "-MessageChars", to_char_codes(message),
    ]

    print(f"\nExecuting: {' '.join(cmd)}\n")
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(f"STDERR: {result.stderr}", file=sys.stderr)

    sys.exit(result.returncode)

if __name__ == "__main__":
    main()
