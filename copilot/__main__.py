"""Command-line entry point.

    python -m copilot login        # interactive sign-in, persists the session
    python -m copilot ask "hi"     # one-shot completion via the pure-HTTP driver
"""

import sys


def main(argv) -> int:
    cmd = argv[0] if argv else "ask"
    if cmd == "login":
        # The browser is used only for interactive sign-in / token capture.
        from .browser import BrowserCopilot

        BrowserCopilot(headless=False).login()
        return 0
    if cmd == "ask":
        prompt = " ".join(argv[1:]) or "Hello!"
        from .client import CopilotClient

        for chunk in CopilotClient().stream(prompt):
            if isinstance(chunk, str):
                print(chunk, end="", flush=True)
        print()
        return 0
    print(f"Unknown command: {cmd!r}. Use 'login' or 'ask <prompt>'.")
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
