import sys

class Logger:
    def __init__(self, debug_mode=False):
        self.debug_mode = debug_mode

    def debug(self, msg):
        if self.debug_mode:
            print(f"[DEBUG] {msg}")

    def info(self, msg):
        print(msg)

    def success(self, msg):
        print(f"✓ {msg}")

    def warning(self, msg):
        print(f"⚠ {msg}")

    def error(self, msg):
        print(f"[ERRO] {msg}")

    def section(self, title):
        print("\n" + "=" * 60)
        print(title)
        print("=" * 60 + "\n")

    def separator(self):
        print("-" * 40)

    def step(self, processo, ano, curr, total):
        print(f"\n[{curr:03d}/{total:03d}]")
        print(f"Processo: {processo}/{ano}")
        self.separator()
