"""Set `triggeredAddProgram` in notify.json to true.

This script updates the `notify.json` file located in the same
directory as this script. If the file doesn't exist or is invalid
JSON, a new object will be created. The key `triggeredAddProgram`
will be set to the boolean `true`.
"""

import json
from pathlib import Path


def set_triggered_add_program():
    base = Path(__file__).resolve().parent
    notify_path = base / "notify.json"

    data = {}
    if notify_path.exists():
        try:
            text = notify_path.read_text(encoding="utf-8")
            data = json.loads(text)
            if not isinstance(data, dict):
                data = {}
        except Exception:
            data = {}
    
    data["triggeredAddProgram"] = True

    # Write atomically
    tmp = notify_path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(notify_path)


if __name__ == "__main__":
    set_triggered_add_program()

