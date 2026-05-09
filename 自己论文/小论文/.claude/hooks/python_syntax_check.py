import json
import py_compile
import sys
from pathlib import Path


def main() -> int:
    try:
        payload = json.load(sys.stdin)
    except json.JSONDecodeError:
        return 0

    tool_input = payload.get("tool_input", {})
    file_path = tool_input.get("file_path")
    if not file_path:
        return 0

    path = Path(file_path)
    if path.suffix != ".py" or not path.exists():
        return 0

    try:
        py_compile.compile(str(path), doraise=True)
    except py_compile.PyCompileError as exc:
        message = str(exc)
        print(json.dumps({
            "systemMessage": f"Python 语法检查失败: {path.name}\n{message}",
            "hookSpecificOutput": {
                "hookEventName": "PostToolUse",
                "additionalContext": f"Python syntax check failed for {path}: {message}"
            }
        }, ensure_ascii=False))
        return 0

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
