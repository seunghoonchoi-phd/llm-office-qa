"""Check PowerPoint, Excel, and Word files for objective Office defects.

This is the user-facing entry point. It dispatches to the format-specific
checkers under tools/.
"""
import argparse
import os
import subprocess
import sys


ROOT = os.path.dirname(os.path.abspath(__file__))
CHECKERS = {
    ".pptx": os.path.join(ROOT, "tools", "check_powerpoint.py"),
    ".xlsx": os.path.join(ROOT, "tools", "check_excel.py"),
    ".docx": os.path.join(ROOT, "tools", "check_word.py"),
}


def main():
    parser = argparse.ArgumentParser(
        description="Check .pptx, .xlsx, and .docx files for objective defects."
    )
    parser.add_argument("files", nargs="+", help="Office files to inspect")
    parser.add_argument("--json", default=None, help="Write JSON report for a single file")
    parser.add_argument("--strict", action="store_true", help="exit nonzero on WARN too")
    parser.add_argument("--quiet", action="store_true", help="suppress human-readable output")
    args = parser.parse_args()

    if args.json and len(args.files) != 1:
        parser.error("--json can only be used with one file")

    final_code = 0
    for path in args.files:
        ext = os.path.splitext(path)[1].lower()
        checker = CHECKERS.get(ext)
        if checker is None:
            print(f"ERROR: unsupported file type: {path}", file=sys.stderr)
            final_code = max(final_code, 2)
            continue

        cmd = [sys.executable, checker, path]
        if args.json:
            cmd += ["--json", args.json]
        if args.strict:
            cmd.append("--strict")
        if args.quiet:
            cmd.append("--quiet")

        result = subprocess.run(cmd)
        if result.returncode == 2:
            final_code = 2
        elif result.returncode != 0 and final_code == 0:
            final_code = result.returncode

    return final_code


if __name__ == "__main__":
    sys.exit(main())
