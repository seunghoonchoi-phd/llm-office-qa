"""
qa_hook.py — PostToolUse dispatcher: auto-lint freshly generated Office files.

Wired as a PostToolUse hook on Bash/PowerShell. It reads the hook JSON on stdin,
looks for .pptx/.xlsx/.docx paths in the command AND in the tool output (scripts
usually print "wrote <path>"), lints any that were just created/modified, and:
  - if any OBJECTIVE ERROR  -> prints details to stderr and exits 2 (fed back to
    the model so it fixes before delivering),
  - if only WARN findings   -> prints a short note to stdout, exits 0 (non-blocking),
  - clean / nothing to lint  -> exits 0 silently.

Never crashes the workflow: any internal error -> exit 0.
"""
import json
import os
import re
import shutil
import subprocess
import sys
import time

TOOLDIR = os.path.dirname(os.path.abspath(__file__))
# Use the interpreter running this hook; fall back to whatever python is on PATH.
PYEXE = sys.executable or shutil.which("python") or shutil.which("py") or "python"
LINTERS = {".pptx": "pptx_lint.py", ".xlsx": "xlsx_lint.py", ".docx": "docx_lint.py"}
RECENT_SEC = 180          # only lint files modified within this window
SELF_SKIP = re.compile(r"_lint\.py|qa_hook|make_test_", re.I)
PATH_RE = re.compile(r"""["']([^"']+\.(?:pptx|xlsx|docx))["']|(\S+\.(?:pptx|xlsx|docx))""", re.I)


def gather_text(data):
    cmd = ""
    ti = data.get("tool_input") or {}
    if isinstance(ti, dict):
        cmd = ti.get("command") or ti.get("file_path") or ""
    parts = [cmd]
    resp = data.get("tool_response")
    if isinstance(resp, str):
        parts.append(resp)
    elif isinstance(resp, dict):
        for k in ("stdout", "output", "stderr", "content"):
            v = resp.get(k)
            if isinstance(v, str):
                parts.append(v)
            elif isinstance(v, list):
                for it in v:
                    if isinstance(it, dict) and isinstance(it.get("text"), str):
                        parts.append(it["text"])
                    elif isinstance(it, str):
                        parts.append(it)
    return cmd, "\n".join(parts)


def find_paths(text, cwd):
    out = set()
    for m in PATH_RE.finditer(text):
        p = (m.group(1) or m.group(2) or "").strip().strip("'\"")
        if not p:
            continue
        if not os.path.isabs(p):
            p = os.path.join(cwd, p)
        p = os.path.normpath(p)
        try:
            if os.path.isfile(p) and (time.time() - os.path.getmtime(p)) < RECENT_SEC:
                out.add(p)
        except OSError:
            pass
    return out


def main():
    raw = sys.stdin.buffer.read().decode("utf-8", "replace")  # locale-independent
    if not raw.strip():
        return 0
    data = json.loads(raw)
    cmd, text = gather_text(data)
    if SELF_SKIP.search(cmd):          # don't re-lint when we're running linters/tests
        return 0
    cwd = data.get("cwd") or os.getcwd()
    paths = find_paths(text, cwd)
    if not paths:
        return 0

    errors, warns = [], []
    for p in sorted(paths):
        linter = LINTERS.get(os.path.splitext(p)[1].lower())
        if not linter:
            continue
        try:
            r = subprocess.run([PYEXE, os.path.join(TOOLDIR, linter), p],
                               capture_output=True, text=True, timeout=90,
                               encoding="utf-8", errors="replace")
        except Exception:
            continue
        body = (r.stdout or "").strip()
        if r.returncode == 1:                       # linter exits 1 only on ERROR
            errors.append((p, body))
        elif "clean --" not in body and body:       # findings but no ERROR
            warns.append((p, body))

    if errors:
        sys.stderr.write("\n[QA hook] OBJECTIVE DEFECTS in a generated file — fix before delivering:\n")
        for p, body in errors:
            sys.stderr.write(f"\n### {p}\n{body}\n")
        if warns:
            for p, body in warns:
                sys.stderr.write(f"\n### {p} (warnings)\n{body}\n")
        return 2
    if warns:
        sys.stdout.write("\n[QA hook] Lint warnings (verify, non-blocking):\n")
        for p, body in warns:
            sys.stdout.write(f"\n### {p}\n{body}\n")
        return 0
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        sys.exit(0)
