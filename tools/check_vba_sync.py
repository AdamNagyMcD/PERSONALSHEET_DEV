#!/usr/bin/env python3
"""Compare the VBA embedded in Personalsheet.xlsm with the exported sources in vba/.

Why this exists: the workbook is the runnable artifact, vba/*.bas|cls is the git
source of truth. They drift apart silently whenever an import did not run (or was
skipped), and the result on Windows is "Sub or Function not defined" for a
procedure that is clearly present in the repository.

The two bootstrap modules (mod_ResetAndImportVBAFiles, mod_CopyData) are never
imported automatically, so they are reported separately: a difference there means
the workbook still runs an older revision and has to be updated by hand in the
VBA editor.

Usage:
    python3 tools/check_vba_sync.py [--workbook Personalsheet.xlsm] [--vba-dir vba]
                                    [--show-diff]

Exit code 0 = workbook and repository agree, 1 = drift found.
"""

from __future__ import annotations

import argparse
import difflib
import os
import sys

# Same skip list as tools/import_vba_and_repair.ps1 and .cursor/rules.md.
BOOTSTRAP_MODULES = ("mod_ResetAndImportVBAFiles", "mod_CopyData")

HEADER_PREFIXES = ("version ", "begin", "end", "multiuse", "attribute ")


def repair_mojibake(text: str) -> str:
    """olevba decodes the VBA streams with a code page; undo the double encoding."""
    try:
        return text.encode("cp1252").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text


def normalize(text: str):
    """Code lines only, lower case, ASCII-folded - VBA rewrites identifier casing."""
    lines = []
    for line in repair_mojibake(text).replace("\r\n", "\n").split("\n"):
        stripped = line.strip()
        low = stripped.lower()
        if not stripped:
            continue
        if any(low.startswith(prefix) for prefix in HEADER_PREFIXES) and (
                low.startswith("attribute ") or low in ("version 1.0 class", "begin", "end")
                or low.startswith("multiuse")):
            continue
        folded = "".join(ch if ord(ch) < 128 else "?" for ch in low)
        lines.append(" ".join(folded.split()))
    return lines


def extract_workbook_modules(workbook_path: str):
    from oletools.olevba import VBA_Parser

    parser = VBA_Parser(workbook_path)
    modules = {}
    for _, _, vba_filename, code in parser.extract_all_macros():
        name = os.path.splitext(os.path.basename(vba_filename))[0]
        modules[name] = code
    parser.close()
    return modules


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--workbook", default=os.path.join(root, "Personalsheet.xlsm"))
    parser.add_argument("--vba-dir", default=os.path.join(root, "vba"))
    parser.add_argument("--show-diff", action="store_true",
                        help="print a unified diff for every module that differs")
    args = parser.parse_args()

    workbook_modules = extract_workbook_modules(args.workbook)
    repo_modules = {}
    for entry in sorted(os.listdir(args.vba_dir)):
        if entry.lower().endswith((".bas", ".cls")):
            with open(os.path.join(args.vba_dir, entry), encoding="utf-8") as handle:
                repo_modules[os.path.splitext(entry)[0]] = handle.read()

    problems = 0
    bootstrap_drift = []

    missing = sorted(set(repo_modules) - set(workbook_modules))
    for name in missing:
        print("MISSING   %s ist in vba/ vorhanden, fehlt aber in der Arbeitsmappe" % name)
        problems += 1

    # Sheet code modules (Tabelle*) exist in the workbook without an exported file.
    extra = sorted(name for name in set(workbook_modules) - set(repo_modules)
                   if not name.lower().startswith("tabelle"))
    for name in extra:
        print("EXTRA     %s liegt in der Arbeitsmappe, hat aber keine Datei in vba/" % name)
        problems += 1

    for name in sorted(set(repo_modules) & set(workbook_modules)):
        repo_lines = normalize(repo_modules[name])
        wb_lines = normalize(workbook_modules[name])
        if repo_lines == wb_lines:
            continue
        changed = sum(1 for line in difflib.unified_diff(wb_lines, repo_lines, n=0)
                      if line.startswith(("+", "-")) and not line.startswith(("+++", "---")))
        if name in BOOTSTRAP_MODULES:
            bootstrap_drift.append((name, changed))
        else:
            print("DRIFT     %s: %d Zeilen unterschiedlich (Import fehlt?)" % (name, changed))
            problems += 1
        if args.show_diff:
            print("\n".join(difflib.unified_diff(
                wb_lines, repo_lines,
                "workbook/%s" % name, "repo/%s" % name, lineterm="", n=1)))

    for name, changed in bootstrap_drift:
        print("BOOTSTRAP %s: %d Zeilen unterschiedlich - wird nie automatisch importiert, "
              "bitte im VBA-Editor von Hand angleichen" % (name, changed))
        problems += 1

    print("")
    print("Workbook-Komponenten: %d   Dateien in vba/: %d"
          % (len(workbook_modules), len(repo_modules)))
    print("Abweichungen: %d" % problems)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
