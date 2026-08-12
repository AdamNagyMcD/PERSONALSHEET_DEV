#!/usr/bin/env python3
"""Find VBA procedures in vba/ that nothing references any more.

VBA is full of indirect calls, so a plain "name not found" search is dangerous.
This script counts a procedure as referenced when its name appears in any of:

  * VBA code of any module (outside its own declaration line)
  * any VBA string literal        -> Shape.OnAction, Application.Run, OnKey, Evaluate
  * a worksheet formula inside Personalsheet.xlsm  -> user defined functions
  * the documentation (README.md, SPEC.md, TEST_CASES.md, docs/, AGENTS.md)
  * the workbook/worksheet event handler naming scheme (Workbook_*, Worksheet_*)
  * tools/*.ps1 helper scripts

Everything it prints is a *candidate*: read the code before deleting. The point is
to shrink the manual search space, not to automate deletion.

Usage:
    python3 tools/vba_deadcode.py [--include-private] [--show-refs NAME]
"""

from __future__ import annotations

import argparse
import glob
import os
import re
import sys

PROC_RE = re.compile(
    r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE,
)
IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
EVENT_PREFIXES = ("workbook_", "worksheet_", "auto_", "app_")


def strip_line(line: str):
    """Return (code_without_strings_and_comment, [string literals])."""
    out, literals, current = [], [], []
    in_string = False
    i = 0
    while i < len(line):
        ch = line[i]
        if in_string:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    current.append('"')
                    i += 2
                    continue
                in_string = False
                literals.append("".join(current))
                current = []
            else:
                current.append(ch)
            i += 1
            continue
        if ch == '"':
            in_string = True
            i += 1
            continue
        if ch == "'":
            break
        out.append(ch)
        i += 1
    return "".join(out), literals


def load_sources(vba_dir):
    modules = {}
    for path in sorted(glob.glob(os.path.join(vba_dir, "*.bas"))
                       + glob.glob(os.path.join(vba_dir, "*.cls"))):
        with open(path, encoding="utf-8") as handle:
            modules[os.path.splitext(os.path.basename(path))[0]] = \
                handle.read().replace("\r\n", "\n").split("\n")
    return modules


def collect_procs(modules):
    procs = []
    for module, lines in modules.items():
        for number, line in enumerate(lines, start=1):
            code, _ = strip_line(line)
            m = PROC_RE.match(code)
            if m:
                procs.append({
                    "name": m.group(3),
                    "module": module,
                    "kind": m.group(2).split()[0].capitalize(),
                    "visibility": (m.group(1) or "Public").capitalize(),
                    "line": number,
                    "is_class": False,
                })
    return procs


def workbook_formula_text(workbook_path):
    if not os.path.exists(workbook_path):
        return ""
    import zipfile
    chunks = []
    with zipfile.ZipFile(workbook_path) as archive:
        for name in archive.namelist():
            if name.startswith("xl/") and name.endswith(".xml"):
                chunks.append(archive.read(name).decode("utf-8", "replace"))
    return "\n".join(chunks)


def doc_text(root):
    chunks = []
    patterns = ["*.md", "docs/*.md", "docs/*.html", "tools/*.ps1", ".cursor/*.md"]
    for pattern in patterns:
        for path in glob.glob(os.path.join(root, pattern)):
            with open(path, encoding="utf-8", errors="replace") as handle:
                chunks.append(handle.read())
    return "\n".join(chunks)


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--vba-dir", default=os.path.join(root, "vba"))
    parser.add_argument("--workbook", default=os.path.join(root, "Personalsheet.xlsm"))
    parser.add_argument("--include-private", action="store_true",
                        help="also list unreferenced Private procedures")
    parser.add_argument("--show-refs", metavar="NAME",
                        help="print every place that mentions NAME")
    args = parser.parse_args()

    modules = load_sources(args.vba_dir)
    procs = collect_procs(modules)

    # index every identifier occurrence and every string literal, per module
    code_hits = {}          # lower name -> list of "module:line"
    literal_hits = {}
    for module, lines in modules.items():
        for number, line in enumerate(lines, start=1):
            code, literals = strip_line(line)
            if PROC_RE.match(code):
                # keep the parameter list, drop the declared name itself
                code = code[code.find("(") if "(" in code else len(code):]
            for m in IDENT_RE.finditer(code):
                code_hits.setdefault(m.group(0).lower(), []).append(
                    "%s:%d" % (module, number))
            for literal in literals:
                for m in IDENT_RE.finditer(literal):
                    literal_hits.setdefault(m.group(0).lower(), []).append(
                        "%s:%d (Text)" % (module, number))

    formulas = workbook_formula_text(args.workbook).lower()
    docs = doc_text(root).lower()

    if args.show_refs:
        key = args.show_refs.lower()
        print("Code   :", code_hits.get(key, []))
        print("Texte  :", literal_hits.get(key, []))
        print("Formeln:", formulas.count(key))
        print("Doku   :", docs.count(key))
        return 0

    unreferenced = []
    for proc in procs:
        key = proc["name"].lower()
        if key.startswith(EVENT_PREFIXES):
            continue
        if not args.include_private and proc["visibility"] == "Private":
            pass  # private procedures are checked too, they are the safest to remove
        in_code = len(code_hits.get(key, []))
        in_literals = len(literal_hits.get(key, []))
        in_formulas = formulas.count(key)
        in_docs = docs.count(key)
        if in_code == 0 and in_literals == 0 and in_formulas == 0 and in_docs == 0:
            unreferenced.append(proc)

    print("Prozeduren gesamt: %d   ohne jede Referenz: %d"
          % (len(procs), len(unreferenced)))
    print("")
    if unreferenced:
        print("%-44s %-26s %-9s %-8s %s"
              % ("Prozedur", "Modul", "Art", "Sichtbar", "Zeile"))
        for proc in sorted(unreferenced, key=lambda p: (p["module"], p["line"])):
            print("%-44s %-26s %-9s %-8s %d"
                  % (proc["name"], proc["module"], proc["kind"],
                     proc["visibility"], proc["line"]))
        print("")
        print("Jeder Treffer ist ein KANDIDAT - vor dem Loeschen den Code lesen.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
