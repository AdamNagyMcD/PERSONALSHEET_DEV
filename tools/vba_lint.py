#!/usr/bin/env python3
"""Static checker for the exported VBA sources in vba/.

Closest available proxy for "Debug > VBAProject kompilieren" on a machine without
Excel. It catches the compile errors that have actually hurt this project:

  * "Sub or Function not defined"  -> call to a PID_* name that exists nowhere
  * "Sub or Function not defined"  -> cross-module call to a Private procedure
  * "Ambiguous name detected"      -> same Public name in two standard modules
  * "Label not defined"            -> GoTo/Resume target missing in the procedure
  * unbalanced Sub/Function/If/For/Do/With/Select blocks
  * missing Option Explicit, VB_Name that does not match the file name
  * formulas that are not Excel 2016 compatible (XLOOKUP, LET, FILTER, ...)

Usage:
    python3 tools/vba_lint.py [--vba-dir vba] [--quiet]

Exit code 0 = no errors (warnings allowed), 1 = at least one error.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass, field

# Standard modules that live in the workbook and are never callable by name from
# other modules: their procedures are class members / event handlers.
CLASS_EXTENSIONS = (".cls",)

VBA_KEYWORDS = {
    "abs", "and", "any", "as", "boolean", "byref", "byte", "byval", "call", "case",
    "cbool", "cbyte", "ccur", "cdate", "cdbl", "cdec", "chdir", "chdrive", "chr",
    "chrw", "cint", "circle", "clng", "close", "const", "cos", "createobject",
    "csng", "cstr", "curdir", "currency", "cvar", "cverr", "date", "dateadd",
    "datediff", "datepart", "dateserial", "datevalue", "day", "ddb", "debug",
    "decimal", "declare", "defbool", "dim", "dir", "do", "doevents", "double",
    "each", "else", "elseif", "empty", "end", "endif", "environ", "eof", "eqv",
    "erase", "err", "error", "exit", "exp", "explicit", "false", "fileattr",
    "filecopy", "filedatetime", "filelen", "fix", "for", "format", "freefile",
    "function", "get", "getattr", "getobject", "gosub", "goto", "hex", "hour",
    "if", "iif", "imp", "in", "input", "instr", "instrrev", "int", "integer",
    "is", "isarray", "isdate", "isempty", "iserror", "ismissing", "isnull",
    "isnumeric", "isobject", "join", "kill", "lbound", "lcase", "left", "len",
    "let", "like", "line", "loc", "lock", "lof", "log", "long", "loop", "ltrim",
    "me", "mid", "minute", "mkdir", "mod", "month", "monthname", "msgbox", "new",
    "next", "not", "nothing", "now", "null", "object", "oct", "on", "open",
    "option", "optional", "or", "paramarray", "preserve", "print", "private",
    "property", "public", "put", "raiseevent", "randomize", "redim", "rem",
    "replace", "reset", "resume", "return", "rgb", "right", "rmdir", "rnd",
    "round", "rtrim", "second", "seek", "select", "set", "setattr", "sgn",
    "shell", "single", "sin", "space", "spc", "split", "sqr", "static", "step",
    "stop", "str", "strcomp", "strconv", "string", "strreverse", "sub", "switch",
    "tab", "tan", "text", "then", "time", "timer", "timeserial", "timevalue",
    "to", "trim", "true", "typename", "typeof", "ubound", "ucase", "unlock",
    "until", "val", "variant", "vartype", "vbcrlf", "vblf", "vbcr", "vbtab",
    "wend", "weekday", "while", "with", "write", "xor", "year",
}

# Object model members that only exist in Microsoft 365. Early bound
# (Application.X = ...) they are a *compile* error in Excel 2016, which On Error
# cannot catch - use a late bound Object variable if the call is really wanted.
POST_2016_APPLICATION_MEMBERS = (
    "formatstalevalues", "autosaveon", "checkperformance",
)
POST_2016_WORKSHEETFUNCTION_MEMBERS = (
    "textjoin", "concat", "ifs", "maxifs", "minifs", "switch", "xlookup", "xmatch",
    "unique", "sortby", "sequence", "randarray", "textbefore", "textafter",
    "textsplit", "vstack", "hstack", "lambda",
)
APPLICATION_MEMBER_RE = re.compile(r"\bApplication\s*\.\s*([A-Za-z_]\w*)", re.IGNORECASE)
WORKSHEETFUNCTION_MEMBER_RE = re.compile(
    r"\bWorksheetFunction\s*\.\s*([A-Za-z_]\w*)", re.IGNORECASE)

# Worksheet functions that Excel 2016 (perpetual) does not know. See .cursor/rules.md.
FORBIDDEN_FORMULA_TOKENS = (
    "XLOOKUP", "XVERWEIS", "LET(", "FILTER(", "UNIQUE(", "SEQUENCE(", "SORTBY(",
    "RANDARRAY(", "TEXTSPLIT(", "TEXTBEFORE(", "TEXTAFTER(", "VSTACK(", "HSTACK(",
    "XMATCH(", "TOCOL(", "TOROW(", "LAMBDA(",
)

# Identifiers that are always available: VBA language, VBA runtime library and the
# Excel object model globals. Everything starting with xl/mso/vb is treated as an
# enum member, everything after a dot is a member access and never checked.
GLOBAL_IDENTIFIERS = {
    "application", "thisworkbook", "activesheet", "activeworkbook", "activecell",
    "activewindow", "selection", "workbooks", "worksheets", "sheets", "range",
    "cells", "columns", "rows", "err", "debug", "me", "nothing", "true", "false",
    "empty", "null", "vba", "excel", "worksheetfunction", "commandbars", "names",
    "shell", "environ", "clipboard", "target", "sh", "cancel", "saveasui",
    # types used in As clauses
    "worksheet", "workbook", "variant", "string", "long", "integer", "double",
    "single", "boolean", "byte", "date", "currency", "object", "collection",
    "shape", "shapes", "chart", "chartobject", "hyperlink", "comment", "name",
    "validation", "listobject", "pivottable", "window",
    # frequently used runtime functions and constants
    "abs", "array", "asc", "atn", "cbool", "cbyte", "ccur", "cdate", "cdbl",
    "cdec", "choose", "chr", "chrw", "cint", "clng", "cos", "createobject",
    "csng", "cstr", "cvar", "cverr", "dateadd", "datediff", "datepart",
    "dateserial", "datevalue", "day", "dir", "doevents", "eof", "error", "exp",
    "fileattr", "filecopy", "filedatetime", "filelen", "fix", "format", "freefile",
    "getattr", "getobject", "hex", "hour", "iif", "inputbox", "instr", "instrrev",
    "int", "isarray", "isdate", "isempty", "iserror", "ismissing", "isnull",
    "isnumeric", "isobject", "join", "kill", "lbound", "lcase", "left", "len",
    "loc", "lof", "log", "ltrim", "mid", "minute", "mkdir", "month", "monthname",
    "msgbox", "now", "oct", "randomize", "replace", "rgb", "right", "rmdir",
    "rnd", "round", "rtrim", "second", "seek", "setattr", "sgn", "sin",
    "space", "split", "sqr", "str", "strcomp", "strconv", "strreverse",
    "tan", "time", "timer", "timeserial", "timevalue", "trim",
    "typename", "ubound", "ucase", "val", "vartype", "weekday", "year",
    # Application methods that may be called unqualified
    "intersect", "union", "evaluate", "volatile", "caller", "run", "calculate",
    "goto", "wait", "inputbox", "index", "match", "transpose",
    # keywords of the Open / Print / Line Input statements
    "output", "input", "append", "binary", "random", "read", "write", "lock",
    "shared", "access", "step", "spc", "tab",
}
GLOBAL_PREFIXES = ("xl", "mso", "vb", "wd", "ol")

PROC_RE = re.compile(
    r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE,
)
END_PROC_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
LABEL_RE = re.compile(r"^\s*([A-Za-z_][A-Za-z0-9_]*)\s*:\s*$")
GOTO_RE = re.compile(r"\b(?:GoTo|GoSub|Resume)\s+([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)
CONST_RE = re.compile(
    r"^\s*(?:(Public|Private|Global)\s+)?Const\s+(.+)$", re.IGNORECASE)
DIM_RE = re.compile(
    r"^\s*(?:Dim|Private|Public|Global|Static|ReDim)\s+(?:Preserve\s+)?(.+)$",
    re.IGNORECASE,
)
IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
ATTRIBUTE_NAME_RE = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.IGNORECASE)


@dataclass
class Proc:
    name: str
    module: str
    kind: str
    visibility: str
    line: int


@dataclass
class Module:
    name: str
    path: str
    is_class: bool
    lines: list = field(default_factory=list)      # raw physical lines
    logical: list = field(default_factory=list)    # (line_no, code_without_strings_comments)


@dataclass
class Finding:
    level: str      # "ERROR" | "WARN"
    module: str
    line: int
    message: str


def strip_code(line: str):
    """Return (code, literals): line with string literals blanked and comment removed."""
    out = []
    literals = []
    in_string = False
    current = []
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
                out.append('"')
            else:
                current.append(ch)
            i += 1
            continue
        if ch == '"':
            in_string = True
            out.append('"')
            i += 1
            continue
        if ch == "'":
            break
        if line[i:i + 4].lower() == "rem " and (i == 0 or not line[i - 1].isalnum()):
            break
        out.append(ch)
        i += 1
    return "".join(out), literals


def join_continuations(raw_lines):
    """Yield (first_line_no, joined_code, joined_literals) for logical lines."""
    buffer_code = ""
    buffer_literals = []
    start_no = None
    for idx, raw in enumerate(raw_lines, start=1):
        code, literals = strip_code(raw)
        if start_no is None:
            start_no = idx
        buffer_literals.extend(literals)
        stripped = code.rstrip()
        if stripped.endswith("_") and (len(stripped) == 1 or stripped[-2].isspace()):
            buffer_code += stripped[:-1]
            continue
        buffer_code += code
        yield start_no, buffer_code, buffer_literals
        buffer_code = ""
        buffer_literals = []
        start_no = None
    if start_no is not None:
        yield start_no, buffer_code, buffer_literals


def split_declared_names(decl: str):
    """Extract declared identifiers from a Dim/Const/parameter list."""
    names = []
    depth = 0
    token = ""
    parts = []
    for ch in decl:
        if ch in "([":
            depth += 1
        elif ch in ")]":
            depth -= 1
        if ch == "," and depth == 0:
            parts.append(token)
            token = ""
        else:
            token += ch
    parts.append(token)
    for part in parts:
        part = part.strip()
        if not part:
            continue
        part = re.split(r"\bAs\b|=", part, maxsplit=1, flags=re.IGNORECASE)[0]
        part = part.strip()
        part = re.sub(r"\(.*\)$", "", part).strip()
        part = re.sub(r"^(?:ByVal|ByRef|Optional|ParamArray)\s+", "", part,
                      flags=re.IGNORECASE).strip()
        # repeat, a parameter can be "Optional ByVal x"
        while re.match(r"^(?:ByVal|ByRef|Optional|ParamArray)\s+", part, re.IGNORECASE):
            part = re.sub(r"^(?:ByVal|ByRef|Optional|ParamArray)\s+", "", part,
                          flags=re.IGNORECASE).strip()
        m = IDENT_RE.match(part)
        if m and m.group(0).lower() not in VBA_KEYWORDS:
            names.append(m.group(0))
    return names


def load_modules(vba_dir: str):
    modules = []
    for entry in sorted(os.listdir(vba_dir)):
        if not entry.lower().endswith((".bas", ".cls")):
            continue
        path = os.path.join(vba_dir, entry)
        with open(path, encoding="utf-8") as handle:
            raw = handle.read()
        lines = raw.replace("\r\n", "\n").split("\n")
        mod = Module(
            name=os.path.splitext(entry)[0],
            path=path,
            is_class=entry.lower().endswith(CLASS_EXTENSIONS),
            lines=lines,
        )
        mod.logical = list(join_continuations(lines))
        modules.append(mod)
    return modules


def collect_symbols(modules):
    procs = {}                 # lower name -> list[Proc]
    module_procs = {}          # module -> set(lower name)
    declared = set()           # every declared identifier, lowercase
    for mod in modules:
        module_procs[mod.name] = set()
        for line_no, code, _ in mod.logical:
            m = PROC_RE.match(code)
            if m:
                visibility = (m.group(1) or "Public").capitalize()
                proc = Proc(name=m.group(3), module=mod.name,
                            kind=m.group(2).split()[0].capitalize(),
                            visibility=visibility, line=line_no)
                procs.setdefault(proc.name.lower(), []).append(proc)
                module_procs[mod.name].add(proc.name.lower())
                declared.add(proc.name.lower())
                params = code[code.find("(") + 1:code.rfind(")")] if "(" in code else ""
                for name in split_declared_names(params):
                    declared.add(name.lower())
                continue
            cm = CONST_RE.match(code)
            if cm:
                for name in split_declared_names(cm.group(2)):
                    declared.add(name.lower())
                continue
            dm = DIM_RE.match(code)
            if dm and not PROC_RE.match(code):
                head = code.strip().split(None, 1)[0].lower()
                if head in {"dim", "private", "public", "global", "static", "redim"}:
                    for name in split_declared_names(dm.group(1)):
                        declared.add(name.lower())
    return procs, module_procs, declared


BLOCK_OPENERS = [
    (re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?Sub\s+", re.I), "Sub"),
    (re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?Function\s+", re.I), "Function"),
    (re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?Property\s+(?:Get|Let|Set)\s+", re.I), "Property"),
    (re.compile(r"^\s*With\b", re.I), "With"),
    (re.compile(r"^\s*Select\s+Case\b", re.I), "Select"),
    (re.compile(r"^\s*(?:For|For\s+Each)\b", re.I), "For"),
    (re.compile(r"^\s*Do\b", re.I), "Do"),
    (re.compile(r"^\s*While\b", re.I), "While"),
]
BLOCK_CLOSERS = [
    (re.compile(r"^\s*End\s+Sub\b", re.I), "Sub"),
    (re.compile(r"^\s*End\s+Function\b", re.I), "Function"),
    (re.compile(r"^\s*End\s+Property\b", re.I), "Property"),
    (re.compile(r"^\s*End\s+With\b", re.I), "With"),
    (re.compile(r"^\s*End\s+Select\b", re.I), "Select"),
    (re.compile(r"^\s*Next\b", re.I), "For"),
    (re.compile(r"^\s*Loop\b", re.I), "Do"),
    (re.compile(r"^\s*Wend\b", re.I), "While"),
    (re.compile(r"^\s*End\s+If\b", re.I), "If"),
]
IF_BLOCK_RE = re.compile(r"^\s*(?:\}|)?\s*If\b.*\bThen\s*$", re.IGNORECASE)
ELSE_RE = re.compile(r"^\s*(?:Else|ElseIf\b.*\bThen)\s*$", re.IGNORECASE)


def check_blocks(mod: Module, findings):
    stack = []
    for line_no, code, _ in mod.logical:
        stripped = code.strip()
        if not stripped:
            continue
        if re.match(r"^\s*Exit\s+(Sub|Function|Property|For|Do)\b", stripped, re.I):
            continue
        if re.match(r"^\s*(?:Declare|Const|Dim)\b", stripped, re.I):
            continue
        if IF_BLOCK_RE.match(stripped):
            stack.append(("If", line_no))
            continue
        if ELSE_RE.match(stripped):
            continue
        opened = False
        for pattern, kind in BLOCK_OPENERS:
            if pattern.match(stripped):
                # single-line "Do ... Loop" is not possible, but "Do While x" is an opener
                stack.append((kind, line_no))
                opened = True
                break
        if opened:
            continue
        for pattern, kind in BLOCK_CLOSERS:
            if pattern.match(stripped):
                if not stack:
                    findings.append(Finding("ERROR", mod.name, line_no,
                                            "Block-Ende ohne Anfang: %s" % stripped))
                elif stack[-1][0] != kind:
                    findings.append(Finding(
                        "ERROR", mod.name, line_no,
                        "Block-Ende passt nicht: '%s' schliesst %s aus Zeile %d"
                        % (stripped, stack[-1][0], stack[-1][1])))
                    stack.pop()
                else:
                    stack.pop()
                break
    for kind, line_no in stack:
        findings.append(Finding("ERROR", mod.name, line_no,
                                "Block '%s' wird nicht geschlossen" % kind))


def check_labels(mod: Module, findings):
    current_proc = None
    labels = set()
    targets = []
    for line_no, code, _ in mod.logical:
        if PROC_RE.match(code):
            current_proc = code.strip()
            labels = set()
            targets = []
            continue
        if END_PROC_RE.match(code):
            for target, target_line in targets:
                if target.lower() not in labels:
                    findings.append(Finding(
                        "ERROR", mod.name, target_line,
                        "Sprungziel '%s' ist in dieser Prozedur nicht definiert" % target))
            current_proc = None
            continue
        if current_proc is None:
            continue
        lm = LABEL_RE.match(code)
        if lm:
            labels.add(lm.group(1).lower())
            continue
        for gm in GOTO_RE.finditer(code):
            target = gm.group(1)
            if target == "0" or target.lower() == "next":
                continue
            targets.append((target, line_no))


def check_calls(modules, procs, module_procs, declared, findings):
    private_lookup = {}
    for name, entries in procs.items():
        for proc in entries:
            private_lookup.setdefault(name, []).append(proc)

    for mod in modules:
        for line_no, code, _ in mod.logical:
            stripped = code.strip()
            if not stripped or stripped.startswith("Attribute "):
                continue
            if PROC_RE.match(stripped):
                continue
            for m in IDENT_RE.finditer(stripped):
                ident = m.group(0)
                lower = ident.lower()
                if not lower.startswith("pid_") and lower not in procs:
                    continue
                if lower in VBA_KEYWORDS:
                    continue
                # member access (obj.PID_x) is not a project-level call
                if m.start() > 0 and stripped[m.start() - 1] in ".!":
                    continue
                if lower not in declared:
                    findings.append(Finding(
                        "ERROR", mod.name, line_no,
                        "'%s' ist nirgends deklariert (Sub or Function not defined)" % ident))
                    continue
                entries = private_lookup.get(lower)
                if not entries:
                    continue
                if any(p.module == mod.name for p in entries):
                    continue
                visible = [p for p in entries
                           if p.visibility.lower() != "private" and not
                           next(x for x in modules if x.name == p.module).is_class]
                if not visible:
                    owner = entries[0]
                    kind = "Klassenmodul" if next(
                        x for x in modules if x.name == owner.module).is_class else "Private"
                    findings.append(Finding(
                        "ERROR", mod.name, line_no,
                        "'%s' ist nur in %s sichtbar (%s) - Aufruf aus %s schlaegt fehl"
                        % (ident, owner.module, kind, mod.name)))


NAMED_ARG_RE = re.compile(r"([A-Za-z_]\w*)\s*:=")
AS_TYPE_RE = re.compile(r"\bAs\s+(?:New\s+)?([A-Za-z_][\w.]*)", re.IGNORECASE)


def collect_module_level_names(mod: Module):
    """Names declared outside any procedure (constants, module variables)."""
    names = set()
    inside_proc = False
    for _, code, _ in mod.logical:
        if PROC_RE.match(code):
            inside_proc = True
            continue
        if END_PROC_RE.match(code):
            inside_proc = False
            continue
        if inside_proc:
            continue
        cm = CONST_RE.match(code)
        if cm:
            names.update(n.lower() for n in split_declared_names(cm.group(2)))
            continue
        dm = DIM_RE.match(code)
        if dm:
            head = code.strip().split(None, 1)[0].lower()
            if head in {"dim", "private", "public", "global", "static"}:
                names.update(n.lower() for n in split_declared_names(dm.group(1)))
    return names


def check_undeclared(modules, procs, findings):
    """Option Explicit violations - the classic "Variable not defined" compile error."""
    known = set()
    for mod in modules:
        known |= collect_module_level_names(mod)
    for proc in procs.values() if isinstance(procs, dict) else []:
        pass
    for entries in procs.values():
        for proc in entries:
            known.add(proc.name.lower())
    known |= GLOBAL_IDENTIFIERS
    known |= VBA_KEYWORDS

    for mod in modules:
        local = set()
        labels = set()
        current_proc = None
        body = []
        for line_no, code, _ in mod.logical:
            pm = PROC_RE.match(code)
            if pm:
                current_proc = pm.group(3)
                local = {current_proc.lower()}
                labels = set()
                body = []
                params = code[code.find("(") + 1:code.rfind(")")] if "(" in code else ""
                local.update(n.lower() for n in split_declared_names(params))
                local.update(m.group(1).lower() for m in AS_TYPE_RE.finditer(code))
                continue
            if current_proc is None:
                continue
            if END_PROC_RE.match(code):
                for use_line, ident in body:
                    low = ident.lower()
                    if low in local or low in labels or low in known:
                        continue
                    if low.startswith(GLOBAL_PREFIXES):
                        continue
                    findings.append(Finding(
                        "ERROR", mod.name, use_line,
                        "'%s' ist in %s nicht deklariert (Option Explicit)"
                        % (ident, current_proc)))
                current_proc = None
                continue

            lm = LABEL_RE.match(code)
            if lm:
                labels.add(lm.group(1).lower())
                continue

            cm = CONST_RE.match(code)
            if cm:
                local.update(n.lower() for n in split_declared_names(cm.group(2)))
                continue
            dm = DIM_RE.match(code)
            if dm:
                head = code.strip().split(None, 1)[0].lower()
                if head in {"dim", "static", "redim"}:
                    local.update(n.lower() for n in split_declared_names(dm.group(1)))
                    continue

            local.update(m.group(1).lower() for m in AS_TYPE_RE.finditer(code))
            skip_positions = set()
            for m in NAMED_ARG_RE.finditer(code):
                skip_positions.add(m.start(1))
            for m in AS_TYPE_RE.finditer(code):
                skip_positions.add(m.start(1))
            for gm in GOTO_RE.finditer(code):
                skip_positions.add(gm.start(1))
            for m in IDENT_RE.finditer(code):
                if m.start() in skip_positions:
                    continue
                if m.start() > 0 and code[m.start() - 1] in ".!#[&":
                    continue
                body.append((line_no, m.group(0)))


def check_duplicates(modules, procs, findings):
    class_names = {m.name for m in modules if m.is_class}
    for lower, entries in sorted(procs.items()):
        public_entries = [p for p in entries
                          if p.visibility.lower() != "private" and p.module not in class_names]
        if len(public_entries) > 1:
            where = ", ".join("%s:%d" % (p.module, p.line) for p in public_entries)
            findings.append(Finding(
                "ERROR", public_entries[0].module, public_entries[0].line,
                "Ambiguous name: '%s' ist mehrfach oeffentlich definiert (%s)"
                % (public_entries[0].name, where)))
        by_module = {}
        for proc in entries:
            by_module.setdefault(proc.module, []).append(proc)
        for module_name, module_entries in by_module.items():
            if len(module_entries) > 1:
                findings.append(Finding(
                    "ERROR", module_name, module_entries[0].line,
                    "'%s' ist in %s mehrfach definiert (Zeilen %s)"
                    % (module_entries[0].name, module_name,
                       ", ".join(str(p.line) for p in module_entries))))


def check_headers(mod: Module, findings):
    declared_name = None
    has_option_explicit = False
    for raw in mod.lines:
        am = ATTRIBUTE_NAME_RE.match(raw)
        if am:
            declared_name = am.group(1)
        if re.match(r"^\s*Option\s+Explicit\s*$", raw, re.IGNORECASE):
            has_option_explicit = True
    if declared_name is None:
        findings.append(Finding("ERROR", mod.name, 1, "Attribute VB_Name fehlt"))
    elif declared_name != mod.name:
        findings.append(Finding(
            "ERROR", mod.name, 1,
            "Attribute VB_Name = '%s' passt nicht zum Dateinamen" % declared_name))
    if not has_option_explicit:
        findings.append(Finding("ERROR", mod.name, 1, "Option Explicit fehlt"))


def check_excel2016(mod: Module, findings):
    """Only literals that really end up in a cell are checked.

    mod_SmokeCheck legitimately mentions XLOOKUP/XVERWEIS in its own detection
    routine, so plain text mentions must not be reported.
    """
    current_proc = ""
    for line_no, code, literals in mod.logical:
        pm = PROC_RE.match(code)
        if pm:
            current_proc = pm.group(3)
        joined = " ".join(literals).upper()
        if not joined:
            continue
        is_formula_context = ("formula" in current_proc.lower()
                              or ".formula" in code.lower()
                              or joined.lstrip().startswith("="))
        if not is_formula_context:
            continue
        for token in FORBIDDEN_FORMULA_TOKENS:
            if token in joined:
                findings.append(Finding(
                    "ERROR", mod.name, line_no,
                    "Formel nicht Excel-2016-kompatibel: %s" % token))


# Bootstrap modules run exactly when the rest of the project is missing or broken.
# An early bound call into another module makes the whole project fail to compile -
# and then the repair macro itself can no longer be started. See .cursor/rules.md.
BOOTSTRAP_MODULES = ("mod_ResetAndImportVBAFiles",)


def check_bootstrap_isolation(modules, procs, findings):
    own = {}
    for mod in modules:
        own[mod.name] = set()
    for entries in procs.values():
        for proc in entries:
            own.setdefault(proc.module, set()).add(proc.name.lower())

    for mod in modules:
        if mod.name not in BOOTSTRAP_MODULES:
            continue
        for line_no, code, _ in mod.logical:
            if PROC_RE.match(code):
                continue
            for m in IDENT_RE.finditer(code):
                lower = m.group(0).lower()
                if lower in own[mod.name] or lower not in procs:
                    continue
                if m.start() > 0 and code[m.start() - 1] in ".!":
                    continue
                owners = sorted({p.module for p in procs[lower] if p.module != mod.name})
                if not owners:
                    continue
                findings.append(Finding(
                    "ERROR", mod.name, line_no,
                    "Bootstrap-Modul ruft '%s' aus %s auf - fehlt das Modul, laesst sich "
                    "auch dieses Makro nicht mehr starten (Application.Run oder eigenen "
                    "Text verwenden)" % (m.group(0), ", ".join(owners))))


CONST_CALL_RE = re.compile(r"\b([A-Za-z_]\w*)\s*\(")


def check_const_expressions(mod: Module, findings):
    """Const values must be compile time constants - Const X = "a" & ChrW(10) is
    "Konstanter Ausdruck erforderlich" in the VBA compiler."""
    for line_no, code, _ in mod.logical:
        cm = CONST_RE.match(code)
        if not cm:
            continue
        value = cm.group(2)
        if "=" not in value:
            continue
        expression = value.split("=", 1)[1]
        for m in CONST_CALL_RE.finditer(expression):
            name = m.group(1)
            if name.lower() in ("array",):
                continue
            findings.append(Finding(
                "ERROR", mod.name, line_no,
                "Const darf keinen Funktionsaufruf enthalten ('%s') - "
                "konstanter Ausdruck erforderlich" % name))


def check_object_model(mod: Module, findings):
    for line_no, code, _ in mod.logical:
        for m in APPLICATION_MEMBER_RE.finditer(code):
            if m.group(1).lower() in POST_2016_APPLICATION_MEMBERS:
                findings.append(Finding(
                    "ERROR", mod.name, line_no,
                    "Application.%s gibt es in Excel 2016 nicht - frueh gebunden ist das "
                    "ein Compile-Fehler; nur ueber ein Object aufrufen" % m.group(1)))
        for m in WORKSHEETFUNCTION_MEMBER_RE.finditer(code):
            if m.group(1).lower() in POST_2016_WORKSHEETFUNCTION_MEMBERS:
                findings.append(Finding(
                    "ERROR", mod.name, line_no,
                    "WorksheetFunction.%s gibt es in Excel 2016 nicht" % m.group(1)))


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--vba-dir", default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "vba"))
    parser.add_argument("--quiet", action="store_true",
                        help="only print findings, no summary table")
    args = parser.parse_args()

    modules = load_modules(args.vba_dir)
    if not modules:
        print("No VBA files found in %s" % args.vba_dir)
        return 1

    procs, module_procs, declared = collect_symbols(modules)
    findings = []
    for mod in modules:
        check_headers(mod, findings)
        check_blocks(mod, findings)
        check_labels(mod, findings)
        check_excel2016(mod, findings)
        check_object_model(mod, findings)
        check_const_expressions(mod, findings)
    check_calls(modules, procs, module_procs, declared, findings)
    check_duplicates(modules, procs, findings)
    check_undeclared(modules, procs, findings)
    check_bootstrap_isolation(modules, procs, findings)

    errors = [f for f in findings if f.level == "ERROR"]
    warnings = [f for f in findings if f.level == "WARN"]

    for finding in sorted(findings, key=lambda f: (f.module, f.line)):
        print("%-5s %s:%d  %s" % (finding.level, finding.module, finding.line,
                                  finding.message))

    if not args.quiet:
        total_lines = sum(len(m.lines) for m in modules)
        print("")
        print("Modules: %d   procedures: %d   lines: %d"
              % (len(modules), sum(len(v) for v in procs.values()), total_lines))
        print("Errors: %d   Warnings: %d" % (len(errors), len(warnings)))

    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
