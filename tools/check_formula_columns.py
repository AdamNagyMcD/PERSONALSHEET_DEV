#!/usr/bin/env python3
"""Report the health of the formula columns G, H, K and L in Personalsheet.xlsm.

Linux twin of the VBA diagnostic PID_PruefeFormelspalten (TR-10). Older versions
cleared B:N when an employee was deleted, which took the four formula columns with
it; such rows stay dead - a new employee entered there gets no Monatslohn, no
Aktuelle Stunden, no Urlaub in Euro and no Letztes Gehalt.

The script only reads the workbook, it never writes. Use it before and after
running "Formeln reparieren" / Full Refresh in Excel to prove the repair worked.

Usage:
    python3 tools/check_formula_columns.py [--workbook Personalsheet.xlsm] [--verbose]

Exit code 0 = every row has its formulas, 1 = gaps found.
"""

from __future__ import annotations

import argparse
import os
import sys
import warnings

MONTHS = ["Januar", "Februar", "Marz", "April", "Mai", "Juni",
          "Juli", "August", "September", "Oktober", "November", "Dezember"]
FORMULA_COLUMNS = ("G", "H", "K", "L")
FIRST_ROW = 3
LAST_ROW = 82


def has_formula(cell) -> bool:
    return isinstance(cell.value, str) and cell.value.startswith("=")


def compress(rows):
    """[3,4,5,9] -> '3-5, 9'"""
    if not rows:
        return ""
    parts = []
    start = previous = rows[0]
    for row in rows[1:]:
        if row == previous + 1:
            previous = row
            continue
        parts.append(str(start) if start == previous else "%d-%d" % (start, previous))
        start = previous = row
    parts.append(str(start) if start == previous else "%d-%d" % (start, previous))
    return ", ".join(parts)


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--workbook", default=os.path.join(root, "Personalsheet.xlsm"))
    parser.add_argument("--verbose", action="store_true",
                        help="also list the affected rows per column")
    args = parser.parse_args()

    warnings.simplefilter("ignore")
    import openpyxl

    workbook = openpyxl.load_workbook(args.workbook, data_only=False)

    total_missing = 0
    bad_index_sheets = []

    print("%-11s %5s | %5s %5s %5s %5s   (Zellen ohne Formel, Zeile %d-%d)"
          % ("Blatt", "A1", *FORMULA_COLUMNS, FIRST_ROW, LAST_ROW))
    for month_index, name in enumerate(MONTHS, start=1):
        if name not in workbook.sheetnames:
            print("%-11s FEHLT" % name)
            total_missing += 1
            continue
        sheet = workbook[name]
        index_cell = sheet["A1"].value
        if index_cell != month_index:
            bad_index_sheets.append("%s (A1=%r)" % (name, index_cell))

        counts = {}
        detail = {}
        for column in FORMULA_COLUMNS:
            rows = [row for row in range(FIRST_ROW, LAST_ROW + 1)
                    if not has_formula(sheet["%s%d" % (column, row)])]
            counts[column] = len(rows)
            detail[column] = rows
            total_missing += len(rows)

        marker = "" if index_cell == month_index else "  <- Monatsindex falsch"
        print("%-11s %5s | %5d %5d %5d %5d%s"
              % (name, index_cell, counts["G"], counts["H"], counts["K"], counts["L"], marker))
        if args.verbose:
            for column in FORMULA_COLUMNS:
                if detail[column]:
                    print("            %s: Zeilen %s" % (column, compress(detail[column])))

    print("")
    if bad_index_sheets:
        print("Monatsindex A1 falsch: %s" % ", ".join(bad_index_sheets))
    print("Fehlende Formelzellen gesamt: %d" % total_missing)
    if total_missing:
        print("Reparatur in Excel: _ADMIN > 'Formeln reparieren' oder Full Refresh, "
              "danach speichern und dieses Skript erneut laufen lassen.")
    return 1 if total_missing or bad_index_sheets else 0


if __name__ == "__main__":
    sys.exit(main())
