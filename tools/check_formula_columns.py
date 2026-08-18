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
# H, K and L are pure formula columns. G is different: RefreshKVLohnForRow writes the
# wage as a plain value, which PID_MonthSheetHasMonatslohnFormula accepts as valid
# (see TEST 15). For G a cell only counts as broken when the row holds an employee
# with KV group and hours but G stays empty.
FORMULA_COLUMNS = ("H", "K", "L")
FIRST_ROW = 3
LAST_ROW = 82
# One wrapping of the Letztes-Gehalt formula is around 1500 characters. Every further
# wrapping doubles it, and above 8192 Excel refuses the formula altogether.
LONG_FORMULA_WARNING = 2200


def has_formula(cell) -> bool:
    return isinstance(cell.value, str) and cell.value.startswith("=")


def is_filled(cell) -> bool:
    return cell.value not in (None, "")


def broken_wage_rows(sheet):
    """Rows with an employee and complete KV input but no wage in G."""
    rows = []
    for row in range(FIRST_ROW, LAST_ROW + 1):
        has_employee = is_filled(sheet["B%d" % row]) or is_filled(sheet["C%d" % row])
        if not has_employee:
            continue
        if not (is_filled(sheet["E%d" % row]) and is_filled(sheet["F%d" % row])):
            continue
        cell = sheet["G%d" % row]
        if not has_formula(cell) and not is_filled(cell):
            rows.append(row)
    return rows


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
    wage_gaps = []
    long_formulas = []

    print("%-11s %5s | %5s %5s %5s | %6s | %8s   (Zellen ohne Formel, Zeile %d-%d)"
          % ("Blatt", "A1", *FORMULA_COLUMNS, "G ohne", "L Laenge", FIRST_ROW, LAST_ROW))
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

        wage_rows = broken_wage_rows(sheet)
        if wage_rows:
            wage_gaps.append((name, wage_rows))

        formula_l = sheet["L%d" % FIRST_ROW].value
        length_l = len(formula_l) if isinstance(formula_l, str) else 0
        if length_l > LONG_FORMULA_WARNING:
            long_formulas.append((name, length_l))

        marker = "" if index_cell == month_index else "  <- Monatsindex falsch"
        if length_l > LONG_FORMULA_WARNING:
            marker += "  <- L doppelt verpackt?"
        print("%-11s %5s | %5d %5d %5d | %6d | %8d%s"
              % (name, index_cell, counts["H"], counts["K"], counts["L"],
                 len(wage_rows), length_l, marker))
        if args.verbose:
            for column in FORMULA_COLUMNS:
                if detail[column]:
                    print("            %s: Zeilen %s" % (column, compress(detail[column])))
            if wage_rows:
                print("            G ohne Lohn: Zeilen %s" % compress(wage_rows))

    print("")
    if bad_index_sheets:
        print("Monatsindex A1 falsch: %s" % ", ".join(bad_index_sheets))
    if wage_gaps:
        print("Mitarbeiterzeilen mit KV-Gruppe und Stunden, aber ohne Lohn in G:")
        for name, rows in wage_gaps:
            print("   %-11s Zeilen %s" % (name, compress(rows)))
    if long_formulas:
        print("Ueberlange Letztes-Gehalt-Formeln (eine Verpackung ist ~1500 Zeichen,")
        print("jede weitere verdoppelt sie; ab 8192 lehnt Excel sie ab):")
        for name, length in long_formulas:
            print("   %-11s %d Zeichen" % (name, length))
    print("Fehlende Formelzellen in H/K/L: %d" % total_missing)
    if total_missing or wage_gaps or long_formulas:
        print("Reparatur in Excel: _ADMIN > 'Formeln reparieren' oder Full Refresh, "
              "danach speichern und dieses Skript erneut laufen lassen.")
    return 1 if total_missing or bad_index_sheets or wage_gaps or long_formulas else 0


if __name__ == "__main__":
    sys.exit(main())
