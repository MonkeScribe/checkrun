from datetime import date
from decimal import Decimal

import openpyxl
import pandas as pd
import pytest

from checkrun import core


def raw_file(tmp_path, rows, name="raw.xlsx"):
    p = tmp_path / name
    pd.DataFrame(rows).to_excel(p, index=False)
    return p


GOOD = [
    {"Vendor": "Acme Motors", "Check Date": "2026-08-15", "Total": "$1,234.50", "Check #": 1001, "Description": "Aug"},
    {"Vendor": "Bob's Parts", "Check Date": "2026-08-15", "Total": 99, "Check #": 1002},
]


def test_end_to_end(tmp_path):
    filled, iif = tmp_path / "f.xlsx", tmp_path / "o.iif"
    assert core.fill_template(raw_file(tmp_path, GOOD), filled).count == 2
    ws = openpyxl.load_workbook(filled)["Check Data"]
    assert ws["B3"].value == "Acme Motors" and ws["C3"].value == "Checking" and ws["D4"].value == "Dealer Checks"
    assert core.convert_template_to_iif(filled, iif) == 2

    lines = iif.read_bytes().decode("cp1252").split("\r\n")
    assert lines[0].startswith("!TRNS\tTRNSID") and lines[2] == "!ENDTRNS"
    assert lines[3] == "TRNS\t\tCHECK\t08/15/2026\tChecking\tAcme Motors\t\t-1234.50\t1001\tAug\tY"
    assert lines[4] == "SPL\t\tCHECK\t08/15/2026\tDealer Checks\t\t1234.50\t1001\tAug"
    assert lines[5] == "ENDTRNS"
    assert lines[6].endswith("-99.00\t1002\t\tY")


def test_blocks_balance():
    c = core.Check(date(2026, 1, 2), "X", "Checking", "Dealer Checks", Decimal("10.10"), "1", "", "")
    t = [l.split("\t") for l in core.build_iif([c]).splitlines() if l.startswith(("TRNS", "SPL"))]
    assert Decimal(t[0][7]) + Decimal(t[1][6]) == 0


def test_missing_column(tmp_path):
    with pytest.raises(core.CheckRunError, match="missing column"):
        core.fill_template(raw_file(tmp_path, [{"Vendor": "A", "Total": 1}]), tmp_path / "f.xlsx")


def test_bad_rows_reported_together(tmp_path):
    rows = GOOD + [
        {"Vendor": "", "Check Date": "nope", "Total": -5, "Check #": 1001},
    ]
    with pytest.raises(core.CheckRunError) as e:
        core.fill_template(raw_file(tmp_path, rows), tmp_path / "f.xlsx")
    msg = str(e.value)
    assert "Row 4: invalid or missing Date" in msg and "missing Payee" in msg
    assert "Amount must be a positive" in msg and "duplicate Check Number 1001" in msg


def test_tabs_and_newlines_stripped(tmp_path):
    rows = [{"Vendor": "A\tB\nC", "Check Date": "2026-01-01", "Total": 5, "Check #": 7}]
    filled, iif = tmp_path / "f.xlsx", tmp_path / "o.iif"
    core.fill_template(raw_file(tmp_path, rows), filled)
    core.convert_template_to_iif(filled, iif)
    assert "A B C" in iif.read_text("cp1252")


def test_gap_in_filled_template_is_caught(tmp_path):
    filled = tmp_path / "f.xlsx"
    core.fill_template(raw_file(tmp_path, GOOD), filled)
    wb = openpyxl.load_workbook(filled)
    wb["Check Data"]["B3"].value = None
    wb.save(filled)
    with pytest.raises(core.CheckRunError, match="no Payee"):
        core.read_template_checks(filled)


def test_unencodable_character(tmp_path):
    rows = [{"Vendor": "日本", "Check Date": "2026-01-01", "Total": 5, "Check #": 7}]
    filled = tmp_path / "f.xlsx"
    core.fill_template(raw_file(tmp_path, rows), filled)
    with pytest.raises(core.CheckRunError, match="Windows-1252"):
        core.convert_template_to_iif(filled, tmp_path / "o.iif")


def test_csv_input(tmp_path):
    p = tmp_path / "r.csv"
    p.write_text("Payee,Date,Amount,Check Number\nZed,8/1/2026,12.5,55\n")
    assert core.fill_template(p, tmp_path / "f.xlsx").count == 1


def test_summarize_flags_far_dates():
    cs = [core.Check(date(2026, 8, 15), "A", "B", "E", Decimal("10"), "1", "", ""),
          core.Check(date(2025, 1, 1), "B", "B", "E", Decimal("5.5"), "2", "", "")]
    out = core.summarize(cs, today=date(2026, 8, 16))
    assert "Checks: 2" in out and "01/01/2025 – 08/15/2026" in out and "$15.50" in out
    assert "1 check(s) are dated more than 30 days" in out and "#2" in out


def booth_report(tmp_path):
    p = tmp_path / "sales_2026-10-01_2026-10-15.xlsx"
    booth = pd.DataFrame({
        "Vendor": ["Ann A", "Concessions", "Zero Z", "Bob B"],
        "Booth": ["001", "700", "002", 3],
        "Payout": [100.1, 130.5, 0, 55.55],
    })
    with pd.ExcelWriter(p) as w:
        pd.DataFrame({"x": [1]}).to_excel(w, sheet_name="Store Totals", index=False)
        booth.to_excel(w, sheet_name="Booth Totals", index=False)
    return p


def test_booth_totals_ignores_concessions_and_zero(tmp_path):
    p = booth_report(tmp_path)
    res = core.fill_template(p, tmp_path / "f.xlsx", check_date=date(2026, 10, 15), start_check_no=7001)
    assert res.count == 2 and res.total == Decimal("155.65") and res.next_check_no == 7003
    assert any("Concessions" in n for n in res.notes) and any("Zero Z" in n for n in res.notes)
    cs = core.read_template_checks(tmp_path / "f.xlsx")
    assert [(c.payee, c.check_no, c.amount) for c in cs] == [("Ann A", "7001", Decimal("100.10")), ("Bob B", "7002", Decimal("55.55"))]
    assert cs[0].date == date(2026, 10, 15) and cs[0].memo == "Booth 001 10/01-10/15/2026" and cs[1].memo.startswith("Booth 3 ")


def test_booth_totals_needs_start_number(tmp_path):
    with pytest.raises(core.CheckRunError, match="starting check number"):
        core.fill_template(booth_report(tmp_path), tmp_path / "f.xlsx")


def test_booth_sort_orders(tmp_path):
    p = tmp_path / "r.xlsx"
    booth = pd.DataFrame({
        "Vendor": ["Zed", "Amy", "Max", "Joe", "Kim"],
        "Booth": ["BAM", "100", "002", "12D", "009"],
        "Payout": [10, 10, 10, 10, 10],
    })
    with pd.ExcelWriter(p) as w:
        booth.to_excel(w, sheet_name="Booth Totals", index=False)

    def order(sort_by):
        core.fill_template(p, tmp_path / "f.xlsx", start_check_no=1, sort_by=sort_by)
        return [(c.payee, c.check_no) for c in core.read_template_checks(tmp_path / "f.xlsx")]

    assert order("booth") == [("Max", "1"), ("Kim", "2"), ("Amy", "3"), ("Joe", "4"), ("Zed", "5")]  # 002 first
    assert order("vendor") == [("Amy", "1"), ("Joe", "2"), ("Kim", "3"), ("Max", "4"), ("Zed", "5")]


def test_unknown_sort(tmp_path):
    with pytest.raises(core.CheckRunError, match="Unknown sort"):
        core.fill_template(booth_report(tmp_path), tmp_path / "f.xlsx", start_check_no=1, sort_by="nope")
