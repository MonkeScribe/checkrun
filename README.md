# checkrun

Turn a spreadsheet of checks into a QuickBooks Desktop **IIF** import file.

1. **Fill template** – raw data is copied into the *Check Data* tab of `templates/Check_to_IIF_Template.xlsx` and saved to your output folder.
2. **Convert to IIF** – the filled template is read back (so you can review/edit it first) and written as a `.iif` file (Windows-1252, CRLF).

## Run

```
pip install -r requirements.txt
python -m checkrun                      # GUI: drop a file, set output folder
python -m checkrun raw.xlsx out_dir     # headless: both steps
```

`tkinterdnd2` enables drag-and-drop; without it the drop box is click-to-browse.
The chosen output folder is remembered in `~/.checkrun/settings.json`.

## Raw file columns

Header names are matched loosely (case-insensitive). Required: **Date**, **Payee** (or Vendor/Name), **Amount**, **Check Number** (or Check #/Num).
Optional: Memo, Class, Bank Account, Expense Account (default to the template's row 2: `Checking` / `Dealer Checks`).

Validation reports every bad row at once: bad dates, non-positive amounts, duplicate or missing check numbers, characters QuickBooks can't take.

## Notes

- The template's own *IIF Output* tab only covers 60 checks; this tool has no such limit.
- Account/payee names must match QuickBooks exactly. Back up the company file before importing; IIF import can't be undone.
- Tests: `pytest`.
