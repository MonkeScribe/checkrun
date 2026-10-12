import sys

from . import core


def main() -> None:
    # `python -m checkrun` opens the GUI; `python -m checkrun raw.xlsx out_dir [start_check_no]` runs both steps headless.
    if len(sys.argv) in (3, 4):
        from pathlib import Path

        raw, out = Path(sys.argv[1]), Path(sys.argv[2])
        start = int(sys.argv[3]) if len(sys.argv) == 4 else None
        out.mkdir(parents=True, exist_ok=True)
        filled = out / (raw.stem + "_template.xlsx")
        iif = out / (raw.stem + ".iif")
        res = core.fill_template(raw, filled, start_check_no=start)
        print(f"Filled {res.count} check(s), ${res.total:,.2f} -> {filled}")
        for n in res.notes:
            print("  " + n)
        print(f"Wrote {core.convert_template_to_iif(filled, iif)} check(s) -> {iif}")
        return
    from .gui import main as gui_main

    gui_main()


main()
