import sys

from . import core


def main() -> None:
    # `python -m checkrun` opens the GUI; `python -m checkrun raw.xlsx out_dir` runs both steps headless.
    if len(sys.argv) == 3:
        from pathlib import Path

        raw, out = Path(sys.argv[1]), Path(sys.argv[2])
        out.mkdir(parents=True, exist_ok=True)
        filled = out / (raw.stem + "_template.xlsx")
        iif = out / (raw.stem + ".iif")
        print(f"Filled {core.fill_template(raw, filled)} check(s) -> {filled}")
        print(f"Wrote {core.convert_template_to_iif(filled, iif)} check(s) -> {iif}")
        return
    from .gui import main as gui_main

    gui_main()


main()
