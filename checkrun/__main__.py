import sys
from pathlib import Path

if __package__ in (None, ""):
    # Run as a plain script (python checkrun/__main__.py or an IDE "Run" button):
    # put the project root on sys.path so the package imports below still work.
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    __package__ = "checkrun"

from checkrun import core  # noqa: E402


def main() -> None:
    # `python -m checkrun` opens the GUI; `python -m checkrun raw.xlsx out_dir` runs both steps headless.
    if len(sys.argv) == 3:
        raw, out = Path(sys.argv[1]), Path(sys.argv[2])
        out.mkdir(parents=True, exist_ok=True)
        filled = out / (raw.stem + "_template.xlsx")
        iif = out / (raw.stem + ".iif")
        print(f"Filled {core.fill_template(raw, filled)} check(s) -> {filled}")
        print(f"Wrote {core.convert_template_to_iif(filled, iif)} check(s) -> {iif}")
        return
    from checkrun.gui import main as gui_main

    gui_main()


if __name__ == "__main__":
    main()
