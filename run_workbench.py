"""ASCII-named entry point for the Windows CMD launcher."""

from pathlib import Path
import runpy


def main():
    application = Path(__file__).with_name("procurement_workbench.py")
    if not application.is_file():
        raise FileNotFoundError(f"Workbench program not found: {application}")
    runpy.run_path(str(application), run_name="__main__")


if __name__ == "__main__":
    main()
