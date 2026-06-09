from __future__ import annotations

from .gmail_monitor import get_gmail_service
from .settings import PROJECT_ROOT


def main() -> None:
    get_gmail_service(PROJECT_ROOT)
    print("Gmail authorization complete.")


if __name__ == "__main__":
    main()
