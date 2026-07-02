"""Run the local estimate web UI: python -m app.web [--host H] [--port P]."""

import argparse
import threading
import time
import webbrowser

import uvicorn

from app.web.app import create_app


def main() -> None:
    parser = argparse.ArgumentParser(description="Estimate AI local web UI")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--base-dir", default=None, help="working directory for uploads")
    parser.add_argument("--no-browser", action="store_true", help="do not open a browser")
    args = parser.parse_args()

    app = create_app(args.base_dir)
    url = f"http://{args.host}:{args.port}/"
    print(f"Estimate AI web UI running at {url}")
    print("Press Ctrl+C to stop.")

    if not args.no_browser:
        _open_browser_when_ready(url)

    uvicorn.run(app, host=args.host, port=args.port, log_level="info")


def _open_browser_when_ready(url: str) -> None:
    def opener() -> None:
        time.sleep(1.0)
        try:
            webbrowser.open(url)
        except Exception:
            pass

    threading.Thread(target=opener, daemon=True).start()


if __name__ == "__main__":
    main()
