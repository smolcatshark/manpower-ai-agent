from __future__ import annotations

import logging
import os
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Final

APP_TITLE: Final = "Manpower Report Agent"
APP_FILENAME: Final = "streamlit_app.py"
CHILD_FLAG: Final = "--streamlit-child"
HOST: Final = "127.0.0.1"
STARTUP_TIMEOUT_SECONDS: Final = 120
HEALTH_PATH: Final = "/_stcore/health"


def is_frozen() -> bool:
    """Return True when running inside a PyInstaller bundle."""
    return bool(getattr(sys, "frozen", False))


def bundle_root() -> Path:
    """Return the directory containing bundled application resources."""
    if is_frozen():
        temporary_root = getattr(sys, "_MEIPASS", None)
        if temporary_root:
            return Path(temporary_root)
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent


def application_script() -> Path:
    """Return the bundled Streamlit application path."""
    return bundle_root() / APP_FILENAME


def log_directory() -> Path:
    """Return a writable per-user application log directory."""
    if os.name == "nt":
        root = Path(
            os.environ.get(
                "LOCALAPPDATA",
                Path.home() / "AppData" / "Local",
            )
        )
    else:
        root = Path.home() / ".local" / "state"

    directory = root / APP_TITLE
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def configure_logging() -> Path:
    """Configure launcher logging and return the log file path."""
    log_path = log_directory() / "launcher.log"

    logging.basicConfig(
        level=logging.INFO,
        format=(
            "%(asctime)s | %(levelname)s | "
            "%(process)d | %(message)s"
        ),
        handlers=[
            logging.FileHandler(
                log_path,
                encoding="utf-8",
            )
        ],
        force=True,
    )
    return log_path


def show_error(title: str, message: str) -> None:
    """Display an error dialog without requiring a console window."""
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showerror(title, message, parent=root)
        root.destroy()
    except Exception:
        logging.exception("Unable to display the error dialog.")


def find_available_port() -> int:
    """Ask Windows for a free localhost TCP port."""
    with socket.socket(
        socket.AF_INET,
        socket.SOCK_STREAM,
    ) as probe:
        probe.bind((HOST, 0))
        return int(probe.getsockname()[1])


def child_port_from_arguments() -> int:
    """Read the Streamlit child-process port from command arguments."""
    try:
        flag_index = sys.argv.index(CHILD_FLAG)
        port_text = sys.argv[flag_index + 1]
        port = int(port_text)
    except (
        ValueError,
        IndexError,
        TypeError,
    ) as exc:
        raise ValueError(
            "The Streamlit child process did not receive a valid port."
        ) from exc

    if not 1 <= port <= 65535:
        raise ValueError(
            f"Invalid Streamlit port: {port}"
        )

    return port


def run_streamlit_child() -> None:
    """Run Streamlit inside the bundled child process."""
    port = child_port_from_arguments()
    app_path = application_script()

    if not app_path.exists():
        raise FileNotFoundError(
            f"Bundled application file not found: {app_path}"
        )

    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = (
        "false"
    )
    os.environ.setdefault(
        "STREAMLIT_BROWSER_GATHER_USAGE_STATS",
        "false",
    )
    os.environ.setdefault(
        "STREAMLIT_SERVER_HEADLESS",
        "true",
    )
    os.environ.setdefault(
        "STREAMLIT_SERVER_FILE_WATCHER_TYPE",
        "none",
    )

    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        f"--server.address={HOST}",
        f"--server.port={port}",
        "--global.developmentMode=false",
        "--server.headless=true",
        "--server.fileWatcherType=none",
        "--server.runOnSave=false",
        "--browser.gatherUsageStats=false",
    ]

    from streamlit.web.cli import main as streamlit_main

    raise SystemExit(streamlit_main())


def server_command(port: int) -> list[str]:
    """Build the command used to start the local Streamlit child."""
    if is_frozen():
        return [
            sys.executable,
            CHILD_FLAG,
            str(port),
        ]

    return [
        sys.executable,
        str(Path(__file__).resolve()),
        CHILD_FLAG,
        str(port),
    ]


def start_streamlit_server(
    port: int,
    log_path: Path,
) -> tuple[subprocess.Popen[bytes], object]:
    """Start Streamlit invisibly and redirect output to the log file."""
    creation_flags = 0
    if os.name == "nt":
        creation_flags = getattr(
            subprocess,
            "CREATE_NO_WINDOW",
            0,
        )

    log_handle = log_path.open("ab", buffering=0)

    process = subprocess.Popen(
        server_command(port),
        cwd=str(bundle_root()),
        env=os.environ.copy(),
        stdin=subprocess.DEVNULL,
        stdout=log_handle,
        stderr=subprocess.STDOUT,
        creationflags=creation_flags,
    )

    logging.info(
        "Started Streamlit child process %s on port %s.",
        process.pid,
        port,
    )
    return process, log_handle


def wait_for_streamlit(
    process: subprocess.Popen[bytes],
    port: int,
) -> str:
    """Wait until Streamlit reports a healthy local server."""
    base_url = f"http://{HOST}:{port}"
    health_url = base_url + HEALTH_PATH
    deadline = time.monotonic() + STARTUP_TIMEOUT_SECONDS
    last_error = "No response received."

    while time.monotonic() < deadline:
        exit_code = process.poll()
        if exit_code is not None:
            raise RuntimeError(
                "Streamlit closed before the app opened "
                f"(exit code {exit_code})."
            )

        try:
            with urllib.request.urlopen(
                health_url,
                timeout=1.5,
            ) as response:
                body = response.read().decode(
                    "utf-8",
                    errors="replace",
                )
                if (
                    response.status == 200
                    and body.strip().lower() == "ok"
                ):
                    logging.info(
                        "Streamlit health check passed."
                    )
                    return base_url
        except (
            urllib.error.URLError,
            TimeoutError,
            ConnectionError,
            OSError,
        ) as exc:
            last_error = str(exc)

        time.sleep(0.35)

    raise TimeoutError(
        "The app did not finish starting within "
        f"{STARTUP_TIMEOUT_SECONDS} seconds. "
        f"Last connection result: {last_error}"
    )


def stop_process(
    process: subprocess.Popen[bytes] | None,
) -> None:
    """Stop the local Streamlit process safely."""
    if process is None or process.poll() is not None:
        return

    logging.info(
        "Stopping Streamlit child process %s.",
        process.pid,
    )
    process.terminate()

    try:
        process.wait(timeout=8)
    except subprocess.TimeoutExpired:
        logging.warning(
            "Streamlit did not close normally; killing it."
        )
        process.kill()
        process.wait(timeout=5)


def open_desktop_window(url: str) -> None:
    """Open the Streamlit interface inside a native desktop window."""
    import webview

    webview.settings["ALLOW_DOWNLOADS"] = True
    webview.settings[
        "OPEN_EXTERNAL_LINKS_IN_BROWSER"
    ] = True

    webview.create_window(
        APP_TITLE,
        url,
        width=1440,
        height=920,
        min_size=(1024, 700),
        resizable=True,
        confirm_close=False,
        background_color="#FFFFFF",
        text_select=True,
        zoomable=True,
    )

    webview.start(
        debug=False,
        private_mode=True,
    )


def run_desktop_application() -> int:
    """Start the server, show the window, and clean up on exit."""
    log_path = configure_logging()
    process: subprocess.Popen[bytes] | None = None
    log_handle = None

    try:
        app_path = application_script()
        if not app_path.exists():
            raise FileNotFoundError(
                f"Application file not found: {app_path}"
            )

        port = find_available_port()
        process, log_handle = start_streamlit_server(
            port,
            log_path,
        )
        app_url = wait_for_streamlit(
            process,
            port,
        )
        open_desktop_window(app_url)
        return 0

    except Exception as exc:
        logging.exception(
            "Desktop application startup failed."
        )
        show_error(
            APP_TITLE,
            (
                "The application could not start.\n\n"
                f"{exc}\n\n"
                "Technical details were written to:\n"
                f"{log_path}"
            ),
        )
        return 1

    finally:
        stop_process(process)
        if log_handle is not None:
            log_handle.close()


def main() -> int:
    """Dispatch either the hidden Streamlit child or desktop parent."""
    if CHILD_FLAG in sys.argv:
        configure_logging()
        try:
            run_streamlit_child()
        except SystemExit:
            raise
        except Exception:
            logging.exception(
                "Streamlit child process failed."
            )
            return 1

    return run_desktop_application()


if __name__ == "__main__":
    raise SystemExit(main())
