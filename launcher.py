"""
Altment converter - self-updating launcher.
===========================================

Build this file into ONE .exe with PyInstaller and hand that .exe to your
coworkers. From then on, updating the converter for everybody is just:

    git commit + git push

...and the next time anyone runs the .exe, they get the new version. No
rebuild, no Drive upload, no "please replace your old file".

HOW IT WORKS
------------
The .exe does not contain the converter. It contains the Python runtime and
the libraries the converter needs (python-docx, lxml, tkinter - see the
"BUNDLED DEPENDENCIES" imports below), plus this ~200 lines of bootstrap.
On every launch it:

  1. downloads convert_to_html2.py from your GitHub repo,
  2. checks the download actually compiles (so a bad push can't brick
     everyone - see _is_valid_python below),
  3. saves it next to the .exe as a cache, and
  4. executes it in-process, as if it had been run directly.

Because the interpreter and the libraries are already inside the .exe, the
downloaded file is plain source that needs nothing installed on the machine.

WHEN YOU STILL NEED TO REBUILD THE .EXE
---------------------------------------
Only when the converter starts importing a library that isn't bundled yet
(add it to the imports below first), or when you change this launcher.
Editing the converter itself never requires a rebuild.

OFFLINE / GITHUB DOWN
---------------------
The cached copy from the previous run is used instead, so the converter
keeps working. Only the very first run genuinely requires network access.
"""

import hashlib
import json
import os
import ssl
import sys
import time
import traceback
import urllib.error
import urllib.request

# =====================================================================
# CONFIGURATION - the only part you need to edit.
# =====================================================================
GITHUB_OWNER = "YOUR-GITHUB-USERNAME-OR-ORG"
GITHUB_REPO = "altment-converter"
GITHUB_BRANCH = "main"
PAYLOAD_PATH_IN_REPO = "convert_to_html2.py"

# No credentials anywhere in this file, by design.
#
# The repo is PUBLIC, so the converter is fetched anonymously. The Mailchimp
# API key it needs is never in the repo either - the converter reads it at
# runtime from a config.json placed next to this .exe during setup, so the
# code can be public and auto-updating while the credential stays local to
# each machine.

# How long to wait for GitHub before giving up and using the cached copy.
# Deliberately short: a slow network should delay the tool by seconds, not
# leave someone staring at a frozen window.
UPDATE_TIMEOUT_SECONDS = 12

# Bumped only when this launcher itself changes in a way the converter
# depends on. The converter can require a minimum by including a line
#     # requires-launcher: 2
# anywhere in its source; older launchers then tell the user to get a new
# .exe instead of failing in a confusing way further down.
LAUNCHER_VERSION = 1
# =====================================================================

# --- BUNDLED DEPENDENCIES -------------------------------------------
# These imports exist so PyInstaller sees them and bundles them into the
# .exe. The launcher itself doesn't use them - the downloaded converter
# does, and by then it's too late for PyInstaller to notice. Add any new
# third-party import the converter picks up here, then rebuild once.
import tkinter  # noqa: F401
import tkinter.filedialog  # noqa: F401
import tkinter.ttk  # noqa: F401
import tkinter.colorchooser  # noqa: F401
import tkinter.messagebox  # noqa: F401
import tkinter.scrolledtext  # noqa: F401
import docx  # noqa: F401  (python-docx)
import lxml.etree  # noqa: F401
import certifi  # noqa: F401  (CA bundle - see _ssl_context)
# ---------------------------------------------------------------------


def _base_dir():
    """The folder the .exe lives in - NOT PyInstaller's temporary unpack
    folder. Same reasoning as SCRIPT_DIR inside the converter: everything
    that has to survive between runs (the cached script, the signature
    photos, the Mailchimp upload cache) belongs next to the executable,
    because sys._MEIPASS is deleted on exit."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = _base_dir()
CACHE_DIR = os.path.join(BASE_DIR, "app")
CACHED_SCRIPT = os.path.join(CACHE_DIR, os.path.basename(PAYLOAD_PATH_IN_REPO))
STATE_PATH = os.path.join(CACHE_DIR, "update_state.json")


def _log(message):
    print(f"[updater] {message}")


def _ssl_contexts():
    """Yields the TLS setups to try, in order, as (description, context).

    Two different environments fail in opposite ways, so neither choice
    works everywhere and the fix is to try both:

      * Under WINE, Python verifies against the Windows certificate store,
        which a fresh prefix leaves essentially empty. GitHub's perfectly
        valid certificate is rejected and it looks exactly like being
        offline. certifi's bundled CA list fixes this.

      * Behind a CORPORATE TLS-INSPECTING PROXY, traffic is re-signed with
        an internal CA that is installed in the system store and is,
        by design, absent from certifi. Here certifi is the one that fails
        and the system store is what works.

    Verification is never disabled in either case: this downloads code that
    is about to be executed, so an unverified fetch would be far worse than
    a failed update."""
    try:
        import certifi
        yield "bundled CA list (certifi)", ssl.create_default_context(cafile=certifi.where())
    except Exception:
        pass
    # Also picks up SSL_CERT_FILE, which is how proxies and some corporate
    # images point Python at their own CA.
    yield "system certificate store", ssl.create_default_context()


# Set by _download_latest so the error dialog can explain what actually went
# wrong instead of blaming the network for everything.
LAST_ERROR = None


def _download_latest():
    """Fetches the current converter source from GitHub, or returns None if
    it can't (offline, token expired, certificate problem, repo renamed...).

    Sets LAST_ERROR to a human-readable explanation on failure, so the
    dialog can say what actually happened instead of blaming the network.

    raw.githubusercontent.com is used rather than the REST API because the
    repo is public and the raw host has no meaningful rate limit, whereas
    the unauthenticated API allows only 60 requests per hour PER IP - an
    office where everyone shares one public IP could exhaust that in a
    morning. A cache-busting query string avoids the CDN serving a stale
    copy right after a push."""
    global LAST_ERROR
    LAST_ERROR = None

    url = (
        f"https://raw.githubusercontent.com/{GITHUB_OWNER}/{GITHUB_REPO}/"
        f"{GITHUB_BRANCH}/{PAYLOAD_PATH_IN_REPO}?t={int(time.time())}"
    )
    request = urllib.request.Request(url)
    request.add_header("User-Agent", f"altment-converter-launcher/{LAUNCHER_VERSION}")
    request.add_header("Cache-Control", "no-cache")

    ssl_failures = []
    for description, context in _ssl_contexts():
        try:
            with urllib.request.urlopen(
                request, timeout=UPDATE_TIMEOUT_SECONDS, context=context
            ) as response:
                return response.read().decode("utf-8")

        except urllib.error.HTTPError as e:
            # GitHub answered, so the network and TLS are both fine - this is
            # a configuration problem, and "you're offline" would send you
            # hunting in entirely the wrong place.
            LAST_ERROR = {
                404: f"GitHub can't find {GITHUB_OWNER}/{GITHUB_REPO} at branch "
                     f"'{GITHUB_BRANCH}', path '{PAYLOAD_PATH_IN_REPO}'. Check "
                     "the CONFIGURATION block - and check the repo is PUBLIC, "
                     "since a private one is invisible without credentials and "
                     "reports 404 rather than a permissions error.",
                429: "GitHub is rate-limiting this network. Try again shortly.",
            }.get(e.code, f"GitHub returned HTTP {e.code}.")
            break

        except urllib.error.URLError as e:
            # Only a certificate failure is worth retrying against a
            # different trust store; DNS or refused-connection errors would
            # fail identically and just double the wait.
            reason = getattr(e, "reason", e)
            if isinstance(reason, ssl.SSLError):
                ssl_failures.append(f"{description}: {reason}")
                continue
            LAST_ERROR = f"Could not reach GitHub: {reason}"
            break

        except (TimeoutError, OSError) as e:
            LAST_ERROR = f"Could not reach GitHub: {e}"
            break

    else:
        # Every trust store was tried and all of them failed on the
        # certificate itself.
        LAST_ERROR = (
            "The secure connection to GitHub failed with every certificate "
            "store tried, so this is a certificate problem rather than a "
            "connectivity one:\n  - " + "\n  - ".join(ssl_failures)
        )

    _log(LAST_ERROR)
    return None


def _diagnose():
    """`AltmentConverter.exe --diagnose` prints everything needed to tell a
    misconfiguration apart from a genuine network problem, without anyone
    having to read a traceback."""
    print("=" * 62)
    print("Altment converter - update diagnostics")
    print("=" * 62)
    print(f"launcher version : {LAUNCHER_VERSION}")
    print(f"frozen (.exe)    : {bool(getattr(sys, 'frozen', False))}")
    print(f"base folder      : {BASE_DIR}")
    print(f"cached script    : {CACHED_SCRIPT} "
          f"({'present' if os.path.isfile(CACHED_SCRIPT) else 'MISSING'})")
    print(f"repo             : {GITHUB_OWNER}/{GITHUB_REPO} @ {GITHUB_BRANCH}")
    print(f"path in repo     : {PAYLOAD_PATH_IN_REPO}")
    config_path = os.path.join(BASE_DIR, "config.json")
    if os.path.isfile(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                key = json.load(f).get("mailchimp_api_key", "")
            print(f"config.json      : found, mailchimp key "
                  f"{'present (...%s)' % key[-6:] if key else 'MISSING from the file'}")
        except Exception as e:
            print(f"config.json      : found but unreadable ({e})")
    else:
        print(f"config.json      : NOT FOUND at {config_path} - image upload "
              "will fail until it's created")

    try:
        import certifi
        print(f"CA bundle        : certifi at {certifi.where()}")
    except Exception:
        print("CA bundle        : certifi NOT bundled - using the system store "
              "(this is what breaks under Wine)")

    print("-" * 62)
    print("Contacting GitHub...")
    source = _download_latest()
    if source is None:
        print(f"RESULT: FAILED - {LAST_ERROR}")
    else:
        print(f"RESULT: OK - downloaded {len(source):,} characters.")
    print("=" * 62)
    input("Press Enter to close...")


def _is_valid_python(source):
    """Compiles the downloaded source without running it.

    This is the safety net that makes push-to-deploy tolerable: a typo
    pushed at 18:00 would otherwise reach every coworker's next run. A file
    that doesn't compile is rejected and the previous working copy is kept."""
    try:
        compile(source, CACHED_SCRIPT, "exec")
        return True
    except SyntaxError as e:
        _log(f"Downloaded update has a syntax error (line {e.lineno}) - ignoring it.")
        return False


def _required_launcher_version(source):
    """Reads an optional '# requires-launcher: N' marker from the converter."""
    for line in source.splitlines()[:80]:
        if line.strip().startswith("# requires-launcher:"):
            try:
                return int(line.split(":", 1)[1].strip())
            except ValueError:
                return 0
    return 0


def _read_state():
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


def _write_state(state):
    try:
        with open(STATE_PATH, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
    except OSError:
        pass  # the state file is a convenience, never load-bearing


def _install(source):
    """Writes the new source to the cache atomically.

    Atomically because the alternative - truncating the real file and
    writing into it - leaves a half-written script on disk if the machine
    is shut down or the write fails midway, and that IS the file the next
    run executes. os.replace swaps it in one step instead."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    tmp_path = CACHED_SCRIPT + ".tmp"
    with open(tmp_path, "w", encoding="utf-8", newline="") as f:
        f.write(source)
    os.replace(tmp_path, CACHED_SCRIPT)


def _sync():
    """Brings the cached converter up to date. Returns the source to run,
    or None if there's nothing runnable at all (first launch with no
    network)."""
    cached_source = None
    if os.path.isfile(CACHED_SCRIPT):
        try:
            with open(CACHED_SCRIPT, "r", encoding="utf-8") as f:
                cached_source = f.read()
        except OSError:
            cached_source = None

    if "--no-update" in sys.argv:
        _log("Update check skipped (--no-update).")
        return cached_source

    _log("Checking for updates...")
    latest = _download_latest()
    if latest is None:
        return cached_source

    if cached_source is not None and latest == cached_source:
        _log("Already up to date.")
        return cached_source

    if not _is_valid_python(latest):
        return cached_source

    needed = _required_launcher_version(latest)
    if needed > LAUNCHER_VERSION:
        _log(
            f"This update needs launcher version {needed} (this one is "
            f"{LAUNCHER_VERSION}). Ask for the new .exe - continuing with the "
            "version already on this PC."
        )
        return cached_source

    try:
        _install(latest)
    except OSError as e:
        _log(f"Could not save the update ({e}). Continuing with the current version.")
        return cached_source

    digest = hashlib.sha256(latest.encode("utf-8")).hexdigest()[:12]
    _write_state({"sha256": digest, "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")})
    _log("Updated to the latest version." if cached_source else "Converter downloaded.")
    return latest


def _fatal(message):
    """First-run-with-no-network is the one case the user must be told
    about explicitly: there's no cached copy to fall back on, so the tool
    simply can't start. A console line isn't enough if the .exe was built
    windowed, hence the dialog."""
    _log(message)
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Converter could not start", message)
        root.destroy()
    except Exception:
        pass
    sys.exit(1)


def main():
    if "--diagnose" in sys.argv:
        _diagnose()
        return

    source = _sync()
    if source is None:
        # Report the ACTUAL reason. The old message assumed "offline", which
        # sent people looking at their wifi when the real cause was usually a
        # token or a certificate.
        reason = LAST_ERROR or "The update could not be downloaded."
        _fatal(
            f"{reason}\n\nThere's also no copy saved on this PC yet, so the "
            "converter can't start.\n\nRun it from a terminal with "
            "--diagnose for details."
        )

    # Run the converter exactly as if it had been launched directly:
    # __name__ == "__main__" is what triggers its entry point.
    #
    # __file__ deliberately points at BASE_DIR rather than at the cache
    # subfolder. The converter derives SCRIPT_DIR from it when it isn't
    # frozen, and SCRIPT_DIR is where the Firmas photos, the signature
    # store and the Mailchimp cache live - those belong beside the .exe,
    # not buried in app/. (In the frozen build the converter uses
    # sys.executable anyway, so this only matters when testing from
    # source.) Tracebacks still name the real cached path, because that's
    # the filename compile() was given.
    namespace = {
        "__name__": "__main__",
        "__file__": os.path.join(BASE_DIR, os.path.basename(CACHED_SCRIPT)),
        "__builtins__": __builtins__,
    }
    exec(compile(source, CACHED_SCRIPT, "exec"), namespace)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        # Without this, any crash in the converter closes the console window
        # instantly and the user has nothing to send you.
        traceback.print_exc()
        print("\nSomething went wrong. Copy the text above and send it over.")
        input("Press Enter to close...")
        sys.exit(1)
