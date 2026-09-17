# Auto-update setup

Goal: **updating the converter for everyone = `git push`.** No rebuilding, no
Drive upload, no asking anyone to replace a file.

## How it works

The `.exe` your coworkers run is not the converter. It's a small launcher that
carries the Python runtime and the libraries (python-docx, lxml, tkinter), and
on each launch it downloads the latest `convert_to_html2.py` from your GitHub
repo, saves it next to itself in `app/`, and runs it.

That means the `.exe` almost never changes. You rebuild it only if:

- the converter starts using a library that isn't bundled yet (add the import
  to the `BUNDLED DEPENDENCIES` block in `launcher.py` first), or
- you edit `launcher.py` itself.

If GitHub is unreachable, the copy from the last run is used, so the tool keeps
working offline. Only the very first run needs network.

---

## 1. Create the repo

Make a **public** repo called e.g. `altment-converter`, with this layout:

```
convert_to_html2.py          <- the converter (the file that gets updated)
launcher.py                  <- the bootstrap that becomes the .exe
.github/workflows/build-launcher.yml
.gitignore
```

`.gitignore`:

```
app/
email_images/
Firmas/
.upload_cache.json
.custom_signatures.json
.signature_photo_overrides.json
__pycache__/
dist/
build/
*.spec
```

Those are per-machine runtime files. `Firmas/` is listed because signature
photos are uploaded through the dialog and shouldn't be overwritten by an
update — if you'd rather ship a standard set of photos to everyone, remove that
line and commit them (the launcher only ever replaces `convert_to_html2.py`,
so you'd need to distribute the folder once alongside the `.exe`).

## 2. Keep the Mailchimp key out of the code

This is the step that makes a public repo safe. `convert_to_html2.py` reads the
key at runtime from a `config.json` next to the `.exe`:

```json
{"mailchimp_api_key": "abc123...-us16"}
```

`config.json` is the first entry in `.gitignore`. Never remove it, and never
paste a key back into the script — a public repo is scraped for credentials
within minutes of a push.

`MAILCHIMP_API_KEY` as an environment variable also works and takes precedence,
which is handy on a build or test machine.

## 3. Configure the launcher

In `launcher.py`, edit the CONFIGURATION block:

```python
GITHUB_OWNER = "your-org-or-username"
GITHUB_REPO  = "altment-converter"
GITHUB_BRANCH = "main"
PAYLOAD_PATH_IN_REPO = "convert_to_html2.py"
```

There is no token to set. The converter is fetched anonymously from
`raw.githubusercontent.com`, which — unlike the REST API's 60-requests-per-hour
anonymous limit, shared across everyone behind one office IP — has no limit
you'll hit in practice.

## 4. Build the .exe (no more wine)

The included workflow builds it on a real Windows runner, with no secrets
involved. Push it, then go to the **Actions** tab → "Build launcher exe" →
**Run workflow**. When it finishes, download the `AltmentConverter` artifact —
that's your `.exe`.

It also runs automatically whenever `launcher.py` changes, which is exactly
when a rebuild is needed.

If you prefer to keep building locally with wine, that still works:

```bash
wine pyinstaller --onefile --name AltmentConverter \
    --collect-all docx --collect-all lxml --collect-all certifi launcher.py
```

No token injection step any more — `launcher.py` builds as-is.

`--collect-all docx` matters: python-docx ships a default template file that
PyInstaller doesn't pick up on its own. `--collect-all certifi` matters too —
the CA bundle is a data file, and without it the updater can't verify GitHub's
certificate under Wine.

## 5. Hand it out once

Send everyone the new `AltmentConverter.exe` plus, **separately and privately**,
the Mailchimp key for their `config.json`. Point them at `INSTALL.md`. They put
the `.exe` where the old one was (same folder, so `Firmas/` and the caches are
found) and delete the old `.exe`.

That's the last manual distribution you should need — `config.json` is created
once per machine and is never touched by updates.

## 6. Day-to-day

```bash
# edit convert_to_html2.py, then:
git add convert_to_html2.py
git commit -m "Add link buttons to the selector"
git push
```

The next person to open the `.exe` gets it. They'll see `[updater] Updated to
the latest version.` in the console.

---

## Safeguards worth knowing about

- **A broken push can't brick everyone.** The launcher compiles the download
  before installing it; a file with a syntax error is rejected and the previous
  working copy keeps running. It does *not* catch a change that's valid Python
  but wrong, so a quick local test before pushing is still worth it.
- **Testing without pushing:** run `python launcher.py` from source, or edit
  `app/convert_to_html2.py` directly and launch with `--no-update`.
- **Rolling back** is `git revert` + push, and it reaches everyone the same way.
- **The write is atomic**, so a machine shut down mid-update doesn't end up with
  a half-written script.

## When it says it can't update

Run it from a terminal with `--diagnose`:

```
AltmentConverter.exe --diagnose        # or: wine AltmentConverter.exe --diagnose
```

That prints the repo it's pointing at, whether `config.json` was found and holds
a key, which CA bundle it's using, and the exact result of contacting GitHub. The message tells
you which of these it is:

| Message | Cause |
| --- | --- |
| "can't find owner/repo" | Wrong owner/repo/branch/path — **or the repo isn't public**, which reads as 404. |
| "rate-limiting this network" | Rare on the raw host; wait a few minutes. |
| "No Mailchimp API key found" | `config.json` missing next to the `.exe` (see INSTALL.md). |
| "secure connection failed ... every certificate store" | Certificate problem, not connectivity. Under Wine, run `winetricks cert` or rebuild with `--collect-all certifi`. |
| "Could not reach GitHub: ..." | Genuinely a network/DNS/firewall issue. |

Note the launcher tries the bundled CA list first and then the system store, so
it works both under Wine (empty Windows store) and behind a corporate
TLS-inspecting proxy (internal CA that certifi doesn't know).

## Security notes

**The old key must be rotated.** Any key that was ever committed — even to a
private repo, even in a commit you later amended — should be considered
compromised and regenerated in Mailchimp. The same goes for a key inside any
`.exe` you've already distributed.

**This repo must contain no credentials, ever.** `config.json` is gitignored;
keep it that way. If a key does get committed, rotating it in Mailchimp is the
real fix — deleting the file in a later commit does not remove it from history.

**A Mailchimp key is powerful.** It's not scoped to the File Manager: it can
read and modify audiences and send campaigns. Treat it like a password, send it
over a password manager or private chat rather than email, and rotate it when
someone leaves.
