# Altment converter

Converts the monthly `.docx` factsheet emails into Mailchimp-ready HTML:
inline styling, accent colours per gestora, signature blocks, disclaimers,
webinar and link buttons, and automatic image upload to Mailchimp's File
Manager (which is necessary because email clients strip embedded images).

**Installing it on a computer:** see [INSTALL.md](INSTALL.md).

## No credentials in this repository

The Mailchimp API key is **not** in the code. The converter reads it at runtime
from a `config.json` sitting next to the executable, which is listed in
`.gitignore` and distributed to each machine once during setup.

That separation is what lets this repo be public while the tool still
auto-updates. If you're adding features, keep it that way — never paste a key
into `convert_to_html2.py`.

## How updates reach people

`AltmentConverter.exe` is a small launcher that carries the Python runtime and
libraries. On each launch it downloads `convert_to_html2.py` from this repo,
verifies it compiles, caches it next to itself, and runs it. So deploying to
everyone is:

```bash
# edit convert_to_html2.py, test locally, then:
git add convert_to_html2.py
git commit -m "what changed"
git push
```

The next person to open the `.exe` gets it. If GitHub is unreachable they run
the last cached version, so the tool still works offline.

**Testing before pushing:**

```bash
python convert_to_html2.py            # run the converter directly
python launcher.py --no-update        # run the cached copy via the launcher
python launcher.py --diagnose         # check update + config plumbing
```

**Rolling back:** `git revert <commit>` + push.

**Rebuilding the `.exe`** is only needed when `launcher.py` changes, or when the
converter starts importing a library that isn't bundled — add the import to the
`BUNDLED DEPENDENCIES` block in `launcher.py` first. GitHub Actions builds it;
download the artifact from the Actions tab.

Full setup notes: [SETUP_AUTOUPDATE.md](SETUP_AUTOUPDATE.md).

## Layout

| File | What it is |
| --- | --- |
| `convert_to_html2.py` | The converter. **This is the file that auto-updates.** |
| `launcher.py` | The bootstrap that becomes the `.exe`. Rarely changes. |
| `config.example.json` | Template for the per-machine `config.json`. |
| `INSTALL.md` | What to send someone setting up a new computer. |
| `.github/workflows/build-launcher.yml` | Builds the `.exe` on a Windows runner. |
