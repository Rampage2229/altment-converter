# Installing on a new computer

Takes about five minutes. You need two things from whoever maintains the tool:

1. `AltmentConverter.exe`
2. the Mailchimp API key (sent privately — chat or password manager, **not**
   email and never committed anywhere)

## 1. Make a folder

Create a folder for the tool, e.g. `Documents\AltmentConverter`. Everything
lives here. Put `AltmentConverter.exe` in it.

## 2. Add the key

In that same folder, create a file called **`config.json`** containing:

```json
{
  "mailchimp_api_key": "paste-the-key-here-us16"
}
```

The key ends in something like `-us16` — that suffix tells the tool which
Mailchimp data center to use, so don't trim it.

> On Windows, make sure the file is really called `config.json` and not
> `config.json.txt`. In File Explorer: View → Show → File name extensions, then
> check the name.

## 3. Add the signature images

Copy the **`Firmas`** folder into the same place, so you end up with:

```
AltmentConverter\
    AltmentConverter.exe
    config.json
    Firmas\
        FIRMA_A.png
        ...
```

You can also skip this and add signatures yourself later from the tool's
"Add signature…" button.

## 4. Run it

Double-click `AltmentConverter.exe`. On first launch it downloads the current
converter (a second or two), then the file picker opens.

That's it. It updates itself from then on — never download or replace it again
unless you're told to.

## Checking it worked

Run it from a terminal with `--diagnose`:

```
AltmentConverter.exe --diagnose
```

It reports whether `config.json` was found, whether the key is in it, and
whether GitHub is reachable.

## If something's wrong

| What you see | What it means |
| --- | --- |
| "No Mailchimp API key found" | `config.json` is missing, in the wrong folder, or named `config.json.txt`. |
| "key looks malformed" | The `-usNN` suffix got cut off when pasting. |
| "can't find … at branch" | The `.exe` is misconfigured — tell the maintainer. |
| "secure connection failed" | Certificate/proxy issue on this network. |
| "Could not reach GitHub" | No internet, or a firewall is blocking it. It runs the last downloaded version instead. |

Anything else: run with `--diagnose`, copy the whole output, and send it over.

## Moving to another PC

Copy the whole folder. `config.json`, `Firmas` and your added signatures come
with it.
