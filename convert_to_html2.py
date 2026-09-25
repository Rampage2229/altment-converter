import os
import sys
import copy
import tkinter as tk
from tkinter import filedialog, ttk, colorchooser, messagebox
import tkinter.scrolledtext as scrolledtext
import base64
import json
import hashlib
import re
import shutil
import urllib.request
import urllib.error
from urllib.parse import quote as _url_quote

from lxml import etree
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

# SIGNATURE CONFIGURATION: the entries in SIGNATURES below are the built-in
# ones shipped with this script. Each entry has "label" (list text), "image"
# (path inside Firmas/), and "lines" (contact lines, each with "text",
# optional "bold", "url", "append"/"appends").
#
# You normally do NOT need to edit this dict any more: the selector dialog
# has "Add signature...", "Remove selected" and "Upload photo..." buttons.
# Signatures added from the UI are stored in .custom_signatures.json next to
# the script (see _load_signature_store) and merged into SIGNATURES on every
# run, so they survive restarts - and they keep working in the frozen .exe
# build too, since that JSON lives beside the executable rather than inside
# the code.

# In a normal `python convert_to_html2.py` run, __file__ points at this
# script and SCRIPT_DIR is simply the folder it lives in. But once this is
# frozen into a standalone .exe with PyInstaller (especially --onefile
# mode), __file__ no longer points anywhere near the .exe - PyInstaller's
# bootloader unpacks everything into a fresh temporary folder on every
# launch (sys._MEIPASS) and __file__ resolves inside THAT, which is (a)
# not next to the .exe or the .docx and (b) deleted automatically when the
# program exits. That's why email_images/, the Firmas photo overrides, and
# the Mailchimp upload cache would silently vanish in the .exe build even
# though HTML conversion itself still worked fine. PyInstaller sets
# sys.frozen at runtime specifically so code can detect this and anchor to
# sys.executable's folder instead.
if getattr(sys, "frozen", False):
    SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# =====================================================================
# GLOBAL STYLE SETTINGS - tweak the look of the generated HTML here.
# =====================================================================
BODY_FONT_FAMILY = "Arial, sans-serif"
# ---------------------------------------------------------------------
# COLOR POLICY
# ---------------------------------------------------------------------
# No color is written into this file except:
#   (a) the accent presets in ACCENT_COLOR_PRESETS / BULLET_ICON_COLOR,
#       which are exactly the palette offered in the selector dialog, and
#   (b) the pure greyscale neutrals right below - body text, page
#       background - which are not brand colors and must stay neutral so
#       the accent is the only thing that changes from gestora to gestora.
# Everything else that has a hue (links, headings, bullets, table borders
# and header tint, disclaimer text, the webinar button...) is DERIVED at
# render time from the accent color chosen in the dialog, via
# _lighten_hex_color / _mute_hex_color / _readable_text_color. Don't
# reintroduce fixed hex values: picking a different accent should re-skin
# the whole email on its own.
# ---------------------------------------------------------------------
NEUTRAL_TEXT_COLOR = "#333333"       # body copy - greyscale on purpose
UI_MUTED_TEXT_COLOR = "#666666"      # tkinter dialog chrome only (never
UI_DISABLED_BG_COLOR = "#f7f7f7"     # reaches the generated HTML) - both
                                      # greyscale, same policy as above
PAGE_BACKGROUND_COLOR = "#ffffff"    # the email's own background; used for
                                      # the hairline gaps between chart bars,
                                      # so it must match whatever background
                                      # the email is sent on
BODY_FONT_COLOR = NEUTRAL_TEXT_COLOR
BODY_LINE_HEIGHT = 1.6
BODY_MAX_WIDTH = "960px"             # OUTER page width: the widest anything in
                                      # the email is allowed to be. Only
                                      # full-width images (see
                                      # LARGE_IMAGE_MIN_WIDTH_PX) actually use
                                      # it all; text is narrower - see
                                      # TEXT_MAX_WIDTH.
TEXT_MAX_WIDTH = "760px"             # INNER column for everything that is read
                                      # rather than looked at: paragraphs,
                                      # lists, headings, tables, signature and
                                      # disclaimer. Keeping this smaller than
                                      # BODY_MAX_WIDTH is what lets a large
                                      # screenshot spill wider than the text
                                      # without turning the copy into
                                      # uncomfortably long lines. Set it equal
                                      # to BODY_MAX_WIDTH for the old
                                      # single-width behaviour.
IMAGE_BLOCK_CLASS = "image-block"    # class given to a paragraph whose only
                                      # content is an image, so it can use the
                                      # full BODY_MAX_WIDTH instead of being
                                      # boxed into TEXT_MAX_WIDTH
TEXT_COLUMN_CLASS = "text-column"    # wrapper <div> the text column is built
                                      # from - see _wrap_in_text_column(). The
                                      # column is a real container rather than
                                      # a max-width applied to each paragraph,
                                      # because Word's own indents arrive as
                                      # inline margin-left styles and an inline
                                      # style always beats a stylesheet rule:
                                      # every indented bullet would break out
                                      # of the column and stick to the left
                                      # edge of the page
BODY_PADDING = "20px"
BODY_TEXT_ALIGN = "justify"          # all body text is justified
BODY_FONT_SIZE_FALLBACK_PT = 10      # used only if a doc's own default size can't be read.
                                      # Text sizes are taken from the .docx as-is:
                                      # the Word document is the single source of
                                      # truth for how big anything is.


# --------------------------------------------------------------------
# Updated heading sizes – make titles larger than the original defaults.
# Adjust the point sizes as needed; here we increase each level by ~30%.
# --------------------------------------------------------------------
HEADING_FONT_SIZES_PT = {
    1: 32,   # was 24
    2: 24,   # was 18
    3: 18,   # was 14
    4: 14,   # was 12
    5: 12,   # was 10
    6: 10,   # was 8
}

# Headings are always rendered bold, no matter how the run was formatted in
# the source .docx - a title line should read as a title in every document,
# not just in the ones whose author remembered to bold it.
HEADING_FONT_WEIGHT = "bold"

# --------------------------------------------------------------------
# Titles in the run's accent color instead of whatever dark blue/slate
# Word's built-in heading styles baked into the source .docx.
#
# This is deliberately generic: it always uses the accent color picked in
# the selector dialog, so the same rule re-skins any gestora's document
# (Altment teal, DJE orange, Gemway navy...) without knowing anything
# about which one it is.
#
# Two mechanisms, both needed:
#   1. Runs inside a real Word Heading paragraph (h1-h6) never write their
#      own inline color any more, so the h1..h6 CSS rule - which uses the
#      accent color - actually wins. (An inline style always beats an
#      element selector, which is why headings used to come out dark blue
#      even though the CSS said otherwise.)
#   2. Paragraphs that only *look* like headings (a bold run in one of
#      Word's default dark blue/slate theme tones, e.g. a category line
#      typed as normal text) are recolored too, but ONLY when the run is
#      bold AND its color is one of the theme slots in TITLE_THEME_SLOTS.
#      Bold black labels are deliberately left alone, and so is any color
#      the author chose by hand outside the theme.
#
# Set RECOLOR_TITLES_TO_ACCENT = False for documents whose own text colors
# must be reproduced exactly; empty TITLE_THEME_SLOTS to keep mechanism 1
# (real headings) while dropping mechanism 2.
#
# Mechanism 2 identifies "heading-ish" colors from the DOCUMENT'S OWN THEME
# (theme1.xml) rather than from a list of hex values typed in here: the blue
# Word paints headings with is whatever the theme's accent1/text2 slots say,
# so reading the theme catches it in any document, in any Word language or
# template, without a single color literal in this file.
# --------------------------------------------------------------------
RECOLOR_TITLES_TO_ACCENT = True
TITLE_THEME_SLOTS = ("accent1", "text2", "dark2")

IMAGE_MAX_WIDTH = "100%"             # fallback/ceiling only - each image's actual
                                      # width normally comes from its own inline
                                      # style (read from the .docx, see
                                      # _inline_image_width_px), so a compact
                                      # banner and a deliberately large screenshot
                                      # keep their different sizes instead of
                                      # both being forced to this same width
IMAGE_MARGIN = "15px auto"           # ...and centered
TABLE_IMAGE_MAX_WIDTH = "520px"      # images inside table cells stay smaller (was 400px; raised with IMAGE_SCALE)

TABLE_BORDER_TINT = 0.55             # how far the accent is blended toward
                                      # white to get the table/cell border
                                      # color (0 = full accent, 1 = white)
TABLE_HEADER_TINT = 0.85             # same idea for the header row background
TABLE_CELL_PADDING = "3px 8px"
TABLE_BLOCK_MARGIN = "28px 0"        # vertical space around every table (was 10px 0)

# --------------------------------------------------------------------
# Native bar-chart tables.
#
# Word charts (DrawingML) and pasted Excel charts never survive this
# conversion: the former isn't a raster image at all, and the latter comes
# through as an EMF/WMF metafile that Mailchimp refuses to host (see
# UNSUPPORTED_IMAGE_EXTENSIONS). So a monthly-performance bar chart has to
# be built as a real Word TABLE and rendered here as a borderless bar.
#
# A table opts into this rendering by setting its alt-text title (Table
# Properties > Alt Text > Title in Word, or w:tblPr/w:tblCaption in the
# XML) to BAR_CHART_TABLE_MARKER.
#
# Each ROW of such a table is emitted as its own one-row <table>. That
# looks redundant but it's the whole point: CSS `table-layout: fixed`
# takes its column widths from the FIRST row only, so a single table
# could never render bars of differing lengths. One table per row means
# every bar gets its own exact width percentages, which is what makes
# the chart precise instead of quantised into visible steps. Rows still
# line up with each other because the label and value columns are given
# the same width in every row.
BAR_CHART_TABLE_MARKER = "grafico-barras"
BAR_CHART_CELL_PADDING = "0px"
BAR_CHART_ROW_HEIGHT = "15px"
BAR_CHART_LABEL_PADDING = "1px 6px 1px 0px"   # label + value columns keep a little breathing room
# Bars are separated by a hairline in the page background colour rather
# than by padding: CSS paints a cell's background across its padding box,
# so padding alone would NOT create a gap between consecutive bars.
BAR_CHART_ROW_GAP_COLOR = PAGE_BACKGROUND_COLOR

# Bullet/numbered list styling (word lists converted via w:numPr -> <ul>/<ol>)
LIST_PADDING_LEFT = "22px"
LIST_MARGIN = "6px 0"
LIST_ITEM_MARGIN = "3px 0"

# The signature block used to be introduced by a 2px rule in the accent color.
# It read as a stray orange line dropped into the middle of the email, so it's
# gone: the block is now separated by whitespace only (SIGNATURE_BLOCK_MARGIN_TOP
# + SIGNATURE_BLOCK_PADDING_TOP).
SIGNATURE_BLOCK_MARGIN_TOP = "2px"
SIGNATURE_BLOCK_PADDING_TOP = "20px"
SIGNATURE_IMAGE_MAX_WIDTH = "480px"  # was 650px - the signature read as oversized next to the body text
SIGNATURE_CELL_SPACING = "20px"  # unused now that the signature is image-only (no adjacent text cell)
SIGNATURE_FONT_SIZE = "10pt"  # unused now that the signature is image-only (no adjacent text cell)

DISCLAIMER_BLOCK_MARGIN_TOP = "20px"
DISCLAIMER_FONT_SIZE = "7pt"          # was 9pt - smaller, less obtrusive legal text
DISCLAIMER_LINE_HEIGHT = 1.3          # was the hardcoded 1.4 in the <style> block
DISCLAIMER_MUTE = 0.75               # how far the accent is muted toward mid
                                      # grey for the small print (0 = full
                                      # accent, 1 = plain grey) - legible and
                                      # discreet, but still tinted by the
                                      # gestora's color
DISCLAIMER_GAP = "16px"              # extra space between the fixed and optional disclaimer

# Pill-shaped call-to-action buttons. Two kinds share this exact styling:
#   * the webinar "confirm attendance" button (a mailto: link), and
#   * any number of custom link buttons (a YouTube video, a landing page...),
#     added from the "Link buttons" tab of the selector dialog.
# Both are off by default and each is independent, so an email can carry the
# webinar button only, link buttons only, both, or neither.
#
# Every button's background defaults to the accent color and its label to
# black or white, whichever reads better on that background - which is what
# keeps the text legible on dark navies like Gemway's. Both can be overridden
# per run from the dialog.
BUTTON_BLOCK_MARGIN = "30px 0"
BUTTON_PADDING = "14px 30px"
BUTTON_BORDER_RADIUS = "30px"
BUTTON_FONT_SIZE = "11pt"
BUTTON_STACK_GAP = "12px"   # vertical gap between consecutive buttons
# =====================================================================

# Bullet-point marker icon. Word's native bullet glyph/color isn't reliably
# preserved across email clients (Outlook desktop in particular ignores
# ::marker/::before styling), so every bullet-type list ("- " in the
# factsheet prompt) is rendered with this literal character + color instead
# of relying on the browser's default disc marker. The color picked here
# (in the selector dialog) doubles as the general ACCENT color for the rest
# of the generated email - links, headings, the table header tint, and the
# rule above the signature all pick it up too, so one color choice re-skins
# the whole email to match whichever gestora's factsheet it is. This
# constant is just the fallback/default if nothing is chosen in the dialog.
BULLET_ICON_CHAR = "\u25c6"  # ◆ black diamond
BULLET_ICON_COLOR = "#1AAE9F"  # Altment blue-green (default accent color)
BULLET_ICON_MARGIN_RIGHT = "7px"

# The accent color actually in use for the current run. It starts as the
# fallback above and is overwritten with the color picked in the selector
# dialog (DJE orange, Gemway navy, ...) as soon as that dialog closes, so
# helpers like process_run can reach it without threading it through every
# call. See RECOLOR_TITLES_TO_ACCENT.
ACCENT_COLOR = BULLET_ICON_COLOR

# Quick-select swatches shown next to the accent-color picker in the
# selector dialog, so switching between gestoras doesn't require the OS
# color-picker every time. Colors are approximate, picked to match each
# gestora's own factsheet/branding (chart lines, logo, header bars) as seen
# in their monthly emails - tweak freely if a gestora updates its palette.
# "Altment" (the blue-green house color) is first and is the overall default.
ACCENT_COLOR_PRESETS = [
    ("Altment", "#1AAE9F"),
    ("DJE", "#F08C00"),
    ("Gemway", "#1B2A5B"),
    ("IRIVEST / Chahine Funds", "#101B45"),
    ("UTI", "#F26522"),
]
# =====================================================================

# Starting values shown (and fully editable) in the webinar section of the
# selector dialog, so the form never opens blank.
WEBINAR_DEFAULTS = {
    "button_text": "CONFIRMO ASISTENCIA AL WEBINAR",
    "recipient": "cjimenez@altment.com",
    "subject": "Lennertz & Co. - WEBINAR",
    "body": "Ruego me envíen una invitación. Muchas gracias.",
}

# Starting values for each new row added in the "Link buttons" tab. The text
# is a placeholder to be edited; the URL starts empty so an unfilled row is
# obviously incomplete (rows without both a label and a URL are skipped).
LINK_BUTTON_DEFAULTS = {
    "text": "MÁS INFORMACIÓN",
    "url": "",
}

# Mailchimp is used because it (like most email clients) strips base64-embedded
# images on send, so every image gets uploaded to Mailchimp's File Manager and
# referenced by its permanent URL instead. Get/create a key at:
# Mailchimp > Account > Extras > API keys (the "-usN" suffix is the data center).
#
# The key is deliberately NOT written in this file. This script lives in a
# public repository and is downloaded over the internet on every launch, so a
# key here would be world-readable and scraped within minutes. It's read at
# runtime instead, from either:
#
#   1. the MAILCHIMP_API_KEY environment variable, or
#   2. a config.json sitting NEXT TO the .exe (never in git - see .gitignore):
#
#          {"mailchimp_api_key": "abc123...-us16"}
#
# config.json is handed to each machine once during setup and is never
# touched by updates, which is the whole point: the code can be public and
# updated freely while the credential stays local.
MAILCHIMP_CONFIG_FILENAME = "config.json"


def _load_mailchimp_api_key():
    """Returns the Mailchimp API key, or "" if none is configured.

    Checked at import time so a missing key is reported clearly by
    upload_image_to_mailchimp rather than as an opaque 401 from Mailchimp
    halfway through a conversion."""
    from_env = os.environ.get("MAILCHIMP_API_KEY", "").strip()
    if from_env:
        return from_env

    config_path = os.path.join(SCRIPT_DIR, MAILCHIMP_CONFIG_FILENAME)
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
    except FileNotFoundError:
        return ""
    except (json.JSONDecodeError, OSError) as e:
        print(f"Warning: could not read {config_path} ({e}).")
        return ""

    if not isinstance(config, dict):
        return ""
    return str(config.get("mailchimp_api_key", "")).strip()


MAILCHIMP_API_KEY = _load_mailchimp_api_key()

# Extensions Mailchimp's File Manager API rejects outright (confirmed via its
# "files with the .x-emf extension are not allowed" error). These show up as
# the content_type of an embedded Word blip when something (usually a chart
# or object pasted from Excel/PowerPoint) was stored as a vector metafile
# instead of a real raster image - there's no raster fallback to use, so
# these get skipped rather than crashing the whole conversion. Add more here
# if Mailchimp ever rejects another format.
UNSUPPORTED_IMAGE_EXTENSIONS = {"x-emf", "emf", "x-wmf", "wmf"}

# Folder (next to this script, shared across every .docx you convert) where a
# reference copy of every uploaded image is kept.
IMAGE_EXPORT_DIRNAME = "email_images"

# Cache of image hash -> hosted URL, so re-running doesn't re-upload unchanged images.
UPLOAD_CACHE_PATH = os.path.join(SCRIPT_DIR, ".mailchimp_upload_cache.json")


def _load_upload_cache():
    if os.path.isfile(UPLOAD_CACHE_PATH):
        try:
            with open(UPLOAD_CACHE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def _save_upload_cache(cache):
    with open(UPLOAD_CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2)


# Records which photo (filename inside Firmas/) each signature key should use,
# once someone has uploaded a replacement via the selector's "Upload photo"
# button. This is what makes an uploaded photo "stick" for future runs
# without touching the SIGNATURES dict in this file.
SIGNATURE_PHOTO_OVERRIDES_PATH = os.path.join(SCRIPT_DIR, ".signature_photo_overrides.json")


def _load_signature_photo_overrides():
    if os.path.isfile(SIGNATURE_PHOTO_OVERRIDES_PATH):
        try:
            with open(SIGNATURE_PHOTO_OVERRIDES_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def _save_signature_photo_overrides(overrides):
    with open(SIGNATURE_PHOTO_OVERRIDES_PATH, "w", encoding="utf-8") as f:
        json.dump(overrides, f, indent=2)


# Signatures created from the selector's "Add signature..." button, plus the
# built-in keys the user chose to hide with "Remove selected". Kept in a JSON
# file next to the script (or next to the .exe when frozen) so the list of
# signatures can be managed entirely from the UI, with no code edits and no
# loss when the script is updated.
#
# Shape:
#   {
#     "signatures": {
#        "MI_FIRMA": {"label": "...", "image_filename": "FIRMA_MI_FIRMA.png",
#                     "lines": [{"text": "..."}]}
#     },
#     "hidden_builtins": ["B"]
#   }
#
# Built-ins are HIDDEN rather than deleted (their definitions live in the code
# above and can't be erased from a JSON file), which is also what makes
# "Restore hidden built-ins" possible.
CUSTOM_SIGNATURES_PATH = os.path.join(SCRIPT_DIR, ".custom_signatures.json")


def _empty_signature_store():
    return {"signatures": {}, "hidden_builtins": []}


def _load_signature_store():
    if not os.path.isfile(CUSTOM_SIGNATURES_PATH):
        return _empty_signature_store()
    try:
        with open(CUSTOM_SIGNATURES_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return _empty_signature_store()

    if not isinstance(data, dict):
        return _empty_signature_store()

    signatures = data.get("signatures")
    hidden = data.get("hidden_builtins")
    return {
        "signatures": signatures if isinstance(signatures, dict) else {},
        "hidden_builtins": list(hidden) if isinstance(hidden, list) else [],
    }


def _save_signature_store(store):
    with open(CUSTOM_SIGNATURES_PATH, "w", encoding="utf-8") as f:
        json.dump(store, f, indent=2, ensure_ascii=False)


def _mailchimp_datacenter(api_key):
    if "-" not in api_key:
        raise ValueError(
            "The Mailchimp API key looks malformed (expected something like "
            f"'abc123...-us4'). Check the value in {MAILCHIMP_CONFIG_FILENAME} "
            "next to the program."
        )
    return api_key.rsplit("-", 1)[-1]


def upload_image_to_mailchimp(image_bytes, filename):
    """Uploads image_bytes to Mailchimp's File Manager and returns its
    permanent hosted URL (the 'full_size_url' Mailchimp assigns). Uses a
    local hash-based cache so identical bytes are never uploaded twice."""
    file_hash = hashlib.sha256(image_bytes).hexdigest()
    cache = _load_upload_cache()
    if file_hash in cache:
        print(f"Already uploaded to Mailchimp (cached): {filename} -> {cache[file_hash]}")
        return cache[file_hash]

    if not MAILCHIMP_API_KEY:
        raise RuntimeError(
            "No Mailchimp API key found, so images can't be uploaded.\n\n"
            f"Create a file named {MAILCHIMP_CONFIG_FILENAME} next to the "
            f"program (in {SCRIPT_DIR}) containing:\n\n"
            '    {"mailchimp_api_key": "your-key-here-us16"}\n\n'
            "Get the key from Mailchimp > Account > Extras > API keys. "
            "(Setting a MAILCHIMP_API_KEY environment variable also works.)"
        )

    dc = _mailchimp_datacenter(MAILCHIMP_API_KEY)
    api_url = f"https://{dc}.api.mailchimp.com/3.0/file-manager/files"
    payload = json.dumps({
        "name": filename,
        "file_data": base64.b64encode(image_bytes).decode("utf-8"),
    }).encode("utf-8")

    auth_token = base64.b64encode(f"anystring:{MAILCHIMP_API_KEY}".encode()).decode()
    req = urllib.request.Request(api_url, data=payload, method="POST")
    req.add_header("Authorization", f"Basic {auth_token}")
    req.add_header("Content-Type", "application/json")

    try:
        with urllib.request.urlopen(req) as resp:
            result = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"Mailchimp upload failed for '{filename}': HTTP {e.code} - {error_body}"
        ) from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Could not reach Mailchimp to upload '{filename}': {e}") from e

    hosted_url = result.get("full_size_url")
    if not hosted_url:
        raise RuntimeError(
            f"Mailchimp upload for '{filename}' succeeded but returned no URL: {result}"
        )

    cache[file_hash] = hosted_url
    _save_upload_cache(cache)
    print(f"Uploaded to Mailchimp: {filename} -> {hosted_url}")
    return hosted_url


def _resolve_firmas_dir():
    """Finds the signatures folder next to the script, regardless of case
    (e.g. 'Firmas', 'FIRMAS', 'firmas') since filesystems on Linux/macOS
    are case-sensitive."""
    candidates = ["Firmas", "FIRMAS", "firmas"]
    for name in candidates:
        candidate_path = os.path.join(SCRIPT_DIR, name)
        if os.path.isdir(candidate_path):
            return candidate_path
    # Fallback: scan script dir for a case-insensitive match
    for entry in os.listdir(SCRIPT_DIR):
        if entry.lower() == "firmas" and os.path.isdir(os.path.join(SCRIPT_DIR, entry)):
            return os.path.join(SCRIPT_DIR, entry)
    # Nothing found: default to the most common spelling so the warning
    # message below shows a sensible expected path.
    return os.path.join(SCRIPT_DIR, "Firmas")


FIRMAS_DIR = _resolve_firmas_dir()
SIGNATURES = {
    "A": {
        "label": "Signature A - Madrid (Alba / Chamorro / Martín) - photo",
        "image": os.path.join(FIRMAS_DIR, "FIRMA_A.png"),
        "lines": [
            {"text": "JOSÉ ALBA, Managing Partner ", "bold": True, "append": {"text": "M - 629 217 130", "url": "tel:+34629217130"}},
            {"text": "E-mail: ", "append": {"text": "jalba@altment.com", "url": "mailto:jalba@altment.com"}},

            {"text": "SERGI MARTÍN AMORÓS, CAIA ", "bold": True, "appends": [
                {"text": "M - +376 362 665", "url": "tel:+376362665"},
                {"text": "T - 932 556 159", "url": "tel:+34932556159"},
            ]},
            {"text": "E-mail: ", "append": {"text": "smartin@altment.com", "url": "mailto:smartin@altment.com"}},

            {"text": "JAVIER CHAMORRO, Director ", "bold": True, "append": {"text": "M - 616 938 708", "url": "tel:+34616938708"}},
            {"text": "E-mail: ", "append": {"text": "jchamorro@altment.com", "url": "mailto:jchamorro@altment.com"}},

            {"text": "Paseo de la Castellana, 194"},
            {"text": "28046 Madrid, Spain"},
            {"text": "Av. Diagonal, 601, 08028, Barcelona"},
            {"text": "Spain"},
            {"text": "www.altment.com", "url": "http://www.altment.com"},
        ],
    },
    "B": {
        "label": "Signature B - Barcelona (Martín) - photo",
        "image": os.path.join(FIRMAS_DIR, "FIRMA_B.png"),
        "lines": [
            {"text": "SERGI MARTÍN ", "bold": True, "appends": [
                {"text": "M - +376 362 665", "url": "tel:+376362665"},
                {"text": "T - 932 556 159", "url": "tel:+34932556159"},
            ]},
            {"text": "Av. Diagonal, 601"},
            {"text": "08028, Barcelona, Spain"},
            {"text": "E-mail: ", "append": {"text": "smartin@altment.com", "url": "mailto:smartin@altment.com"}},
            {"text": "www.altment.com", "url": "http://www.altment.com"},
        ],
    },
    "C": {
        "label": "Signature C - Madrid (Alba / Chamorro) - generic logo",
        "image": os.path.join(FIRMAS_DIR, "FIRMA_LOGO.png"),
        "lines": [
            {"text": "JOSÉ ALBA, Managing Partner ", "bold": True, "append": {"text": "M - 629 217 130", "url": "tel:+34629217130"}},
            {"text": "E-mail: ", "append": {"text": "jalba@altment.com", "url": "mailto:jalba@altment.com"}},

            {"text": "JAVIER CHAMORRO, Director ", "bold": True, "append": {"text": "M - 616 938 708", "url": "tel:+34616938708"}},
            {"text": "E-mail: ", "append": {"text": "jchamorro@altment.com", "url": "mailto:jchamorro@altment.com"}},

            {"text": "Paseo de la Castellana, 194"},
            {"text": "28046 Madrid, Spain"},
            {"text": "Office: ", "append": {"text": "+34 932 556 159", "url": "tel:+34932556159"}},
            {"text": "www.altment.com", "url": "http://www.altment.com"},
        ],
    },
    "D": {
        "label": "Signature D - Barcelona (Chamorro / Alba / Martín) - generic logo",
        "image": os.path.join(FIRMAS_DIR, "FIRMA_LOGO.png"),
        "lines": [
            {"text": "JOSÉ ALBA, Managing Partner ", "bold": True, "append": {"text": "M - 629 217 130", "url": "tel:+34629217130"}},
            {"text": "E-mail: ", "append": {"text": "jalba@altment.com", "url": "mailto:jalba@altment.com"}},

            {"text": "SERGI MARTÍN AMORÓS, CAIA ", "bold": True, "appends": [
                {"text": "M - +376 362 665", "url": "tel:+376362665"},
                {"text": "T - 932 556 159", "url": "tel:+34932556159"},
            ]},
            {"text": "E-mail: ", "append": {"text": "smartin@altment.com", "url": "mailto:smartin@altment.com"}},

            {"text": "JAVIER CHAMORRO, Responsable of Portugal ", "bold": True, "append": {"text": "M - 616 938 708", "url": "tel:+34616938708"}},
            {"text": "E-mail: ", "append": {"text": "jchamorro@altment.com", "url": "mailto:jchamorro@altment.com"}},

            {"text": "Paseo de la Castellana, 194"},
            {"text": "28046 Madrid, Spain"},
            {"text": "Av. Diagonal, 601"},
            {"text": "08028, Barcelona, Spain"},
            {"text": "www.altment.com", "url": "http://www.altment.com"},
        ],
    },
}

# Pristine snapshot of the signatures defined in the code above. SIGNATURES
# itself is rebuilt from this snapshot every time the user adds, removes or
# restores something in the dialog, so hiding a built-in is always reversible.
_BUILTIN_SIGNATURES = copy.deepcopy(SIGNATURES)


def _slugify_signature_key(label, existing_keys):
    """Turns a user-typed label into a unique, filename-safe dict key
    (e.g. "Firma Lisboa (Alba)" -> "FIRMA_LISBOA_ALBA"). The key ends up in
    the photo filename (FIRMA_<key>.png), so it can't contain anything the
    filesystem would object to."""
    base = re.sub(r"[^A-Za-z0-9]+", "_", label).strip("_").upper()
    base = base[:40] or "FIRMA"
    if base[0].isdigit():
        base = f"F_{base}"
    key = base
    n = 2
    while key in existing_keys:
        key = f"{base}_{n}"
        n += 1
    return key


def _rebuild_signatures():
    """Repopulates SIGNATURES in place from three layers, in order:

      1. the built-in definitions above, minus any the user hid;
      2. the signatures the user added from the dialog;
      3. the photo overrides uploaded via "Upload photo..." (which win, so a
         signature always shows its most recently uploaded image).

    Mutating the existing dict rather than rebinding the name matters:
    build_signature_html and the rest of the module read the module-level
    SIGNATURES directly. Returns the loaded store so callers can reuse it."""
    store = _load_signature_store()
    hidden = set(store["hidden_builtins"])

    SIGNATURES.clear()
    for key, data in _BUILTIN_SIGNATURES.items():
        if key not in hidden:
            SIGNATURES[key] = copy.deepcopy(data)

    for key, entry in store["signatures"].items():
        if not isinstance(entry, dict):
            continue
        image_filename = entry.get("image_filename") or ""
        SIGNATURES[key] = {
            "label": entry.get("label", key),
            "image": os.path.join(FIRMAS_DIR, image_filename) if image_filename else "",
            "lines": entry.get("lines") or [],
            "custom": True,   # marks it as removable (built-ins are only hideable)
        }

    for key, override_filename in _load_signature_photo_overrides().items():
        if key in SIGNATURES:
            SIGNATURES[key]["image"] = os.path.join(FIRMAS_DIR, override_filename)

    return store


# Apply the user's saved signatures / hidden built-ins / uploaded photos from
# previous runs. The files live permanently in FIRMAS_DIR, so each signature
# automatically uses its latest uploaded photo.
_rebuild_signatures()

# DISCLAIMER CONFIGURATION: DISCLAIMER_FIXED_TEXT is always appended, no matter
# which option is chosen below (including "no additional disclaimer"). Each
# entry in DISCLAIMERS is an extra, content-specific disclaimer that gets
# appended after the fixed one. Use "\n" inside "text" for line breaks.
# TODO: replace these placeholder texts with the real disclaimer wording.
DISCLAIMER_FIXED_TEXT = (
    """This message is directed exclusively to its addressee. It contains confidential information the disclosure of which is prohibited by law. If you have received this message by mistake, please let us know by sending an email to cjimenez@altment.com or by calling +34 932 556 159 and destroy it immediately. You should
know that both the reading and the copy or any other use of it is prohibited. We would also like to inform you that, in accordance with Spanish Law 3/2018 on
Protection of Personal Data, the data used to send you this information are stored in a database. To exercise the right to access, rectify or cancel your data you
can send a message to the above address. All services provided in the European Economic Area (EEA) by Altment Capital Partners, S.L. as Agent of Solventis
S.V. S,A.  are subject to current Spanish law and come under the supervision of the Spanish Securities Commission (“CNMV”)."""
)

DISCLAIMERS = {
    "A": {
        "label": "DJE",
        "text": """This is a marketing advertisement. Please read the prospectus of the relevant fund and the PRIIPs KID before making a final investment decision. It also
contains detailed information on opportunities and risks. These documents can be obtained free of charge in German at www.dje.de, www.dws.com under the
relevant fund. A summary of investor rights can be accessed in German free of charge in electronic form on the website at www.dje.de/summary-of-investor-
rights, www.dws.com/footer/legal-resources/. The Funds described in this Marketing Announcement may have been notified for distribution in different EU
Member States. Investors should note that the relevant management company may decide to discontinue the arrangements it has made for the distribution of
the units of your funds in accordance with Directive 2009/65/EC and Article 32a of Directive 2011/61/EU. All information published here is for your information
only, is subject to change and does not constitute investment advice or any other recommendation. The sole binding basis for the acquisition of the relevant fund
is the above-mentioned documents in conjunction with the associated annual report and/or the semi-annual report. The statements contained in this document
reflect the current assessment of DJE Kapital AG. The opinions expressed may change at any time without prior notice. All information in this overview has
been provided with due care in accordance with the state of knowledge at the time of preparation. However, no guarantee or liability can be assumed for the
correctness and completeness.""",
    },
    "B": {
        "label": "IRIVEST CHAHINE",
        "text": """The future performance of an investment cannot be deduced from previous market value, I.E. the value of an investment may fall as well as rise. An investment
may lose value due to changes in rates of foreign exchange. IRIVEST Investment Managers cannot guarantee that any capital invested will maintain or increase
in value. This is a marketing communication.""",
    },
    "C": {
        "label": "UTI",
        "text": """This email message (including attachments, if any) is intended for the use of the individual or entity to which it is addressed and may contain information that is
privileged, proprietary, confidential and exempt from disclosures. Any comments or statements made herein do not necessarily reflect those of UTI International
(Singapore) Private Limited, its sponsors &amp; affiliates. Unless stated otherwise, this email is not intended to constitute a marketing communication, an offer,
invitation or solicitation to buy or sell any financial instrument or investment product. Where this email contains marketing material, the applicable regulatory
disclosures and offering documents should be referred to before making any investment decision. If you are not the intended recipient, you are notified that any
dissemination, distribution or copying of this communication is strictly prohibited. If you have received this communication in error, please notify the sender and
erase this email message immediately.""",
    },
    "D": {
        "label": "GEMWAY",
        "text": """ This document, intended for professional investors, is not of a contractual nature. It may not be reproduced, distributed or passed on to third parties in whole or
in part without the prior written authorisation of Gemway Assets SAS. The purpose of this document, which is a commercial in nature, is to inform investors of
the fund&#39;s characteristics in a simplified way. GemEquity, GemAsia, GemChina are primarily invested in equities and present a risk of capital loss. For more
information, please refer to the KIID or consult your usual contact. """,
    },
}

def _html_escape(text):
    return (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _readable_text_color(hex_color):
    """Returns black or white - whichever reads better on top of the given
    '#rrggbb' background. These two are contrast values, not palette colors:
    they're the only fixed colors allowed to appear here (see COLOR POLICY)."""
    try:
        hex_color = hex_color.lstrip("#")
        r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    except (ValueError, IndexError):
        return "#000000"
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return "#000000" if luminance > 0.6 else "#ffffff"


def _lighten_hex_color(hex_color, amount=0.85):
    """Blends '#rrggbb' toward white by `amount` (0 = no change, 1 = white).
    Used to derive soft tints of the chosen accent color (table header
    backgrounds, borders), so the accent shows up without overpowering the
    content."""
    return _blend_hex_color(hex_color, "#ffffff", amount)


def _mute_hex_color(hex_color, amount=0.6):
    """Blends '#rrggbb' toward mid grey by `amount`. Used where the accent
    would be too loud but a fixed grey would be a hardcoded color: small
    print, secondary text."""
    return _blend_hex_color(hex_color, "#808080", amount)


def _is_title_theme_color(run_color, theme_colors):
    """True when a run's color is one of the document theme's heading slots
    (see TITLE_THEME_SLOTS). That's how a paragraph that only LOOKS like a
    heading - bold text in the template's heading blue, typed as normal text -
    is recognised without hardcoding any hex value: the reference colors come
    from the .docx's own theme."""
    if not run_color or not theme_colors:
        return False
    target = run_color.lstrip("#").upper()
    return any(
        (theme_colors.get(slot) or "").lstrip("#").upper() == target
        for slot in TITLE_THEME_SLOTS
    )


def _blend_hex_color(hex_color, toward, amount):
    """Linear blend between two '#rrggbb' colors. On malformed input it falls
    back to the default accent (BULLET_ICON_COLOR) rather than to some fixed
    grey, so no color that isn't part of the accent palette can leak into the
    output."""
    def _rgb(value):
        value = (value or "").lstrip("#")
        return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))

    try:
        r, g, b = _rgb(hex_color)
    except (ValueError, IndexError):
        try:
            r, g, b = _rgb(BULLET_ICON_COLOR)
        except (ValueError, IndexError):
            return "#000000"
    try:
        tr, tg, tb = _rgb(toward)
    except (ValueError, IndexError):
        tr, tg, tb = (255, 255, 255)

    amount = min(max(amount, 0.0), 1.0)
    r = round(r + (tr - r) * amount)
    g = round(g + (tg - g) * amount)
    b = round(b + (tb - b) * amount)
    return f"#{r:02x}{g:02x}{b:02x}"


# --- File selector ---
def select_docx_files():
    print("Opening file selector...")
    root = tk.Tk()
    root.withdraw()
    root.call('wm', 'attributes', '.', '-topmost', True)

    file_paths = filedialog.askopenfilenames(
        title="Select Word Document(s) to convert to HTML",
        filetypes=[("Word Documents", "*.docx")]
    )
    root.destroy()
    return file_paths

# --- Combined signature + disclaimer + buttons selector ---
def select_email_options():
    """Shows a single, resizable dialog for picking the signature, the extra
    disclaimer, and (optionally) a webinar "confirm attendance" mailto
    button, all in one pass.

    Signature/disclaimer are scrollable lists; picking an entry shows its
    full text in a read-only preview pane below the list (word-wrapped,
    scrollable) so nothing gets cut off the way the old fixed-width buttons
    did. The webinar section is a checkbox that enables/disables a small
    form (button text, recipient, subject, body) pre-filled with sensible
    defaults but fully editable before you hit Continue.

    The "Link buttons" tab adds any number of custom buttons (a YouTube
    video, a landing page...) in the same pill format, each with its own
    label, URL and color. It's independent of the webinar toggle, so an
    email can have the webinar button only, link buttons only, both, or
    neither. The "Button colors" tab holds the two settings that apply to
    all buttons at once: one shared background color, and the label color
    (automatic black/white by contrast, or a fixed color).

    Returns (signature_key, disclaimer_key, webinar_dict, link_buttons,
    bullet_color). signature_key/disclaimer_key may be None. webinar_dict
    is either None (toggle left off) or {"enabled": True, "button_text",
    "recipient", "subject", "body", "button_color", "text_color"}.
    link_buttons is a (possibly empty) list of {"text", "url", "color",
    "text_color"}. In both, text_color of None means "choose black or
    white automatically per button". bullet_color is a hex string (e.g.
    "#1AAE9F") - despite the name, it's used as the overall accent color
    for the document: every bullet-list marker icon, link, heading, the
    table header tint and every button color derive from it."""
    print("Opening signature / disclaimer / buttons selector...")

    NO_SIGNATURE_LABEL = "No signature"
    NO_DISCLAIMER_LABEL = "No additional disclaimer"

    sig_keys = [None] + list(SIGNATURES.keys())
    sig_labels = [NO_SIGNATURE_LABEL] + [
        SIGNATURES[k].get("label", k) for k in SIGNATURES
    ]

    disc_keys = [None] + list(DISCLAIMERS.keys())
    disc_labels = [NO_DISCLAIMER_LABEL] + [
        DISCLAIMERS[k].get("label", k) for k in DISCLAIMERS
    ]

    result = {
        "signature": None,
        "disclaimer": None,
        "webinar": None,
        "link_buttons": [],
        "bullet_color": BULLET_ICON_COLOR,
    }

    root = tk.Tk()
    root.title("Select Signature, Disclaimer & Webinar Button")
    root.call('wm', 'attributes', '.', '-topmost', True)
    root.geometry("940x760")
    root.minsize(780, 640)

    tk.Label(
        root,
        text="Choose the signature and extra disclaimer to append, and "
             "optionally add webinar or link buttons. Click an item "
             "to preview it in full below its list. Use the buttons under "
             "the signature list to add, remove or re-photo signatures.",
        padx=20, pady=12, justify="left", anchor="w",
        font=("TkDefaultFont", 10, "bold"),
    ).pack(fill="x")

    columns_frame = tk.Frame(root, padx=20)
    columns_frame.columnconfigure(0, weight=1)
    columns_frame.columnconfigure(1, weight=1)
    columns_frame.rowconfigure(0, weight=0)
    columns_frame.rowconfigure(1, weight=1)
    columns_frame.rowconfigure(2, weight=1)
    columns_frame.rowconfigure(3, weight=0)

    def _build_column(parent, col, title, labels):
        tk.Label(parent, text=title, font=("TkDefaultFont", 10, "bold"), anchor="w") \
            .grid(row=0, column=col, sticky="w", pady=(6, 2), padx=(0 if col == 0 else 10, 0))

        list_frame = tk.Frame(parent)
        list_frame.grid(row=1, column=col, sticky="nsew", padx=(0 if col == 0 else 10, 0))
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        listbox = tk.Listbox(
            list_frame, exportselection=False, activestyle="dotbox",
            yscrollcommand=scrollbar.set, height=5,
        )
        scrollbar.config(command=listbox.yview)
        listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        for label in labels:
            listbox.insert("end", label)
        listbox.selection_set(0)  # default to "No signature" / "No additional disclaimer"

        preview = scrolledtext.ScrolledText(
            parent, wrap="word", height=8, font=("TkDefaultFont", 9),
            state="disabled", bg=UI_DISABLED_BG_COLOR,
        )
        preview.grid(row=2, column=col, sticky="nsew", padx=(0 if col == 0 else 10, 0), pady=(6, 0))

        return listbox, preview

    sig_listbox, sig_preview = _build_column(columns_frame, 0, "Signature", sig_labels)
    disc_listbox, disc_preview = _build_column(columns_frame, 1, "Disclaimer", disc_labels)

    def _set_preview(widget, text):
        widget.config(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.config(state="disabled")

    def _sig_preview_text(key):
        if key is None:
            return "No signature will be appended."
        data = SIGNATURES[key]
        origin = "Added from the UI" if data.get("custom") else "Built-in"
        parts = [
            f"{origin} - key: {key}",
            f"Image: {os.path.basename(data.get('image', '')) or '(none)'}",
            "",
        ]
        for line in data.get("lines", []):
            piece = line.get("text", "")
            appends = line.get("appends") or ([line["append"]] if line.get("append") else [])
            piece += "  ".join(a.get("text", "") for a in appends)
            parts.append(piece)
        return "\n".join(parts)

    def _disc_preview_text(key):
        if key is None:
            return DISCLAIMER_FIXED_TEXT.strip() + \
                "\n\n(This fixed disclaimer above is always included. " \
                "No extra disclaimer selected.)"
        return DISCLAIMER_FIXED_TEXT.strip() + "\n\n---\n\n" + DISCLAIMERS[key].get("text", "").strip()

    def _on_sig_select(_event=None):
        sel = sig_listbox.curselection()
        idx = sel[0] if sel else 0
        _set_preview(sig_preview, _sig_preview_text(sig_keys[idx]))

    def _on_disc_select(_event=None):
        sel = disc_listbox.curselection()
        idx = sel[0] if sel else 0
        _set_preview(disc_preview, _disc_preview_text(disc_keys[idx]))

    sig_listbox.bind("<<ListboxSelect>>", _on_sig_select)
    disc_listbox.bind("<<ListboxSelect>>", _on_disc_select)
    _on_sig_select()
    _on_disc_select()

    def _upload_signature_photo():
        sel = sig_listbox.curselection()
        idx = sel[0] if sel else 0
        key = sig_keys[idx]
        if key is None:
            messagebox.showwarning(
                "No signature selected",
                "Select a specific signature (not \"No signature\") before uploading a photo for it.",
                parent=root,
            )
            return

        file_path = filedialog.askopenfilename(
            title=f"Choose a photo for: {SIGNATURES[key].get('label', key)}",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif"), ("All files", "*.*")],
        )
        if not file_path:
            return

        ext = os.path.splitext(file_path)[1].lower() or ".png"
        target_filename = f"FIRMA_{key}{ext}"
        target_path = os.path.join(FIRMAS_DIR, target_filename)

        try:
            os.makedirs(FIRMAS_DIR, exist_ok=True)
            shutil.copyfile(file_path, target_path)
        except OSError as e:
            messagebox.showerror("Upload failed", f"Could not save the photo:\n{e}", parent=root)
            return

        # Use it for the rest of this run...
        SIGNATURES[key]["image"] = target_path

        # ...and remember it permanently, so future runs pick it up too
        # without any code changes.
        overrides = _load_signature_photo_overrides()
        overrides[key] = target_filename
        _save_signature_photo_overrides(overrides)

        _on_sig_select()  # refresh the preview to show the new filename
        messagebox.showinfo(
            "Photo saved",
            f"Saved to:\n{target_path}\n\n"
            f"\"{SIGNATURES[key].get('label', key)}\" will use this photo from now on.",
            parent=root,
        )

    def _refresh_signature_list(select_key=None):
        """Repopulates the signature listbox after the underlying SIGNATURES
        dict changed. sig_keys is updated IN PLACE (slice assignment) because
        _confirm and the preview callbacks close over this exact list
        object."""
        new_keys = [None] + list(SIGNATURES.keys())
        sig_keys[:] = new_keys

        sig_listbox.delete(0, "end")
        for key in new_keys:
            sig_listbox.insert(
                "end",
                NO_SIGNATURE_LABEL if key is None else SIGNATURES[key].get("label", key),
            )

        idx = new_keys.index(select_key) if select_key in new_keys else 0
        sig_listbox.selection_clear(0, "end")
        sig_listbox.selection_set(idx)
        sig_listbox.see(idx)
        _on_sig_select()

    def _selected_signature_key():
        sel = sig_listbox.curselection()
        return sig_keys[sel[0] if sel else 0]

    def _copy_photo_into_firmas(source_path, key):
        """Copies a chosen image into Firmas/ as FIRMA_<key><ext>, avoiding
        clobbering an unrelated file that happens to share the name. Returns
        the stored filename."""
        ext = os.path.splitext(source_path)[1].lower() or ".png"
        filename = f"FIRMA_{key}{ext}"
        target_path = os.path.join(FIRMAS_DIR, filename)
        n = 2
        # The key is already unique, so a collision here means an unrelated
        # file in Firmas/ happens to own that name - bump rather than
        # overwrite it. (Unless it IS the file being added, in which case
        # copying it onto itself is a no-op we can skip.)
        while os.path.isfile(target_path) and \
                os.path.abspath(source_path) != os.path.abspath(target_path):
            filename = f"FIRMA_{key}_{n}{ext}"
            target_path = os.path.join(FIRMAS_DIR, filename)
            n += 1

        if os.path.abspath(source_path) == os.path.abspath(target_path):
            return filename

        os.makedirs(FIRMAS_DIR, exist_ok=True)
        shutil.copyfile(source_path, target_path)
        return filename

    def _add_signature():
        """Modal form for creating a new signature: a label, the photo file
        (that image is what actually gets rendered into the email), and
        optional contact lines shown in the preview pane."""
        dialog = tk.Toplevel(root)
        dialog.title("Add signature")
        dialog.transient(root)
        dialog.grab_set()
        dialog.geometry("620x420")

        body = tk.Frame(dialog, padx=16, pady=12)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(3, weight=1)

        tk.Label(body, text="Label (shown in the list):", anchor="w") \
            .grid(row=0, column=0, sticky="w", pady=4)
        label_var = tk.StringVar()
        tk.Entry(body, textvariable=label_var).grid(row=0, column=1, columnspan=2, sticky="ew", pady=4)

        tk.Label(body, text="Signature image:", anchor="w") \
            .grid(row=1, column=0, sticky="w", pady=4)
        image_var = tk.StringVar()
        tk.Entry(body, textvariable=image_var, state="readonly") \
            .grid(row=1, column=1, sticky="ew", pady=4)

        def _browse_image():
            path = filedialog.askopenfilename(
                title="Choose the signature image",
                filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif"), ("All files", "*.*")],
                parent=dialog,
            )
            if path:
                image_var.set(path)

        tk.Button(body, text="Browse...", command=_browse_image) \
            .grid(row=1, column=2, sticky="w", padx=(8, 0), pady=4)

        tk.Label(
            body,
            text="Contact lines (optional, one per line - preview only,\n"
                 "the image itself is what gets sent in the email):",
            anchor="w", justify="left", fg=UI_MUTED_TEXT_COLOR,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 2))

        lines_text = tk.Text(body, height=8, wrap="word")
        lines_text.grid(row=3, column=0, columnspan=3, sticky="nsew")

        def _save_new_signature():
            label = label_var.get().strip()
            source_path = image_var.get().strip()

            if not label:
                messagebox.showwarning("Missing label", "Give the signature a label.", parent=dialog)
                return
            if not source_path or not os.path.isfile(source_path):
                messagebox.showwarning(
                    "Missing image",
                    "Choose the image file for this signature.",
                    parent=dialog,
                )
                return

            store = _load_signature_store()
            taken = set(_BUILTIN_SIGNATURES) | set(store["signatures"])
            key = _slugify_signature_key(label, taken)

            try:
                filename = _copy_photo_into_firmas(source_path, key)
            except OSError as e:
                messagebox.showerror("Could not save image", f"{e}", parent=dialog)
                return

            lines = [
                {"text": raw.strip()}
                for raw in lines_text.get("1.0", "end").splitlines()
                if raw.strip()
            ]

            store["signatures"][key] = {
                "label": label,
                "image_filename": filename,
                "lines": lines,
            }
            _save_signature_store(store)
            _rebuild_signatures()
            dialog.destroy()
            _refresh_signature_list(select_key=key)

        footer = tk.Frame(dialog, pady=10, padx=16)
        footer.pack(fill="x")
        tk.Button(footer, text="Save signature", width=16, command=_save_new_signature) \
            .pack(side="right")
        tk.Button(footer, text="Cancel", width=10, command=dialog.destroy) \
            .pack(side="right", padx=(0, 8))

        dialog.wait_window()

    def _remove_signature():
        key = _selected_signature_key()
        if key is None:
            messagebox.showwarning(
                "No signature selected",
                "Select a specific signature (not \"No signature\") to remove.",
                parent=root,
            )
            return

        label = SIGNATURES[key].get("label", key)
        store = _load_signature_store()
        is_custom = key in store["signatures"]

        if is_custom:
            question = f"Remove \"{label}\" from the list?\n\n" \
                       "Its image file stays in the Firmas folder."
        else:
            question = f"Hide the built-in signature \"{label}\"?\n\n" \
                       "It can be brought back later with \"Restore hidden built-ins\"."

        if not messagebox.askyesno("Remove signature", question, parent=root):
            return

        if is_custom:
            store["signatures"].pop(key, None)
        elif key not in store["hidden_builtins"]:
            store["hidden_builtins"].append(key)
        _save_signature_store(store)

        # Drop any uploaded-photo override for a key that no longer exists,
        # so a later signature reusing the same key doesn't inherit it.
        if is_custom:
            overrides = _load_signature_photo_overrides()
            if overrides.pop(key, None) is not None:
                _save_signature_photo_overrides(overrides)

        _rebuild_signatures()
        _refresh_signature_list()

    def _restore_hidden_builtins():
        store = _load_signature_store()
        if not store["hidden_builtins"]:
            messagebox.showinfo(
                "Nothing hidden",
                "No built-in signatures are currently hidden.",
                parent=root,
            )
            return

        count = len(store["hidden_builtins"])
        store["hidden_builtins"] = []
        _save_signature_store(store)
        _rebuild_signatures()
        _refresh_signature_list()
        messagebox.showinfo(
            "Restored",
            f"{count} built-in signature(s) restored to the list.",
            parent=root,
        )

    sig_buttons = tk.Frame(columns_frame)
    sig_buttons.grid(row=3, column=0, sticky="ew", pady=(8, 0))
    for text, command in (
        ("Add signature...", _add_signature),
        ("Remove selected", _remove_signature),
        ("Upload photo...", _upload_signature_photo),
        ("Restore hidden built-ins", _restore_hidden_builtins),
    ):
        tk.Button(sig_buttons, text=text, command=command).pack(side="left", padx=(0, 6))

    # --- Extra email features live as tabs in this notebook, not as
    # stacked full-width frames: "Webinar button", "Link buttons" and the
    # shared "Button colors". More can be added the same way - e.g. a
    # "Countdown banner" tab - without the window growing taller every
    # time: feature_notebook.add(new_tab, text="...").
    feature_notebook = ttk.Notebook(root)

    webinar_tab = tk.Frame(feature_notebook, padx=15, pady=10)
    webinar_tab.columnconfigure(1, weight=1)
    feature_notebook.add(webinar_tab, text="Webinar button")

    webinar_enabled = tk.BooleanVar(value=False)

    def _toggle_webinar_fields():
        # Enabling/disabling the webinar fields is now just one part of a
        # wider refresh (the color pickers are also governed by the "same
        # color for all buttons" option), so it delegates. The name is kept
        # because this function object is what the checkbox below binds to.
        _refresh_button_states()

    tk.Checkbutton(
        webinar_tab,
        text="Add a webinar confirmation button (opens a pre-filled email when clicked)",
        variable=webinar_enabled,
        command=_toggle_webinar_fields,
        anchor="w",
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

    button_text_var = tk.StringVar(value=WEBINAR_DEFAULTS["button_text"])
    recipient_var = tk.StringVar(value=WEBINAR_DEFAULTS["recipient"])
    subject_var = tk.StringVar(value=WEBINAR_DEFAULTS["subject"])

    tk.Label(webinar_tab, text="Button text:", anchor="w").grid(row=1, column=0, sticky="w", pady=3)
    entry_button_text = tk.Entry(webinar_tab, textvariable=button_text_var)
    entry_button_text.grid(row=1, column=1, sticky="ew", pady=3)

    tk.Label(webinar_tab, text="Recipient email:", anchor="w").grid(row=2, column=0, sticky="w", pady=3)
    entry_recipient = tk.Entry(webinar_tab, textvariable=recipient_var)
    entry_recipient.grid(row=2, column=1, sticky="ew", pady=3)

    tk.Label(webinar_tab, text="Email subject:", anchor="w").grid(row=3, column=0, sticky="w", pady=3)
    entry_subject = tk.Entry(webinar_tab, textvariable=subject_var)
    entry_subject.grid(row=3, column=1, sticky="ew", pady=3)

    tk.Label(webinar_tab, text="Email body:", anchor="nw").grid(row=4, column=0, sticky="nw", pady=3)
    body_text = tk.Text(webinar_tab, height=4, wrap="word")
    body_text.insert("1.0", WEBINAR_DEFAULTS["body"])
    body_text.grid(row=4, column=1, sticky="ew", pady=3)

    # --- Shared button appearance state ----------------------------------
    # These govern BOTH the webinar button and every custom link button, so
    # they're defined before either set of widgets and read by all of them.
    #
    # "Overridden" flags exist so a color the user deliberately picked isn't
    # silently reset when they later click a gestora preset: colors follow
    # the accent until they're chosen by hand, and then they stop.
    uniform_color_enabled = tk.BooleanVar(value=False)   # one color for all buttons
    uniform_color_var = tk.StringVar(value=BULLET_ICON_COLOR)
    auto_text_color = tk.BooleanVar(value=True)          # black/white by contrast
    text_color_var = tk.StringVar(value=_readable_text_color(BULLET_ICON_COLOR))

    def _style_swatch(widget, hex_color):
        """Paints a color-picker button with the color it represents, using
        a readable label on top - the same contrast rule the generated email
        applies, so the dialog previews what the button will look like."""
        widget.config(
            bg=hex_color, activebackground=hex_color,
            text=hex_color, fg=_readable_text_color(hex_color),
        )

    # The webinar button starts on the accent color and keeps following it
    # while the user switches gestora presets, until they deliberately pick
    # a different color from the picker.
    button_color_var = tk.StringVar(value=BULLET_ICON_COLOR)
    button_color_overridden = [False]

    def _pick_button_color():
        chosen = colorchooser.askcolor(color=button_color_var.get(), title="Choose webinar button color", parent=root)
        hex_color = chosen[1] if chosen else None
        if hex_color:
            button_color_overridden[0] = True
            button_color_var.set(hex_color)
            _style_swatch(color_swatch, hex_color)
            _refresh_button_states()

    tk.Label(webinar_tab, text="Button color:", anchor="w").grid(row=5, column=0, sticky="w", pady=3)
    color_swatch = tk.Button(
        webinar_tab, text=BULLET_ICON_COLOR, width=12,
        bg=BULLET_ICON_COLOR, activebackground=BULLET_ICON_COLOR,
        fg=_readable_text_color(BULLET_ICON_COLOR),
        command=_pick_button_color,
    )
    color_swatch.grid(row=5, column=1, sticky="w", pady=3)

    # color_swatch is deliberately NOT in this list: its state depends on
    # the webinar checkbox AND on the "same color for all buttons" option,
    # so _refresh_button_states works it out separately.
    webinar_field_widgets = [entry_button_text, entry_recipient, entry_subject, body_text]

    # =====================================================================
    # "Link buttons" tab - any number of custom buttons (YouTube video,
    # landing page, ...) rendered in exactly the same pill format as the
    # webinar button. It's fully independent of the webinar toggle, so the
    # email can carry the webinar button only, link buttons only, both, or
    # neither. A row with no label or no URL is simply ignored.
    # =====================================================================
    links_tab = tk.Frame(feature_notebook, padx=15, pady=10)
    feature_notebook.add(links_tab, text="Link buttons")
    links_tab.columnconfigure(0, weight=1)
    links_tab.rowconfigure(1, weight=1)

    tk.Label(
        links_tab,
        text="Add one button per link. Each has its own label, URL and color "
             "(a URL without http:// gets https:// added automatically).",
        anchor="w", justify="left", fg=UI_MUTED_TEXT_COLOR,
    ).grid(row=0, column=0, sticky="w", pady=(0, 6))

    # The rows live inside a canvas so a long list scrolls instead of
    # stretching the dialog: the notebook is packed with a fixed height, and
    # there's no limit on how many buttons someone might add.
    rows_container = tk.Frame(links_tab)
    rows_container.grid(row=1, column=0, sticky="nsew")
    rows_container.columnconfigure(0, weight=1)
    rows_container.rowconfigure(0, weight=1)

    rows_canvas = tk.Canvas(rows_container, height=110, highlightthickness=0)
    rows_scroll = tk.Scrollbar(rows_container, orient="vertical", command=rows_canvas.yview)
    rows_canvas.configure(yscrollcommand=rows_scroll.set)
    rows_canvas.grid(row=0, column=0, sticky="nsew")
    rows_scroll.grid(row=0, column=1, sticky="ns")

    rows_frame = tk.Frame(rows_canvas)
    rows_window = rows_canvas.create_window((0, 0), window=rows_frame, anchor="nw")

    def _sync_rows_scrollregion(_event=None):
        rows_canvas.configure(scrollregion=rows_canvas.bbox("all"))
        rows_canvas.itemconfig(rows_window, width=rows_canvas.winfo_width())

    rows_frame.bind("<Configure>", _sync_rows_scrollregion)
    rows_canvas.bind("<Configure>", _sync_rows_scrollregion)

    link_rows = []          # one dict per row: its vars + its widgets
    link_swatches = []      # every per-row color button, for the uniform-color toggle

    empty_hint = tk.Label(
        rows_frame,
        text="No link buttons yet - click \"Add link button\" below.",
        anchor="w", fg=UI_MUTED_TEXT_COLOR,
    )

    def _refresh_empty_hint():
        if link_rows:
            empty_hint.pack_forget()
        else:
            empty_hint.pack(anchor="w", pady=6)

    def _add_link_row(text=None, url=None, color=None):
        row_frame = tk.Frame(rows_frame)
        row_frame.pack(fill="x", pady=2)

        text_var = tk.StringVar(value=LINK_BUTTON_DEFAULTS["text"] if text is None else text)
        url_var = tk.StringVar(value=LINK_BUTTON_DEFAULTS["url"] if url is None else url)
        color_var = tk.StringVar(value=color or bullet_color_var.get())
        color_overridden = [color is not None]

        tk.Label(row_frame, text="Text:").pack(side="left")
        tk.Entry(row_frame, textvariable=text_var, width=22).pack(side="left", padx=(4, 8))

        tk.Label(row_frame, text="URL:").pack(side="left")
        tk.Entry(row_frame, textvariable=url_var).pack(side="left", padx=(4, 8), fill="x", expand=True)

        def _pick_row_color():
            chosen = colorchooser.askcolor(
                color=color_var.get(), title="Choose this button's color", parent=root
            )
            hex_color = chosen[1] if chosen else None
            if hex_color:
                color_overridden[0] = True
                color_var.set(hex_color)
                _style_swatch(row_swatch, hex_color)
                _refresh_button_states()

        row_swatch = tk.Button(row_frame, width=10, command=_pick_row_color)
        _style_swatch(row_swatch, color_var.get())
        row_swatch.pack(side="left", padx=(0, 6))

        row = {
            "frame": row_frame,
            "text_var": text_var,
            "url_var": url_var,
            "color_var": color_var,
            "overridden": color_overridden,
            "swatch": row_swatch,
        }

        def _remove_row():
            row_frame.destroy()
            link_rows.remove(row)
            link_swatches.remove(row_swatch)
            _refresh_empty_hint()
            _sync_rows_scrollregion()

        tk.Button(row_frame, text="Remove", command=_remove_row).pack(side="left")

        link_rows.append(row)
        link_swatches.append(row_swatch)
        _refresh_empty_hint()
        _refresh_button_states()
        _sync_rows_scrollregion()

    links_footer = tk.Frame(links_tab)
    links_footer.grid(row=2, column=0, sticky="w", pady=(8, 0))
    tk.Button(links_footer, text="Add link button", command=_add_link_row).pack(side="left")

    # =====================================================================
    # "Button colors" tab - settings that apply to EVERY button at once:
    # one shared background color, and the label color (black/white by
    # contrast, which is what keeps text legible on dark blues, or a color
    # of your choosing).
    # =====================================================================
    button_style_tab = tk.Frame(feature_notebook, padx=15, pady=10)
    feature_notebook.add(button_style_tab, text="Button colors")
    button_style_tab.columnconfigure(1, weight=1)

    def _pick_uniform_color():
        chosen = colorchooser.askcolor(
            color=uniform_color_var.get(), title="Choose the color for all buttons", parent=root
        )
        hex_color = chosen[1] if chosen else None
        if hex_color:
            uniform_color_var.set(hex_color)
            _style_swatch(uniform_swatch, hex_color)
            _refresh_button_states()

    tk.Checkbutton(
        button_style_tab,
        text="Use the same color for all buttons (webinar + link buttons)",
        variable=uniform_color_enabled,
        command=lambda: _refresh_button_states(),
        anchor="w",
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

    tk.Label(button_style_tab, text="Color for all buttons:", anchor="w") \
        .grid(row=1, column=0, sticky="w", pady=3)
    uniform_swatch = tk.Button(button_style_tab, width=12, command=_pick_uniform_color)
    _style_swatch(uniform_swatch, uniform_color_var.get())
    uniform_swatch.grid(row=1, column=1, sticky="w", pady=3)

    def _pick_text_color():
        chosen = colorchooser.askcolor(
            color=text_color_var.get(), title="Choose the button text color", parent=root
        )
        hex_color = chosen[1] if chosen else None
        if hex_color:
            text_color_var.set(hex_color)
            _style_swatch(text_color_swatch, hex_color)

    tk.Checkbutton(
        button_style_tab,
        text="Automatic text color (black or white, whichever is readable on the button)",
        variable=auto_text_color,
        command=lambda: _refresh_button_states(),
        anchor="w",
    ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 6))

    tk.Label(button_style_tab, text="Button text color:", anchor="w") \
        .grid(row=3, column=0, sticky="w", pady=3)
    text_color_swatch = tk.Button(button_style_tab, width=12, command=_pick_text_color)
    _style_swatch(text_color_swatch, text_color_var.get())
    text_color_swatch.grid(row=3, column=1, sticky="w", pady=3)

    tk.Label(
        button_style_tab,
        text="Leave the automatic option on unless you need a specific label color:\n"
             "it recomputes per button, so a dark navy button gets white text and a\n"
             "light one gets black.",
        anchor="w", justify="left", fg=UI_MUTED_TEXT_COLOR,
    ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def _effective_button_base_color():
        """The background the automatic text color should be computed
        against for the PREVIEW swatch: the shared color when one is in
        force, otherwise the accent (each real button still gets its own
        contrast decision at render time)."""
        if uniform_color_enabled.get():
            return uniform_color_var.get()
        return bullet_color_var.get()

    def _refresh_button_states():
        """Single place that decides which color controls are live.

        Per-button pickers go dead while "same color for all buttons" is on
        (leaving them clickable would let someone set a color that then
        silently gets ignored), the shared picker goes dead while it's off,
        and the text-color picker goes dead while the automatic option is
        on - in which case its swatch previews what the automatic choice
        would be."""
        webinar_state = "normal" if webinar_enabled.get() else "disabled"
        for widget in webinar_field_widgets:
            widget.config(state=webinar_state)

        uniform_on = uniform_color_enabled.get()
        color_swatch.config(
            state="normal" if (webinar_enabled.get() and not uniform_on) else "disabled"
        )
        for swatch in link_swatches:
            swatch.config(state="disabled" if uniform_on else "normal")
        uniform_swatch.config(state="normal" if uniform_on else "disabled")

        text_color_swatch.config(state="disabled" if auto_text_color.get() else "normal")
        if auto_text_color.get():
            auto_color = _readable_text_color(_effective_button_base_color())
            text_color_var.set(auto_color)
            _style_swatch(text_color_swatch, auto_color)

    # --- Accent color (always active - every factsheet uses bullets, and
    # this same color is the single source for every other hue in the email:
    # links, headings, table borders/header tint, small print, webinar
    # button) ---
    bullet_frame = tk.Frame(root)
    bullet_frame.columnconfigure(1, weight=1)

    bullet_color_var = tk.StringVar(value=BULLET_ICON_COLOR)

    def _apply_bullet_color(hex_color):
        bullet_color_var.set(hex_color)
        _style_swatch(bullet_swatch, hex_color)

        # Every button color that hasn't been deliberately overridden
        # follows the accent, so switching gestora preset re-skins the
        # whole email - buttons included - in one click.
        if not button_color_overridden[0]:
            button_color_var.set(hex_color)
            _style_swatch(color_swatch, hex_color)

        for row in link_rows:
            if not row["overridden"][0]:
                row["color_var"].set(hex_color)
                _style_swatch(row["swatch"], hex_color)

        if not uniform_color_enabled.get():
            # The shared color tracks the accent while it's switched off;
            # once it's on, it's the user's own choice and stays put.
            uniform_color_var.set(hex_color)
            _style_swatch(uniform_swatch, hex_color)

        _refresh_button_states()

    def _pick_bullet_color():
        chosen = colorchooser.askcolor(color=bullet_color_var.get(), title="Choose accent color", parent=root)
        hex_color = chosen[1] if chosen else None
        if hex_color:
            _apply_bullet_color(hex_color)

    tk.Label(bullet_frame, text=f"Accent color (bullets {BULLET_ICON_CHAR}, links, headings...):", anchor="w").grid(row=0, column=0, sticky="w", pady=3)
    bullet_swatch = tk.Button(
        bullet_frame, text=BULLET_ICON_COLOR, width=12,
        bg=BULLET_ICON_COLOR, activebackground=BULLET_ICON_COLOR,
        fg=_readable_text_color(BULLET_ICON_COLOR),
        command=_pick_bullet_color,
    )
    bullet_swatch.grid(row=0, column=1, sticky="w", pady=3)

    # Quick-select row: one swatch per gestora, so switching companies is a
    # single click instead of a trip through the OS color-picker each time.
    preset_row = tk.Frame(bullet_frame)
    preset_row.grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 4))
    tk.Label(preset_row, text="Quick select:", fg=UI_MUTED_TEXT_COLOR).pack(side="left", padx=(0, 6))
    for preset_label, preset_hex in ACCENT_COLOR_PRESETS:
        tk.Button(
            preset_row, text=preset_label, width=len(preset_label) + 2,
            bg=preset_hex, activebackground=preset_hex,
            fg=_readable_text_color(preset_hex),
            command=lambda h=preset_hex: _apply_bullet_color(h),
        ).pack(side="left", padx=2)

    tk.Label(
        bullet_frame,
        text="Applied to bullet lists, links, headings, table borders and header tint,\n"
             "the small print and the webinar button - every color in the email is\n"
             "derived from this one.",
        anchor="w", justify="left", fg=UI_MUTED_TEXT_COLOR,
    ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 4))

    button_frame = tk.Frame(root, pady=14)

    def _confirm():
        sig_sel = sig_listbox.curselection()
        disc_sel = disc_listbox.curselection()
        result["signature"] = sig_keys[sig_sel[0] if sig_sel else 0]
        result["disclaimer"] = disc_keys[disc_sel[0] if disc_sel else 0]

        accent = bullet_color_var.get().strip() or BULLET_ICON_COLOR

        def _button_bg(own_color):
            """The shared color wins when 'same color for all buttons' is
            on; otherwise each button keeps its own."""
            if uniform_color_enabled.get():
                return uniform_color_var.get().strip() or accent
            return (own_color or "").strip() or accent

        # None means "let the builder decide black or white per button" -
        # that's the automatic contrast behaviour, and it's why a dark blue
        # button never ends up with unreadable dark text.
        chosen_text_color = None if auto_text_color.get() else (text_color_var.get().strip() or None)

        if webinar_enabled.get():
            result["webinar"] = {
                "enabled": True,
                "button_text": button_text_var.get().strip(),
                "recipient": recipient_var.get().strip(),
                "subject": subject_var.get().strip(),
                "body": body_text.get("1.0", "end").strip(),
                "button_color": _button_bg(button_color_var.get()),
                "text_color": chosen_text_color,
            }
        else:
            result["webinar"] = None

        # Rows left half-filled are dropped here rather than emitting a
        # button that goes nowhere.
        result["link_buttons"] = [
            {
                "text": row["text_var"].get().strip(),
                "url": row["url_var"].get().strip(),
                "color": _button_bg(row["color_var"].get()),
                "text_color": chosen_text_color,
            }
            for row in link_rows
            if row["text_var"].get().strip() and row["url_var"].get().strip()
        ]

        result["bullet_color"] = bullet_color_var.get().strip() or BULLET_ICON_COLOR

        root.destroy()

    tk.Button(button_frame, text="Continue", width=18, command=_confirm) \
        .pack(side="right", padx=20)

    # Packing order matters here: everything below is packed with
    # side="bottom", innermost-last, so each one claims its own fixed
    # height from the bottom of the window FIRST and is therefore always
    # visible - the "Continue" button was getting clipped off the bottom
    # of the window (only reachable by manually resizing it taller)
    # because columns_frame used to be packed first with expand=True,
    # which let it claim space before the fixed-height sections below it
    # had a chance to reserve theirs. columns_frame is packed last here,
    # with expand=True, so it's the one that shrinks (its listboxes and
    # preview panes already scroll) when the window is too short, instead
    # of pushing the button off-screen.
    # Set the initial enabled/disabled state of every color control now
    # that all of them (and bullet_color_var) exist. Doing this at widget
    # creation time would have been too early.
    _refresh_button_states()
    _refresh_empty_hint()

    button_frame.pack(side="bottom", fill="x")
    bullet_frame.pack(side="bottom", fill="x", padx=20, pady=(14, 0))
    feature_notebook.pack(side="bottom", fill="x", padx=20, pady=(14, 0))
    columns_frame.pack(fill="both", expand=True)

    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()

    return (
        result["signature"], result["disclaimer"], result["webinar"],
        result["link_buttons"], result["bullet_color"],
    )

# --- Signature HTML builder ---
def _resolve_image_path(image_path):
    """Returns the actual path to the image, tolerating case differences
    in the filename (e.g. FIRMA_A.png vs firma_a.png)."""
    if not image_path:
        return None
    if os.path.isfile(image_path):
        return image_path

    directory = os.path.dirname(image_path)
    target_name = os.path.basename(image_path).lower()
    if os.path.isdir(directory):
        for entry in os.listdir(directory):
            if entry.lower() == target_name:
                return os.path.join(directory, entry)
    return None


def _image_src_for_bytes(image_bytes, filename, images_dir):
    """Saves a local copy of image_bytes (for your own reference) and
    uploads it to Mailchimp's File Manager, returning the hosted URL to use
    as the <img> src. This replaces base64 embedding, which Mailchimp
    strips out on send.

    Every converted document now shares one images_dir, so it's possible
    for two different files (e.g. two .docx with the same name in
    different folders) to propose the same filename. If a file already
    exists at that path with different content, this disambiguates with a
    numeric suffix ("_2", "_3", ...) instead of silently overwriting the
    earlier file's local reference copy. If the existing file's content is
    identical, it's reused as-is (no redundant write or upload)."""
    os.makedirs(images_dir, exist_ok=True)

    name, ext = os.path.splitext(filename)
    candidate = filename
    n = 2
    while True:
        candidate_path = os.path.join(images_dir, candidate)
        if not os.path.isfile(candidate_path):
            break
        with open(candidate_path, "rb") as existing:
            if existing.read() == image_bytes:
                break  # identical file already saved here - reuse it
        candidate = f"{name}_{n}{ext}"
        n += 1

    out_path = os.path.join(images_dir, candidate)
    with open(out_path, "wb") as f:
        f.write(image_bytes)

    return upload_image_to_mailchimp(image_bytes, candidate)


def build_signature_html(sig_key, images_dir):
    """Builds the HTML block (image + hyperlinks) for the chosen signature."""
    if sig_key is None or sig_key not in SIGNATURES:
        return ""

    sig_data = SIGNATURES[sig_key]
    resolved_path = _resolve_image_path(sig_data.get("image"))

    img_html = ""
    if resolved_path:
        with open(resolved_path, "rb") as img_file:
            image_bytes = img_file.read()
        filename = os.path.basename(resolved_path)
        img_src = _image_src_for_bytes(image_bytes, filename, images_dir)
        # Size inline as well as in the stylesheet: with <style> stripped, an
        # unsized signature renders at the image file's native width.
        img_html = (
            f'<img src="{img_src}" class="signature-image" '
            f'style="width: 100%; max-width: {SIGNATURE_IMAGE_MAX_WIDTH}; height: auto; display: block;" />'
        )
        print(f"Signature image exported: {os.path.join(images_dir, filename)} -> src=\"{img_src}\"")
    else:
        print(f"Warning: signature image not found at '{sig_data.get('image')}'. Skipping image.")

    if not img_html:
        return ""

    return (
        '<table class="signature-block" cellpadding="0" cellspacing="0"><tr>'
        f'<td class="signature-image-cell">{img_html}</td>'
        '</tr></table>'
    )


def _escape_disclaimer(text):
    # Rejoin words/URLs that got hyphen-split across a line-wrap (e.g. from
    # pasted text), then collapse remaining whitespace/newlines into single
    # spaces so each disclaimer renders as one continuous line.
    joined = re.sub(r"-\s*\n\s*", "-", text)
    single_line = re.sub(r"\s+", " ", joined).strip()
    return (
        single_line.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def build_disclaimer_html(disc_key):
    """Builds the disclaimer HTML: the fixed disclaimer always first, followed
    by the chosen extra disclaimer's text (if any)."""
    lines_html = [f'<div class="disclaimer-line">{_escape_disclaimer(DISCLAIMER_FIXED_TEXT)}</div>']

    if disc_key is not None and disc_key in DISCLAIMERS:
        disc_text = DISCLAIMERS[disc_key].get("text", "")
        if disc_text:
            lines_html.append(f'<div class="disclaimer-line disclaimer-extra">{_escape_disclaimer(disc_text)}</div>')

    return f'<div class="disclaimer-block">{"".join(lines_html)}</div>'


def _build_mailto_href(recipient, subject, body):
    """Builds a mailto: URL, then HTML-escapes the '&' between params so
    it's safe to drop straight into an href="..." attribute.

    The recipient address is deliberately left un-percent-encoded: per
    RFC 6068 a normal email address doesn't need escaping, and encoding
    its '@' into '%40' is what was causing some mail clients (notably
    several desktop mailto handlers) to mis-parse the "to" field and show
    a stray leading '/' instead of the actual address. Subject/body are
    free-form text and still need percent-encoding for spaces, accents,
    line breaks, etc."""
    params = []
    if subject:
        params.append(f"subject={_url_quote(subject)}")
    if body:
        params.append(f"body={_url_quote(body)}")

    href = f"mailto:{recipient}"
    if params:
        href += "?" + "&".join(params)
    return href.replace("&", "&amp;")


def _normalize_link_url(url):
    """Makes a user-typed URL safe to use in href="...".

    A bare "youtube.com/watch?v=..." (no scheme) would otherwise be treated
    as a RELATIVE path by every mail client and resolve against the campaign
    URL, so anything that isn't already absolute - and isn't a mailto:/tel:
    link - gets https:// prepended. The '&' separators are escaped last so
    the result can be dropped straight into an attribute."""
    url = (url or "").strip()
    if not url:
        return ""
    # Pasting from a browser or chat app often brings invisible characters
    # along (zero-width spaces, non-breaking spaces, a trailing newline). They
    # survive into href and produce a URL that looks right but 404s - drop
    # all whitespace-class characters, which are never valid inside a URL.
    url = re.sub(r"[\s\u200b\u200c\u200d\u2060\ufeff]+", "", url)
    if url.startswith("//"):
        url = "https:" + url
    elif not re.match(r"^[a-zA-Z][a-zA-Z0-9+.\-]*:", url):
        url = "https://" + url
    return _attr_escape(url)


def _build_button_html(text, href, background_color, text_color=None):
    """Renders one pill button as a "bulletproof" email button: a one-cell
    table, centered, with EVERY style inline.

    The earlier version was a <div class="email-button-block"><a class=
    "email-button">, relying on the <style> block for centering, padding,
    shape and display. Mailchimp content blocks and several email clients
    drop <style>, and then the button collapsed into a bare colored link at
    the left edge - which is what "not centered" and "links look broken"
    were. With nothing depending on the stylesheet, the button survives:

      * align="center" on the table AND margin:auto centers it everywhere,
        including Outlook, which ignores margin:auto on its own;
      * bgcolor on the <td> keeps the color in Outlook, which drops
        background-color on <a>;
      * the padding sits on the <a> (display:inline-block) so the WHOLE
        pill is clickable, not just the letters;
      * width/border/padding are reset inline because the stylesheet's
        generic table/td rules (100% width, cell borders) would otherwise
        hit this table too.

    href must already be escaped for an attribute (see _normalize_link_url
    and _build_mailto_href). text_color None = black or white by contrast."""
    if not text or not href:
        return ""
    label_color = text_color or _readable_text_color(background_color)
    return (
        '<table role="presentation" class="email-button" align="center" border="0" '
        'cellpadding="0" cellspacing="0" '
        'style="margin: 0 auto; width: auto; border-collapse: separate; border: none;">'
        '<tr>'
        f'<td align="center" bgcolor="{background_color}" '
        f'style="background-color: {background_color}; border-radius: {BUTTON_BORDER_RADIUS}; '
        'border: none; padding: 0; text-align: center;">'
        f'<a href="{href}" target="_blank" rel="noopener" '
        f'style="display: inline-block; padding: {BUTTON_PADDING}; '
        f'font-family: {BODY_FONT_FAMILY}; font-size: {BUTTON_FONT_SIZE}; font-weight: bold; '
        f'line-height: 1.2; color: {label_color}; text-decoration: none; '
        f'border-radius: {BUTTON_BORDER_RADIUS}; letter-spacing: 0.5px;">'
        f'{_html_escape(text)}</a>'
        '</td></tr></table>'
    )


def build_buttons_block_html(*button_html_groups):
    """Wraps every button (webinar + link buttons) in one centered block.

    Each button is its own table; spacing between consecutive buttons comes
    from a spacer row rather than margins, because vertical margins between
    tables are unreliable in email clients (Outlook ignores them)."""
    buttons = [html for group in button_html_groups for html in _split_buttons(group)]
    if not buttons:
        return ""
    spacer = f'<div style="height: {BUTTON_STACK_GAP}; line-height: {BUTTON_STACK_GAP}; font-size: 1px;">&nbsp;</div>'
    return (
        f'<div class="email-buttons" style="margin: {BUTTON_BLOCK_MARGIN}; text-align: center;">'
        + spacer.join(buttons)
        + "</div>"
    )


def _split_buttons(html):
    """The builders return their buttons concatenated; split them back into
    individual tables so build_buttons_block_html can space them evenly."""
    if not html:
        return []
    marker = '<table role="presentation" class="email-button"'
    return [marker + part for part in html.split(marker) if part]


def build_webinar_html(webinar_data, accent_color):
    """Builds the optional webinar "confirm attendance" button: a pill-style
    mailto link that opens the user's email client with the recipient,
    subject, and body pre-filled. webinar_data is either None (feature
    toggled off) or the dict returned by select_email_options() containing
    "enabled", "button_text", "recipient", "subject", "body", "button_color"
    and "text_color".

    button_color (a "#rrggbb" string picked in the selector's color picker)
    overrides the accent color for this run; accent_color is the fallback,
    since no fixed button color exists any more. text_color is None when the
    dialog's "automatic" text color is left on, in which case the label is
    painted black or white depending on the background."""
    if not webinar_data or not webinar_data.get("enabled"):
        return ""

    recipient = (webinar_data.get("recipient") or "").strip()
    if not recipient:
        print("Warning: webinar button enabled but no recipient email was set - skipping button.")
        return ""

    button_text = webinar_data.get("button_text") or "CONFIRM ATTENDANCE"
    href = _build_mailto_href(recipient, webinar_data.get("subject", ""), webinar_data.get("body", ""))
    background = webinar_data.get("button_color") or accent_color

    return _build_button_html(button_text, href, background, webinar_data.get("text_color"))


def build_link_buttons_html(link_buttons, accent_color):
    """Builds the custom link buttons (YouTube video, landing page, ...) in
    the same pill format as the webinar button.

    link_buttons is the list returned by select_email_options(): each entry
    is {"text", "url", "color", "text_color"}. It's independent of the
    webinar button, so any combination of the two features is possible.
    Rows missing a label or a URL are skipped with a warning rather than
    emitting a dead button."""
    if not link_buttons:
        return ""

    parts = []
    for button in link_buttons:
        text = (button.get("text") or "").strip()
        href = _normalize_link_url(button.get("url"))
        if not text or not href:
            print(f"Warning: skipping incomplete link button (text={text!r}, url={button.get('url')!r}).")
            continue
        background = button.get("color") or accent_color
        parts.append(_build_button_html(text, href, background, button.get("text_color")))

    return "".join(parts)


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _load_theme_colors(doc_part):
    """Reads the document's theme (theme1.xml) and returns a dict mapping
    Word's theme-color names (e.g. 'accent1', 'hyperlink') to hex RGB
    strings. Needed because text colored via a theme color (very common in
    table styles/header rows) has no literal RGB value in the run itself -
    without this, that color was silently dropped."""
    try:
        theme_part = next(
            rel.target_part for rel in doc_part.rels.values()
            if rel.reltype.endswith("/theme")
        )
    except StopIteration:
        return {}

    a_ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    root = etree.fromstring(theme_part.blob)
    scheme = root.find(".//a:clrScheme", a_ns)
    if scheme is None:
        return {}

    slots = {}
    for child in scheme:
        slot_name = etree.QName(child).localname  # dk1, lt1, accent1, ...
        srgb = child.find("a:srgbClr", a_ns)
        sys_clr = child.find("a:sysClr", a_ns)
        if srgb is not None:
            slots[slot_name] = srgb.get("val")
        elif sys_clr is not None:
            slots[slot_name] = sys_clr.get("lastClr")

    # Word's logical theme-color names map onto the theme's color slots
    name_map = {
        "dark1": "dk1", "text1": "dk1",
        "light1": "lt1", "background1": "lt1",
        "dark2": "dk2", "text2": "dk2",
        "light2": "lt2", "background2": "lt2",
        "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
        "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
        "hyperlink": "hlink", "followedHyperlink": "folHlink",
    }
    return {name: slots[slot] for name, slot in name_map.items() if slot in slots}


def _load_doc_default_size_pt(doc):
    """Reads the document's docDefaults (the <w:docDefaults> block in
    styles.xml) for its default run font size, in points.

    Most of the body text in a typical Word document has no font-size set
    directly on the run at all - it just relies on this document-wide
    fallback (commonly ~9-10pt for a dense factsheet). Word always applies
    that fallback when rendering, but the old HTML output had no
    equivalent: any run without an explicit size got no inline
    font-size, so the browser fell back to ITS OWN default (usually
    16px/12pt) - noticeably bigger than the Word document's real default.
    That's what made small titles and explicitly-sized text look
    disproportionately small next to (accidentally oversized) body text
    in the HTML, even though the title/small text itself was rendering at
    the correct size. Returns None if no default size is declared."""
    try:
        rpr_default = doc.styles.element.find(
            f"{{{W_NS}}}docDefaults/{{{W_NS}}}rPrDefault/{{{W_NS}}}rPr"
        )
        if rpr_default is None:
            return None
        sz = rpr_default.find(f"{{{W_NS}}}sz")
        if sz is None:
            return None
        val = sz.get(f"{{{W_NS}}}val")
        return int(val) / 2 if val else None  # half-points -> points
    except Exception:
        return None


def _load_numbering_formats(doc):
    """Reads the document's numbering part (word/numbering.xml) and returns
    a {numId: {ilvl: numFmt}} mapping, e.g. {"3": {"0": "bullet"}}.

    Word list paragraphs (bullets AND numbered lists) don't store the
    bullet/number glyph as text at all - they just reference a numId/ilvl
    via w:numPr, and Word generates the marker at render time from this
    numbering definition. Without reading this part, there is no way to
    tell a bullet paragraph apart from a plain paragraph, which is why
    bullets were silently rendered as marker-less <p> tags before."""
    try:
        numbering_part = doc.part.numbering_part
    except Exception:
        return {}
    if numbering_part is None:
        return {}
    root = numbering_part.element

    # abstractNumId -> {ilvl: numFmt}
    abstract_fmts = {}
    for abstract_num in root.findall(f"{{{W_NS}}}abstractNum"):
        abstract_id = abstract_num.get(f"{{{W_NS}}}abstractNumId")
        lvl_fmts = {}
        for lvl in abstract_num.findall(f"{{{W_NS}}}lvl"):
            ilvl = lvl.get(f"{{{W_NS}}}ilvl")
            numfmt_el = lvl.find(f"{{{W_NS}}}numFmt")
            fmt = numfmt_el.get(f"{{{W_NS}}}val") if numfmt_el is not None else "bullet"
            lvl_fmts[ilvl] = fmt
        abstract_fmts[abstract_id] = lvl_fmts

    # numId -> abstractNumId (a "num" can override some levels, but for our
    # purposes mapping straight to its abstractNum's formats is enough)
    result = {}
    for num in root.findall(f"{{{W_NS}}}num"):
        num_id = num.get(f"{{{W_NS}}}numId")
        abstract_ref = num.find(f"{{{W_NS}}}abstractNumId")
        if abstract_ref is not None:
            abstract_id = abstract_ref.get(f"{{{W_NS}}}val")
            result[num_id] = abstract_fmts.get(abstract_id, {})
    return result


def _paragraph_list_info(p, numbering_formats):
    """Returns (num_id, ilvl, fmt) if this paragraph is part of a Word list
    (bullet or numbered), or None otherwise. fmt is the raw w:numFmt value:
    'bullet', 'decimal', 'lowerLetter', 'lowerRoman', etc.

    Checks direct paragraph formatting first (w:pPr/w:numPr on the
    paragraph itself), then falls back to the paragraph's style (list
    formatting is often defined once on a "List Paragraph"-type style
    rather than repeated on every paragraph)."""
    numPr = None
    pPr = p._p.pPr
    if pPr is not None:
        numPr = pPr.find(f"{{{W_NS}}}numPr")
    if numPr is None and p.style is not None:
        style_pPr = p.style.element.find(f"{{{W_NS}}}pPr")
        if style_pPr is not None:
            numPr = style_pPr.find(f"{{{W_NS}}}numPr")
    if numPr is None:
        return None

    numid_el = numPr.find(f"{{{W_NS}}}numId")
    if numid_el is None:
        return None
    num_id = numid_el.get(f"{{{W_NS}}}val")
    if not num_id or num_id == "0":
        return None  # numId 0 is Word's explicit "remove numbering" override

    ilvl_el = numPr.find(f"{{{W_NS}}}ilvl")
    ilvl = int(ilvl_el.get(f"{{{W_NS}}}val")) if ilvl_el is not None else 0

    fmt = numbering_formats.get(num_id, {}).get(str(ilvl), "bullet")
    return num_id, ilvl, fmt


# Word list numFmt values that are conceptually "ordered" -> rendered as
# <ol>, with a CSS list-style-type approximating the original marker style.
# Anything not in this map (bullet, and any unrecognized/custom format) is
# treated as an unordered list.
ORDERED_LIST_CSS_TYPE = {
    "decimal": None,             # None = plain <ol> default (1, 2, 3...)
    "lowerLetter": "lower-alpha",
    "upperLetter": "upper-alpha",
    "lowerRoman": "lower-roman",
    "upperRoman": "upper-roman",
}


def _bullet_marker_html(bullet_color):
    """Inline markup for the colored diamond bullet marker, inserted right
    after a list item's opening <li> tag (see prefix_html in
    process_paragraph). Plain inline-styled <span> - no ::marker/::before -
    so it renders consistently in Outlook desktop as well as web/mobile
    email clients."""
    return (
        f'<span style="color: {bullet_color}; margin-right: {BULLET_ICON_MARGIN_RIGHT}; '
        f'font-weight: normal;">{BULLET_ICON_CHAR}</span>'
    )


class ListRenderer:
    """Tracks currently-open <ul>/<ol> elements while a document's blocks
    are iterated in order, so consecutive Word list paragraphs get grouped
    into real, nested HTML lists (with an actual visible marker) instead of
    each becoming its own disconnected, marker-less <p>.

    Call open_or_adjust(ilvl, fmt) for every list paragraph encountered, and
    close_all() whenever a non-list paragraph, a table, or the end of the
    document is reached - Word lists never wrap around those."""

    def __init__(self):
        self.stack = []  # list of (ilvl, tag)

    def open_or_adjust(self, ilvl, fmt):
        is_ordered = fmt in ORDERED_LIST_CSS_TYPE
        tag = "ol" if is_ordered else "ul"
        html = []

        # Close any open levels deeper than this one, or at this same level
        # but with a mismatched list type (e.g. bullet list switching to a
        # numbered one at the same indent).
        while self.stack and (self.stack[-1][0] > ilvl or
                               (self.stack[-1][0] == ilvl and self.stack[-1][1] != tag)):
            _, close_tag = self.stack.pop()
            html.append(f"</{close_tag}>")

        if not self.stack or self.stack[-1][0] < ilvl:
            css_type = ORDERED_LIST_CSS_TYPE.get(fmt) if is_ordered else None
            if is_ordered:
                style_attr = f' style="list-style-type: {css_type};"' if css_type else ""
            else:
                # Bullet lists get their marker rendered manually as inline
                # text (see BULLET_ICON_CHAR / prefix_html) instead of the
                # browser's native disc, since ::marker/::before styling
                # (needed to color/replace that native marker) isn't
                # supported by Outlook's desktop rendering engine.
                style_attr = ' style="list-style: none;"'
            html.append(f"<{tag}{style_attr}>")
            self.stack.append((ilvl, tag))
        # else: already inside the right list at the right level, nothing to open

        return "".join(html)

    def close_all(self):
        html = "".join(f"</{tag}>" for _, tag in reversed(self.stack))
        self.stack = []
        return html


def _effective_run_size_pt(run):
    """Returns the effective font size (in points) for a run: its own
    direct size if set, otherwise the size inherited from its character
    style (walking the basedOn chain) - the same style-inheritance gap
    that caused the missing-bold bug, but for size. Returns None if
    nothing in that chain sets a size either, in which case the caller
    leaves it to the document-wide default set on <body> (see
    _load_doc_default_size_pt), matching how Word itself falls back to
    docDefaults."""
    if run.font.size is not None:
        return run.font.size.pt

    style = run.style
    seen_ids = set()
    while style is not None and id(style) not in seen_ids:
        seen_ids.add(id(style))
        if style.font.size is not None:
            return style.font.size.pt
        style = getattr(style, "base_style", None)
    return None


def _resolve_color(color_element, theme_colors):
    """Given a <w:color> element, returns its hex value: the literal w:val
    if present, otherwise the resolved theme color it points to."""
    if color_element is None:
        return None
    val = color_element.get(f"{{{W_NS}}}val")
    if val and val.lower() not in ("auto", "none"):
        return val
    theme_name = color_element.get(f"{{{W_NS}}}themeColor")
    if theme_name:
        return theme_colors.get(theme_name)
    return None


def _paragraph_border_css(p):
    """Reads the paragraph's border (w:pBdr) and returns CSS border/padding
    declarations for whichever sides are set. Some documents fake a heading
    look (bold text + a colored rule) via direct paragraph formatting rather
    than a real Word Heading style - without this, that rule was dropped."""
    pPr = p._p.pPr
    if pPr is None:
        return []
    pbdr = pPr.find(f"{{{W_NS}}}pBdr")
    if pbdr is None:
        return []

    declarations = []
    for side in ("top", "bottom", "left", "right"):
        side_el = pbdr.find(f"{{{W_NS}}}{side}")
        if side_el is None:
            continue
        val = side_el.get(f"{{{W_NS}}}val")
        if not val or val in ("nil", "none"):
            continue
        color = side_el.get(f"{{{W_NS}}}color") or "000000"
        if color.lower() == "auto":
            color = "000000"
        sz = side_el.get(f"{{{W_NS}}}sz")
        width_pt = (int(sz) / 8) if sz else 1  # sz is in eighths of a point
        declarations.append(f"border-{side}: {width_pt}pt solid #{color};")
        space = side_el.get(f"{{{W_NS}}}space")
        if space:
            declarations.append(f"padding-{side}: {space}pt;")
    return declarations


def _paragraph_spacing_css(p):
    """Reads the paragraph's own explicit before/after spacing (if set) so
    manually-styled section titles keep the vertical rhythm they had in Word."""
    declarations = []
    pf = p.paragraph_format
    if pf.space_before is not None:
        declarations.append(f"margin-top: {pf.space_before.pt}pt;")
    if pf.space_after is not None:
        declarations.append(f"margin-bottom: {pf.space_after.pt}pt;")
    return declarations


def _paragraph_indent_css(p):
    """Reads the paragraph's own indentation (w:ind in pPr) and returns CSS
    margin declarations for it, converting Word's twentieths-of-a-point
    units to points (1pt = 20 twips).

    Without this, any paragraph indented in Word to sit "under" a bullet
    above it (e.g. the Entradas:/Salidas: lines nested under a "Top 10"
    bullet, with no bullet glyph of their own) collapses back to the left
    margin in the HTML output - the indentation is simply dropped, since
    only borders/spacing/shading were being read from pPr before this.
    hanging (a negative first-line indent, used by real bulleted/numbered
    paragraphs) is intentionally NOT emitted here: that's already handled
    by the list-rendering path via padding-left on <ul>/<ol>, and double-
    applying it would push real list items too far right."""
    declarations = []
    pPr = p._p.pPr
    if pPr is None:
        return declarations
    ind = pPr.find(f"{{{W_NS}}}ind")
    if ind is None:
        return declarations

    left = ind.get(f"{{{W_NS}}}left") or ind.get(f"{{{W_NS}}}start")
    right = ind.get(f"{{{W_NS}}}right") or ind.get(f"{{{W_NS}}}end")
    first_line = ind.get(f"{{{W_NS}}}firstLine")

    if left:
        try:
            declarations.append(f"margin-left: {int(left) / 20}pt;")
        except ValueError:
            pass
    if right:
        try:
            declarations.append(f"margin-right: {int(right) / 20}pt;")
        except ValueError:
            pass
    if first_line:
        try:
            declarations.append(f"text-indent: {int(first_line) / 20}pt;")
        except ValueError:
            pass

    return declarations


def _paragraph_shading_css(p):
    """Reads the paragraph's own background shading (w:shd in pPr) and
    returns a background-color declaration for it. This is how Word lets
    you fake a colored title band directly on a paragraph (no table
    involved) - e.g. white bold text on a solid blue paragraph background.
    The old code only picked up shading (w:shd) on table cells, so a
    paragraph shaded this way lost its background in the HTML and its
    white text became invisible white-on-white. A touch of horizontal
    padding is added so the color band doesn't hug the text edge-to-edge."""
    pPr = p._p.pPr
    if pPr is None:
        return []
    shd = pPr.find(f"{{{W_NS}}}shd")
    if shd is None:
        return []
    fill = shd.get(f"{{{W_NS}}}fill")
    if not fill or fill.lower() in ("auto", "none"):
        return []
    return [f"background-color: #{fill};", "padding: 4px 10px;"]


# --- Word document parsing ---
def iter_block_items(parent):
    """Yields paragraphs and tables in the exact order they appear in the document."""
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        return

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def _effective_run_format(run, attr):
    """Returns the effective value of run.bold / run.italic / run.underline,
    falling back to the run's character style (and that style's basedOn
    chain) when the run has no direct formatting for it.

    Word lets you apply bold/italic/underline via a named character style
    (e.g. a "Strong" style with rStyle="Textoennegrita" carrying <w:b/>)
    instead of literal direct formatting on the run. python-docx's
    run.bold/italic/underline only look at direct formatting and return
    None when it's applied this way - so a run that was, say, bold via
    its character style AND had a direct color override lost its bold
    entirely, since the color makes the run take the "has direct
    formatting, use it as-is" path while bold was never seen at all.
    """
    direct = getattr(run, attr)
    if direct is not None:
        return direct

    style = run.style
    seen_ids = set()
    while style is not None and id(style) not in seen_ids:
        seen_ids.add(id(style))
        value = getattr(style.font, attr, None)
        if value is not None:
            return value
        style = getattr(style, "base_style", None)
    return None


def _table_alt_text(table):
    """Returns the table's alt-text title/description (w:tblCaption or
    w:tblDescription), lowercased and stripped, or "".

    Word exposes these under Table Properties > Alt Text. They're the only
    per-table label that survives a round trip through Word without needing
    a custom table style, which makes them a convenient opt-in marker (see
    BAR_CHART_TABLE_MARKER)."""
    tbl_pr = table._tbl.find(f"{{{W_NS}}}tblPr")
    if tbl_pr is None:
        return ""
    for tag in ("tblCaption", "tblDescription"):
        el = tbl_pr.find(f"{{{W_NS}}}{tag}")
        if el is not None:
            val = el.get(f"{{{W_NS}}}val")
            if val:
                return val.strip().lower()
    return ""


def _is_bar_chart_table(table):
    return BAR_CHART_TABLE_MARKER in _table_alt_text(table)


def _row_cell_width_percents(row):
    """Returns a list of column-width percentages for one table row, read
    from each cell's w:tcW, or None if the widths can't be determined.

    Normal tables don't need this (they're auto-laid-out at width:100%),
    but a bar chart does: its bar is made of many narrow slot cells whose
    relative widths ARE the chart. Without explicit widths the browser
    would give every slot the same size and the bars would all look
    identical."""
    widths = []
    for cell in row.cells:
        tc_pr = cell._tc.find(f"{{{W_NS}}}tcPr")
        if tc_pr is None:
            return None
        tc_w = tc_pr.find(f"{{{W_NS}}}tcW")
        if tc_w is None:
            return None
        try:
            widths.append(float(tc_w.get(f"{{{W_NS}}}w") or 0))
        except (TypeError, ValueError):
            return None
    total = sum(widths)
    if total <= 0:
        return None
    return [w * 100.0 / total for w in widths]


WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
EMU_PER_PX = 9525  # 914400 EMU/inch ÷ 96 px/inch
MAX_IMG_WIDTH_PX = 920  # a hair under the 960px page width (BODY_MAX_WIDTH
                        # minus its 20px padding on each side). Images may use
                        # the full page width; text is kept to TEXT_MAX_WIDTH.

# Any image the .docx declares at LARGE_IMAGE_MIN_WIDTH_PX or wider is treated
# as a "full-bleed" image and rendered at 100% of the page width instead of at
# its own pixel width - a screenshot or chart inserted large in Word stays
# readable without zooming, and spills wider than the text column. Images
# inserted deliberately small (banners, logos, inline figures) keep their own
# size. Raise the threshold if something that should stay compact starts
# getting stretched.
LARGE_IMAGE_MIN_WIDTH_PX = 680

# Every image that ISN'T full-bleed (banners, logos, inline charts) is shown at
# its Word width multiplied by this factor, capped at MAX_IMG_WIDTH_PX. Full-bleed
# images already fill the page, so they can't grow further without widening the
# whole email. Set to 1.0 to go back to exact Word sizes.
IMAGE_SCALE = 1.3


def _inline_image_width_px(blip):
    """Reads the image's actual display width as set in the source .docx
    (wp:extent's cx, in EMU) and converts it to CSS pixels.

    Returns None if no wp:extent is found (falls back to the default CSS
    sizing instead of forcing a width).

    This exists so that two images with deliberately different sizes in
    Word - e.g. a compact header banner vs. a full-width fund screenshot at
    the bottom of the DJE End of Month document - keep that size
    difference in the HTML output. Previously every <img> was capped at a
    single hardcoded max-width regardless of how it was sized in Word, so
    the banner and the large screenshot always rendered at the same width.
    """
    extent = blip.xpath('ancestor::*[local-name()="inline" or local-name()="anchor"][1]'
                         f'/*[local-name()="extent"]')
    if not extent:
        return None
    cx = extent[0].get("cx")
    if not cx:
        return None
    try:
        return round(int(cx) / EMU_PER_PX)
    except ValueError:
        return None


def _renderable_blips(element):
    """Returns the image blips inside element that should actually be shown.

    Word wraps many drawings in <mc:AlternateContent>: an <mc:Choice> with
    the modern version and an <mc:Fallback> with a copy for older readers.
    Both usually carry their own <a:blip>, often pointing at DIFFERENT image
    parts (the fallback may be an older, stale or lower-quality rendition).
    Collecting every blip - which is what this code used to do - emitted the
    same banner twice, or the stale fallback banner, depending on order.
    That was the "two banners / wrong banner" bug.

    Rule: a blip inside an mc:Fallback is dropped whenever the Choice branch
    of the same AlternateContent already supplies an image. A Fallback is
    kept only when it is the sole source of an image."""
    result = []
    for blip in element.xpath('.//*[local-name()="blip"]'):
        fallback = blip.xpath('ancestor::*[local-name()="Fallback"][1]')
        if fallback:
            alt = fallback[0].getparent()
            choice_has_image = alt is not None and alt.xpath(
                './*[local-name()="Choice"]//*[local-name()="blip"]'
            )
            if choice_has_image:
                continue
        result.append(blip)
    return result


def process_run(run, doc_part, images_dir, image_counter, theme_colors, is_heading=False, file_prefix=""):
    """Extracts text, inline formatting, and embedded images from a single run.

    is_heading: True when this run lives inside a heading paragraph (h1-h6).
    Word almost always stores an explicit font size on every run - even
    inside a Heading-styled paragraph - and that per-run size used to be
    written out as an inline <span style="font-size:...">, which wins over
    the h1..h6 CSS rules (inline styles beat element selectors). That's why
    every heading rendered at the same size as body text. For heading runs
    we skip the inline font-size entirely and let the heading's own CSS
    rule (HEADING_FONT_SIZES_PT) control the size instead.

    file_prefix: a sanitized version of the source .docx's filename, used
    to namespace embedded image filenames (e.g. "myfile_body_image_1.png").
    All converted documents now share a single images_dir (see
    IMAGE_EXPORT_DIRNAME), so without this prefix two different .docx files
    converted in the same run would both try to write "body_image_1.png"
    and overwrite each other's local reference copy.
    """
    html_parts = []
    text = run.text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>")

    if text:
        styles = []
        color_element = run._element.find(f".//{{{W_NS}}}color")
        run_color = _resolve_color(color_element, theme_colors)
        is_bold = _effective_run_format(run, "bold")
        if RECOLOR_TITLES_TO_ACCENT:
            if is_heading:
                # Drop the inline color entirely so the h1..h6 CSS rule
                # (accent color) applies - see RECOLOR_TITLES_TO_ACCENT.
                run_color = None
            elif is_bold and run_color and _is_title_theme_color(run_color, theme_colors):
                # Bold + dark blue/slate = a category header faked with
                # direct formatting. Repaint it in the accent color.
                run_color = ACCENT_COLOR.lstrip("#")
        if run_color:
            styles.append(f"color: #{run_color};")
        run_size_pt = _effective_run_size_pt(run)
        if run_size_pt and not is_heading:
            styles.append(f"font-size: {run_size_pt}pt;")

        start_tags, end_tags = "", ""
        if is_bold:
            start_tags += "<b>"
            end_tags = "</b>" + end_tags
        if _effective_run_format(run, "italic"):
            start_tags += "<i>"
            end_tags = "</i>" + end_tags
        if _effective_run_format(run, "underline"):
            start_tags += "<u>"
            end_tags = "</u>" + end_tags

        if styles:
            style_str = " ".join(styles)
            start_tags += f'<span style="{style_str}">'
            end_tags = "</span>" + end_tags

        html_parts.append(f"{start_tags}{text}{end_tags}")

    # Handle Inline Embedded Images
    for blip in _renderable_blips(run._element):
        embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        if embed_id and embed_id in doc_part.related_parts:
            img_part = doc_part.related_parts[embed_id]
            ext = img_part.content_type.split("/")[-1]
            ext = "jpg" if ext == "jpeg" else ext

            # Word stores some pasted content (charts/objects copied from
            # Excel or PowerPoint, in particular) as an EMF/WMF vector
            # metafile rather than a real raster image. Mailchimp's File
            # Manager rejects those outright ("files with the x-emf
            # extension are not allowed"), and there's no raster fallback
            # to fall back to here - so skip embedding it rather than
            # crashing the whole conversion. If a fund's chart goes
            # missing from the output, this is almost always why: re-paste
            # it into the source .docx as a picture/PNG instead of an
            # embedded object, or use "Paste Special > Picture (PNG)".
            if ext.lower() in UNSUPPORTED_IMAGE_EXTENSIONS:
                print(f"Warning: skipping unsupported '.{ext}' embedded image (Mailchimp can't host metafile/vector formats). Re-paste it as a PNG/JPEG in the source document if it needs to appear.")
                continue

            image_counter[0] += 1
            name_prefix = f"{file_prefix}_" if file_prefix else ""
            filename = f"{name_prefix}body_image_{image_counter[0]}.{ext}"
            try:
                img_src = _image_src_for_bytes(img_part.blob, filename, images_dir)
            except RuntimeError as e:
                print(f"Warning: skipping image '{filename}' - {e}")
                continue
            width_px = _inline_image_width_px(blip)
            # Sanity-clamp: a wildly oversized value (e.g. a full-resolution
            # screenshot pasted at 100% without resizing in Word) would blow
            # past the 800px content column - cap it there, but otherwise
            # trust the size the document actually specifies. MAX_IMG_WIDTH_PX
            # is intentionally larger than the old blanket 650px cap so a
            # deliberately large image (like the fund screenshot) can still
            # render bigger than a deliberately compact one (like the banner).
            if width_px:
                if width_px >= LARGE_IMAGE_MIN_WIDTH_PX:
                    # Full-bleed (the factsheet screenshot): fill the column
                    # rather than sticking to the exact pixel width Word
                    # recorded, which is usually a bit narrower.
                    width_attr = f' style="width: 100%; max-width: {MAX_IMG_WIDTH_PX}px;"'
                else:
                    width_px = min(round(width_px * IMAGE_SCALE), MAX_IMG_WIDTH_PX)
                    width_attr = f' style="width: {width_px}px; max-width: 100%;"'
            else:
                width_attr = ""
            html_parts.append(f'<img src="{img_src}"{width_attr} />')

    return "".join(html_parts)

def _wrap_in_text_column(chunks):
    """Wraps every run of consecutive non-image chunks in a single
    <div class="text-column"> container, leaving image-only paragraphs
    (IMAGE_BLOCK_CLASS) outside it.

    Why a container instead of `max-width` on each paragraph: Word emits its
    own indentation as an inline `margin-left`, and an inline style beats any
    stylesheet rule - so `margin-left: auto` never applied to indented
    bullets and they escaped the column, hugging the left edge of the page.
    Inside a container, those inline indents are measured from the column
    instead, which is what they mean in Word anyway. Tables, lists, the
    signature and the disclaimer all sit in the same container, so the whole
    body of the email shares one width and only images spill wider."""
    wrapped = []
    is_open = False
    for chunk in chunks:
        if not chunk or not chunk.strip():
            wrapped.append(chunk)
            continue
        is_image_block = chunk.startswith(f'<p class="{IMAGE_BLOCK_CLASS}"')
        if is_image_block:
            if is_open:
                wrapped.append("</div>")
                is_open = False
            wrapped.append(chunk)
            continue
        if not is_open:
            wrapped.append(f'<div class="{TEXT_COLUMN_CLASS}">')
            is_open = True
        wrapped.append(chunk)
    if is_open:
        wrapped.append("</div>")
    return wrapped


def _attr_escape(value):
    """Escapes a value for use inside a double-quoted HTML attribute.
    _html_escape alone leaves '"' untouched, so a URL containing a quote
    would end the href early and swallow the rest of the tag."""
    return _html_escape(value).replace('"', "&quot;")


_ALIGNMENT_CSS = {
    "center": "center",
    "right": "right",
    "end": "right",
    "both": "justify",
    "distribute": "justify",
}


def _paragraph_alignment_css(p):
    """Returns a text-align declaration for the paragraph's alignment (w:jc),
    read from direct formatting first and then up the paragraph style chain.

    Only center / right / justify are emitted. Left-aligned paragraphs are
    deliberately left to the stylesheet's BODY_TEXT_ALIGN (justify), because
    justified body text was an explicit choice for these emails, and many
    generated documents stamp "left" on every paragraph by default - honouring
    that literally would silently undo it. Centered titles and right-aligned
    figures, which used to be flattened to justify, now keep their layout."""
    jc = None
    pPr = p._p.pPr
    if pPr is not None:
        jc_el = pPr.find(f"{{{W_NS}}}jc")
        if jc_el is not None:
            jc = jc_el.get(f"{{{W_NS}}}val")
    style = p.style
    seen = set()
    while jc is None and style is not None and id(style) not in seen:
        seen.add(id(style))
        found = style._element.xpath('./*[local-name()="pPr"]/*[local-name()="jc"]')
        if found:
            jc = found[0].get(f"{{{W_NS}}}val")
        style = getattr(style, "base_style", None)
    css = _ALIGNMENT_CSS.get((jc or "").lower())
    return [f"text-align: {css};"] if css else []


_HYPERLINK_INSTR_RE = re.compile(r'HYPERLINK\s+(?:"([^"]*)"|(\S+))(.*)', re.IGNORECASE | re.DOTALL)
_HYPERLINK_ANCHOR_RE = re.compile(r'\\l\s+"([^"]*)"')


def _hyperlink_url_from_instr(instr):
    r"""Extracts the target URL from a HYPERLINK field instruction, e.g.
    HYPERLINK "https://example.com" \o "tooltip". Returns None for internal
    bookmark links (\l "anchor" with no URL), which have nothing to open in
    an email."""
    match = _HYPERLINK_INSTR_RE.search(instr or "")
    if not match:
        return None
    url = (match.group(1) or match.group(2) or "").strip()
    if url.startswith("\\"):  # e.g. HYPERLINK \l "anchor": no URL at all
        return None
    anchor = _HYPERLINK_ANCHOR_RE.search(match.group(3) or "")
    if url and anchor:
        url = f"{url}#{anchor.group(1)}"
    return url or None


def _link_html(url, inner_html):
    return f'<a href="{_attr_escape(url)}" target="_blank" rel="noopener">{inner_html}</a>'


# Containers whose children are ordinary paragraph content. Word wraps runs
# in these for tracked insertions (ins, moveTo), content controls (sdt),
# smart tags, custom XML and bidi overrides. The old loop only recognised
# direct w:r and w:hyperlink children, so any text inside one of these was
# silently DROPPED from the email - a whole sentence could vanish.
_TRANSPARENT_INLINE_CONTAINERS = {"ins", "moveTo", "smartTag", "customXml", "bdo", "dir", "sdtContent"}
# Tracked deletions: text the author removed. Must not be shown.
_SKIPPED_INLINE_CONTAINERS = {"del", "moveFrom"}


class _InlineRenderer:
    """Renders a paragraph's inline content to HTML, including the parts the
    old loop skipped:

      * runs nested inside content controls / tracked insertions / smart tags;
      * simple fields (w:fldSimple), e.g. HYPERLINK;
      * complex fields - the fldChar begin / instrText / separate / result /
        end sequence Word uses for most hyperlinks it creates itself. Before,
        their visible text came through but the link target was lost, so the
        text looked like a link in Word and was dead in the email.

    Field state spans sibling elements (a field routinely begins in one run
    and ends several runs later), so it lives on the instance for the whole
    paragraph rather than in any one recursive call."""

    def __init__(self, p, render_run):
        self.p = p
        self.render_run = render_run
        self.field_stack = []  # each: {"instr": [...], "in_result": bool, "html": [...]}

    def _emit(self, html, out):
        if self.field_stack:
            top = self.field_stack[-1]
            if top["in_result"]:
                top["html"].append(html)
            # text in a field's INSTRUCTION phase is code, never content
            return
        out.append(html)

    def _handle_field_run(self, r_el, out):
        """Processes fldChar / instrText in a run. Returns True if the run
        was a field-control run (and so must not be rendered as text)."""
        handled = False
        for child in r_el:
            name = etree.QName(child).localname
            if name == "fldChar":
                handled = True
                kind = child.get(f"{{{W_NS}}}fldCharType")
                if kind == "begin":
                    self.field_stack.append({"instr": [], "in_result": False, "html": []})
                elif kind == "separate" and self.field_stack:
                    self.field_stack[-1]["in_result"] = True
                elif kind == "end" and self.field_stack:
                    field = self.field_stack.pop()
                    self._emit(self._finish_field(field), out)
            elif name == "instrText":
                handled = True
                if self.field_stack and not self.field_stack[-1]["in_result"]:
                    self.field_stack[-1]["instr"].append(child.text or "")
        return handled

    def _finish_field(self, field):
        inner = "".join(field["html"])
        url = _hyperlink_url_from_instr("".join(field["instr"]))
        return _link_html(url, inner) if (url and inner) else inner

    def render_children(self, parent_el, out):
        for child in parent_el:
            name = etree.QName(child).localname
            if name == "r":
                if self._handle_field_run(child, out):
                    continue
                self._emit(self.render_run(Run(child, self.p)), out)
            elif name == "hyperlink":
                self._render_hyperlink(child, out)
            elif name == "fldSimple":
                inner = []
                self.render_children(child, inner)
                url = _hyperlink_url_from_instr(child.get(f"{{{W_NS}}}instr"))
                html = "".join(inner)
                self._emit(_link_html(url, html) if (url and html) else html, out)
            elif name == "sdt":
                content = child.find(f"{{{W_NS}}}sdtContent")
                if content is not None:
                    self.render_children(content, out)
            elif name in _TRANSPARENT_INLINE_CONTAINERS:
                self.render_children(child, out)
            elif name in _SKIPPED_INLINE_CONTAINERS:
                continue
            # anything else (pPr, bookmarks, proofErr, comments...) carries
            # no visible content

    def _render_hyperlink(self, h_el, out):
        rId = h_el.get(f"{{{R_NS}}}id")
        anchor = h_el.get(f"{{{W_NS}}}anchor")
        url = ""
        rels = self.p.part.rels
        if rId and rId in rels:
            url = rels[rId].target_ref
            if anchor:
                url = f"{url}#{anchor}"
        # An anchor-only hyperlink jumps to a bookmark inside the .docx;
        # there's nothing for it to point at in the email, so it becomes text.
        inner = []
        self.render_children(h_el, inner)
        html = "".join(inner)
        self._emit(_link_html(url, html) if (url and html) else html, out)

    def render(self):
        out = []
        self.render_children(self.p._p, out)
        # A field left open at the end of the paragraph (it spans into the
        # next one) - keep whatever result text it gathered rather than
        # losing it.
        while self.field_stack:
            field = self.field_stack.pop()
            self._emit("".join(field["html"]), out)
        return out


def process_paragraph(p, doc_part, images_dir, image_counter, theme_colors, file_prefix="", force_tag=None, prefix_html=None):
    """Parses a paragraph, determining if it is a heading, text, or contains hyperlinks, and extracts style colors.

    force_tag: when set (e.g. "li"), overrides whatever tag would otherwise
    be used (p, or h1-h6 for a Heading-styled paragraph). Used to render
    Word list paragraphs as real <li> items - see ListRenderer/
    _paragraph_list_info, which detect list paragraphs via w:numPr.

    prefix_html: raw HTML inserted immediately after the opening tag, before
    any of the paragraph's own runs. Used to inject the colored diamond
    bullet marker (see BULLET_ICON_CHAR) as literal inline markup rather
    than relying on ::marker/::before, which Outlook's desktop rendering
    engine ignores."""

    tag = "p"
    style_declarations = []

    if p.style:
        # Determine if it's a heading
        if p.style.name and p.style.name.startswith("Heading"):
            try:
                level = int(p.style.name.split()[-1])
                if 1 <= level <= 6:
                    tag = f"h{level}"
            except ValueError:
                pass

        # Extract color from the paragraph style (direct value or theme color).
        # Skipped for headings when RECOLOR_TITLES_TO_ACCENT is on: Word's
        # built-in Heading styles carry their own blue, and emitting it here as
        # an inline style would beat the h1..h6 accent-color rule.
        if not (RECOLOR_TITLES_TO_ACCENT and tag != "p"):
            color_elements = p.style._element.xpath('.//*[local-name()="color"]')
            if color_elements:
                color_hex = _resolve_color(color_elements[0], theme_colors)
                if color_hex:
                    style_declarations.append(f"color: #{color_hex};")

    # Preserve direct paragraph formatting (border/spacing/shading) - this is
    # how some documents fake a heading look without using a real Word
    # Heading style.
    border_declarations = _paragraph_border_css(p)
    style_declarations.extend(_paragraph_alignment_css(p))
    style_declarations.extend(border_declarations)
    style_declarations.extend(_paragraph_spacing_css(p))
    style_declarations.extend(_paragraph_shading_css(p))
    if force_tag != "li":
        # Real Word list items (numPr -> <li>) get their indentation from
        # the <ul>/<ol> padding-left instead; applying w:ind's margin-left
        # on top of that would double-indent them.
        style_declarations.extend(_paragraph_indent_css(p))

    is_heading = tag != "p"

    if force_tag:
        tag = force_tag
        is_heading = False  # a list item is never a heading, even if the
        # paragraph happened to carry Heading styling by mistake

    inline_style = f' style="{" ".join(style_declarations)}"' if style_declarations else ""
    p_html = [f"<{tag}{inline_style}>"]
    if prefix_html:
        p_html.append(prefix_html)
    renderer = _InlineRenderer(
        p,
        lambda run: process_run(run, doc_part, images_dir, image_counter, theme_colors, is_heading, file_prefix),
    )
    content_html = renderer.render()
    p_html.extend(content_html)

    inner = "".join(content_html)
    visible_text = re.sub(r"<[^>]+>", "", inner).replace("&nbsp;", "").strip()

    # Empty paragraphs carrying only a border are deliberately KEPT: the
    # factsheet prompt builds its accent-coloured separator between funds
    # exactly that way. Whether a separator should exist at a given point is
    # a content decision, so it's controlled in the prompt, not here.

    p_html.append(f"</{tag}>")
    html = "".join(p_html)

    # A paragraph whose only content is an image is not "text": it must be
    # allowed to use the full page width rather than the narrower text column
    # (see TEXT_MAX_WIDTH). Tag it so the stylesheet can exempt it.
    if "<img" in inner and not visible_text:
        html = html.replace(f"<{tag}", f'<{tag} class="{IMAGE_BLOCK_CLASS}"', 1)

    return html

def _render_block_sequence(blocks, ctx, out):
    """Renders paragraphs (with list detection) and nested tables in order.
    Shared by table cells so a cell gets exactly the same treatment as the
    document body - including tables nested inside it, which the old cell
    loop (cell.paragraphs only) silently dropped."""
    list_renderer = ListRenderer()
    for block in blocks:
        if isinstance(block, Paragraph):
            info = _paragraph_list_info(block, ctx["numbering_formats"])
            if info:
                _num_id, ilvl, fmt = info
                out.append(list_renderer.open_or_adjust(ilvl, fmt))
                marker = None if fmt in ORDERED_LIST_CSS_TYPE else _bullet_marker_html(ctx["bullet_color"])
                out.append(process_paragraph(
                    block, ctx["doc_part"], ctx["images_dir"], ctx["image_counter"],
                    ctx["theme_colors"], ctx["file_prefix"], force_tag="li", prefix_html=marker,
                ))
            else:
                out.append(list_renderer.close_all())
                out.append(process_paragraph(
                    block, ctx["doc_part"], ctx["images_dir"], ctx["image_counter"],
                    ctx["theme_colors"], ctx["file_prefix"],
                ))
        elif isinstance(block, Table):
            out.append(list_renderer.close_all())
            out.extend(_render_table(block, ctx))
    out.append(list_renderer.close_all())


def _cell_style_declarations(tc):
    """Background and vertical alignment from the cell's OWN properties
    (w:tcPr). The old lookup searched the whole cell (.//shd), so shading on
    a paragraph inside the cell could be mistaken for the cell's fill."""
    styles = []
    tcPr = tc.tcPr
    if tcPr is None:
        return styles
    shd = tcPr.find(f"{{{W_NS}}}shd")
    if shd is not None:
        fill = shd.get(f"{{{W_NS}}}fill")
        if fill and fill.lower() not in ("auto", "none"):
            styles.append(f"background-color: #{fill};")
    v_align = tcPr.find(f"{{{W_NS}}}vAlign")
    if v_align is not None:
        css = {"top": "top", "center": "middle", "bottom": "bottom"}.get(v_align.get(f"{{{W_NS}}}val"))
        if css:
            styles.append(f"vertical-align: {css};")
    return styles


def _table_cell_layout(table):
    """Maps the table's real <w:tc> elements onto grid columns.

    python-docx's row.cells returns one entry PER GRID COLUMN, repeating a
    merged cell for every column it spans - so a gestora band spanning six
    columns rendered its text six times across the row. This works from the
    actual cells instead, recording each one's starting column, its colspan
    (w:gridSpan) and its vertical-merge state (w:vMerge), and then computes
    rowspans by looking down for "continue" cells in the same column."""
    rows = []
    for tr in table._tbl.tr_lst:
        col = 0
        trPr = tr.trPr
        if trPr is not None:
            before = trPr.find(f"{{{W_NS}}}gridBefore")
            if before is not None:
                try:
                    col += int(before.get(f"{{{W_NS}}}val", "0"))
                except ValueError:
                    pass
        row = []
        for tc in tr.tc_lst:
            span = tc.grid_span or 1
            row.append({"tc": tc, "col": col, "colspan": span, "vmerge": tc.vMerge, "rowspan": 1})
            col += span
        rows.append(row)

    for r_index, row in enumerate(rows):
        for cell in row:
            if cell["vmerge"] != "restart":
                continue
            span = 1
            for below in rows[r_index + 1:]:
                match = next((c for c in below if c["col"] == cell["col"]), None)
                if match is not None and match["vmerge"] == "continue":
                    span += 1
                else:
                    break
            cell["rowspan"] = span
    return rows


def _render_table(table, ctx):
    """Renders one Word table to HTML chunks. Bar-chart tables keep their
    dedicated one-table-per-row layout (their column widths ARE the data);
    everything else gets real colspan/rowspan and nested-table support."""
    out = []
    if _is_bar_chart_table(table):
        # Wrapped so the chart as a whole gets the same breathing room as
        # any other table, while its rows stay flush against each other.
        out.append(f'<div class="bar-chart-block" style="margin: {TABLE_BLOCK_MARGIN};">')
        for row in table.rows:
            out.append('<table class="bar-chart-row">')
            out.append("<tr>")
            width_percents = _row_cell_width_percents(row)
            for cell_index, cell in enumerate(row.cells):
                cell_styles = _cell_style_declarations(cell._tc)
                if width_percents and cell_index < len(width_percents):
                    cell_styles.append(f"width: {width_percents[cell_index]:.4f}%;")
                style_attr = f' style="{" ".join(cell_styles)}"' if cell_styles else ""
                out.append(f"<td{style_attr}>")
                _render_block_sequence(iter_block_items(cell), ctx, out)
                out.append("</td>")
            out.append("</tr>")
            out.append("</table>")
        out.append("</div>")
        return out

    out.append("<table>")
    for row in _table_cell_layout(table):
        out.append("<tr>")
        for cell in row:
            if cell["vmerge"] == "continue":
                continue  # covered by the rowspan of the cell above
            attrs = ""
            if cell["colspan"] > 1:
                attrs += f' colspan="{cell["colspan"]}"'
            if cell["rowspan"] > 1:
                attrs += f' rowspan="{cell["rowspan"]}"'
            cell_styles = _cell_style_declarations(cell["tc"])
            if cell_styles:
                attrs += f' style="{" ".join(cell_styles)}"'
            out.append(f"<td{attrs}>")
            _render_block_sequence(iter_block_items(_Cell(cell["tc"], table)), ctx, out)
            out.append("</td>")
        out.append("</tr>")
    out.append("</table>")
    return out


# ---------------------------------------------------------------------
# CSS inlining
# ---------------------------------------------------------------------
# Mailchimp (and several other senders) drop a <style> block when the HTML
# is pasted into a content block: the campaign PREVIEW renders the file
# whole and looks perfect, then the delivered email arrives with no table
# borders, no header tints and text running edge to edge. Outlook makes it
# worse by ignoring parts of <style> even when it survives.
#
# So the stylesheet is treated as a source, not as the delivery mechanism:
# every rule is copied onto the elements it matches as an inline style
# attribute, which no client strips. The <style> block is still emitted (it
# costs nothing and keeps the standalone .html file readable), but nothing
# depends on it any more.
#
# This is a deliberately small implementation covering exactly the selector
# shapes this script emits - tag, .class, tag.class, descendant and child
# combinators, comma-separated lists. It's not a general CSS engine, and it
# doesn't need to be: it only ever sees the stylesheet built below.

_CSS_COMMENT_RE = re.compile(r"/\*.*?\*/", re.DOTALL)
_CSS_RULE_RE = re.compile(r"([^{}]+)\{([^{}]*)\}", re.DOTALL)
_SIMPLE_SELECTOR_RE = re.compile(r"^([a-zA-Z][\w-]*)?((?:\.[\w-]+)*)$")
# Properties that only make sense for on-screen layout, or that an email
# client would reject on an element: skipped rather than inlined.
_NON_INLINABLE_PROPERTIES = {"cursor"}


def _parse_css_rules(css):
    """Returns [(selector, declarations, source_order), ...] with comments
    and at-rules removed."""
    css = _CSS_COMMENT_RE.sub("", css)
    rules = []
    for order, (selector_group, declarations) in enumerate(_CSS_RULE_RE.findall(css)):
        declarations = declarations.strip()
        if not declarations:
            continue
        for selector in selector_group.split(","):
            selector = selector.strip()
            if selector and not selector.startswith("@"):
                rules.append((selector, declarations, order))
    return rules


def _selector_to_xpath(selector):
    """Converts a simple CSS selector to XPath, or returns None for anything
    outside the supported subset (pseudo-classes, attribute selectors, ...),
    which is then left to the <style> block alone."""
    tokens = selector.replace(">", " > ").split()
    xpath = "."
    child_combinator = False
    for token in tokens:
        if token == ">":
            child_combinator = True
            continue
        match = _SIMPLE_SELECTOR_RE.match(token)
        if not match:
            return None
        tag = match.group(1) or "*"
        predicates = "".join(
            f"[contains(concat(' ', normalize-space(@class), ' '), ' {cls} ')]"
            for cls in match.group(2).split(".") if cls
        )
        xpath += ("/" if child_combinator else "//") + tag + predicates
        child_combinator = False
    return xpath


def _selector_specificity(selector):
    """(classes, tags) - enough ordering for this stylesheet, which has no
    IDs or attribute selectors."""
    return (selector.count("."), len(re.findall(r"(?:^|[\s>])([a-zA-Z][\w-]*)", selector)))


def _split_declarations(declarations):
    """Splits a declaration block into an ordered {property: value} dict,
    dropping !important (meaningless once inline) and properties that don't
    belong in an email."""
    parsed = {}
    for declaration in declarations.split(";"):
        if ":" not in declaration:
            continue
        prop, _, value = declaration.partition(":")
        prop = prop.strip().lower()
        value = value.replace("!important", "").strip()
        if prop and value and prop not in _NON_INLINABLE_PROPERTIES:
            parsed[prop] = value
    return parsed


def _inline_stylesheet(html):
    """Copies the document's own <style> rules onto matching elements as
    inline styles and returns the rewritten HTML.

    Precedence follows CSS: rules are applied least-specific first, and any
    style already written inline by the converter (a cell's shading, a run's
    colour, a button's background) is applied LAST so it always wins - those
    carry the actual content, while the stylesheet only carries defaults."""
    try:
        from lxml import html as lxml_html
    except ImportError:  # pragma: no cover - lxml is a hard dependency
        return html

    tree = lxml_html.fromstring(html)
    style_elements = tree.xpath("//style")
    if not style_elements:
        return html

    rules = []
    for element in style_elements:
        rules.extend(_parse_css_rules(element.text or ""))
    rules.sort(key=lambda rule: (_selector_specificity(rule[0]), rule[2]))

    pending = {}  # element -> merged declarations from the stylesheet
    for selector, declarations, _order in rules:
        xpath = _selector_to_xpath(selector)
        if xpath is None:
            continue
        try:
            matches = tree.xpath(xpath)
        except Exception:
            continue
        parsed = _split_declarations(declarations)
        for element in matches:
            pending.setdefault(element, {}).update(parsed)

    # Walk the tree root-first, tracking what each element would inherit, and
    # drop any declaration that merely repeats an inherited value. On a long
    # factsheet this removes thousands of redundant "text-align:justify" and
    # "color:#333333" copies - which matters only because of Gmail's ~102 kB
    # clipping threshold, but at these document sizes it genuinely decides
    # whether the email arrives whole.
    #
    # font-family and font-size are deliberately NOT pruned: Outlook resets
    # the font inside <table>, so those have to be restated on table cells
    # even though CSS says they would inherit.
    inheritable = ("color", "text-align", "line-height", "font-style",
                   "font-weight", "letter-spacing")

    def _apply(element, inherited):
        declarations = pending.get(element, {})
        own = _split_declarations(element.get("style", ""))
        declarations.update(own)  # the element's own inline style wins

        for prop in inheritable:
            if prop in declarations and declarations[prop] == inherited.get(prop):
                del declarations[prop]

        if declarations:
            element.set("style", "".join(f"{k}:{v};" for k, v in declarations.items()))
        elif element.get("style") is not None:
            del element.attrib["style"]

        child_inherited = dict(inherited)
        for prop in inheritable:
            if prop in declarations:
                child_inherited[prop] = declarations[prop]
        # Presentational attributes set alignment too, and they are what the
        # children actually inherit. Missing this once cost the whole email
        # its justification: the button/container wrapper carries
        # align="center", so pruning a "redundant" text-align:justify below
        # it let the centre alignment take over instead.
        align = element.get("align")
        if align:
            child_inherited["text-align"] = align.lower()
        for child in element:
            _apply(child, child_inherited)

    _apply(tree, {})

    # The <style> block is now redundant - every rule it held is inline on
    # the elements. Dropping it matters because of Gmail's ~102 kB clipping
    # threshold: past it Gmail truncates the message and shows "View entire
    # message", which on a long factsheet would cut the email mid-table.
    for element in style_elements:
        element.getparent().remove(element)

    return lxml_html.tostring(tree, encoding="unicode", doctype="<!DOCTYPE html>")


def _email_container_open(font_family, font_color, font_size_pt, line_height,
                          max_width, padding, text_align):
    """Opens the wrapper that replaces <body> for layout purposes.

    Pasting into a Mailchimp content block discards <html>/<body> entirely,
    so any styling on <body> - the font, the 960px column, the centering -
    goes with them. This reproduces all of it on elements that survive: an
    outer 100%-width table (Outlook centres via align/table, not margin:auto)
    around an inner div carrying the width cap and the typography."""
    return (
        '<table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" '
        'style="width: 100%; border-collapse: collapse; border: none;">'
        '<tr><td align="center" style="border: none; padding: 0;">'
        f'<div class="email-container" style="max-width: {max_width}; margin: 0 auto; '
        f'padding: {padding}; font-family: {font_family}; color: {font_color}; '
        f'font-size: {font_size_pt}pt; line-height: {line_height}; text-align: {text_align};">'
    )


_EMAIL_CONTAINER_CLOSE = "</div></td></tr></table>"


# --- Execution pipeline ---
# --- Progress window shown while converting ---
class ProgressWindow:
    """A small always-on-top window with a determinate progress bar and a
    status line. Call .step(message) once per unit of work (there are 3
    steps per file: parsing, rendering/uploading images, writing the HTML),
    and .close() when everything is done."""

    def __init__(self, total_steps, title="Converting..."):
        self.root = tk.Tk()
        self.root.title(title)
        self.root.call('wm', 'attributes', '.', '-topmost', True)
        self.root.resizable(False, False)
        self.root.geometry("460x130")

        self.status_label = tk.Label(self.root, text="Starting...", pady=8, wraplength=420, justify="left")
        self.status_label.pack(fill="x", padx=20)

        self.bar = ttk.Progressbar(
            self.root, orient="horizontal", length=420,
            mode="determinate", maximum=max(total_steps, 1),
        )
        self.bar.pack(pady=8, padx=20)

        self.count_label = tk.Label(self.root, text=f"0 / {total_steps}")
        self.count_label.pack()

        self.total_steps = total_steps
        self.current = 0
        self.root.update()

    def step(self, message):
        self.current += 1
        self.bar["value"] = self.current
        self.status_label.config(text=message)
        self.count_label.config(text=f"{self.current} / {self.total_steps}")
        self.root.update_idletasks()
        self.root.update()

    def close(self):
        self.root.destroy()


# --- HTML source viewer / copy-to-clipboard window ---
def show_html_output_window(html_by_file, accent_color=BULLET_ICON_COLOR):
    """Shown once all files are converted. Lists every converted filename;
    selecting one displays its full HTML source (read-only, still
    selectable/scrollable) with a button to copy that source straight to
    the clipboard - handy for pasting into Mailchimp's code editor without
    having to open the .html file separately.

    html_by_file: dict mapping display filename -> full HTML source string,
    in the order the files were converted."""
    if not html_by_file:
        return

    root = tk.Tk()
    root.title("Generated HTML")
    root.call('wm', 'attributes', '.', '-topmost', True)
    root.geometry("980x680")
    root.minsize(760, 520)

    tk.Label(
        root,
        text="Conversion complete. Select a file to view its HTML source, "
             "then copy it to the clipboard.",
        padx=20, pady=12, justify="left", anchor="w",
        font=("TkDefaultFont", 10, "bold"),
    ).pack(fill="x")

    main_frame = tk.Frame(root, padx=20)
    main_frame.pack(fill="both", expand=True)
    main_frame.columnconfigure(0, weight=0)
    main_frame.columnconfigure(1, weight=1)
    main_frame.rowconfigure(0, weight=1)

    filenames = list(html_by_file.keys())

    list_frame = tk.Frame(main_frame)
    list_frame.grid(row=0, column=0, sticky="ns", padx=(0, 12))
    tk.Label(list_frame, text="Converted files", font=("TkDefaultFont", 10, "bold"), anchor="w") \
        .pack(fill="x")

    list_scrollbar = tk.Scrollbar(list_frame, orient="vertical")
    file_listbox = tk.Listbox(
        list_frame, exportselection=False, activestyle="dotbox",
        width=34, height=24, yscrollcommand=list_scrollbar.set,
    )
    list_scrollbar.config(command=file_listbox.yview)
    file_listbox.pack(side="left", fill="y", expand=True)
    list_scrollbar.pack(side="left", fill="y")

    for name in filenames:
        file_listbox.insert("end", name)

    text_frame = tk.Frame(main_frame)
    text_frame.grid(row=0, column=1, sticky="nsew")
    text_frame.rowconfigure(0, weight=1)
    text_frame.columnconfigure(0, weight=1)

    html_text = scrolledtext.ScrolledText(text_frame, wrap="none", font=("Courier New", 9))
    html_text.grid(row=0, column=0, sticky="nsew")

    def _show_file(idx):
        name = filenames[idx]
        html_text.config(state="normal")
        html_text.delete("1.0", "end")
        html_text.insert("1.0", html_by_file[name])
        html_text.config(state="disabled")

    def _on_select(_event=None):
        sel = file_listbox.curselection()
        if sel:
            _show_file(sel[0])

    file_listbox.bind("<<ListboxSelect>>", _on_select)
    file_listbox.selection_set(0)
    _show_file(0)

    status_var = tk.StringVar(value="")

    def _copy():
        sel = file_listbox.curselection()
        idx = sel[0] if sel else 0
        content = html_by_file[filenames[idx]]
        root.clipboard_clear()
        root.clipboard_append(content)
        root.update()  # push the clipboard content now, before the window might close
        status_var.set(f"Copied {filenames[idx]} to clipboard!")
        root.after(2500, lambda: status_var.set(""))

    button_bar = tk.Frame(root, pady=14)
    button_bar.pack(fill="x")
    tk.Button(button_bar, text="Copy to Clipboard", width=20, command=_copy) \
        .pack(side="left", padx=20)
    # The confirmation message uses the run's accent color rather than a
    # fixed green, so nothing outside the accent palette is hardcoded.
    tk.Label(button_bar, textvariable=status_var, fg=accent_color).pack(side="left")
    tk.Button(button_bar, text="Close", width=12, command=root.destroy) \
        .pack(side="right", padx=20)

    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()


if __name__ == "__main__":
    selected_files = select_docx_files()

    if not selected_files:
        print("No files were selected. Exiting script.")
        exit()

    (selected_signature_key, selected_disclaimer_key, selected_webinar,
     selected_link_buttons, selected_bullet_color) = select_email_options()
    # Publish the chosen accent color so process_run can recolor titles with it.
    ACCENT_COLOR = selected_bullet_color
    if selected_signature_key:
        print(f"Selected signature: {SIGNATURES[selected_signature_key].get('label', selected_signature_key)}")
    else:
        print("No signature selected. Files will be converted without a signature.")

    if selected_disclaimer_key:
        print(f"Selected disclaimer: {DISCLAIMERS[selected_disclaimer_key].get('label', selected_disclaimer_key)}")
    else:
        print("No additional disclaimer selected. Only the fixed disclaimer will be added.")

    if selected_webinar:
        print(f"Webinar button enabled: \"{selected_webinar['button_text']}\" -> {selected_webinar['recipient']}")
    else:
        print("No webinar button. Files will be converted without one.")

    if selected_link_buttons:
        for button in selected_link_buttons:
            print(f"Link button: \"{button['text']}\" -> {button['url']} ({button['color']})")
    else:
        print("No custom link buttons.")

    print(f"Bullet icon color: {selected_bullet_color}")

    progress = ProgressWindow(len(selected_files) * 3, title="Converting documents...")
    html_by_file = {}

    for docx_path in selected_files:
        filename = os.path.basename(docx_path)
        base_name, _ = os.path.splitext(docx_path)
        html_path = f"{base_name}.html"
        images_dir = os.path.join(SCRIPT_DIR, IMAGE_EXPORT_DIRNAME)
        # All documents share images_dir now, so namespace each one's body
        # images with its own filename to avoid collisions between files.
        file_prefix = re.sub(r"[^A-Za-z0-9_-]+", "_", os.path.basename(base_name)).strip("_")

        print(f"Converting: {filename}...")
        progress.step(f"Parsing {filename}...")

        doc = Document(docx_path)
        theme_colors = _load_theme_colors(doc.part)
        doc_default_size_pt = _load_doc_default_size_pt(doc)
        numbering_formats = _load_numbering_formats(doc)
        html_body = []
        image_counter = [0]  # mutable int so nested calls can increment it
        list_renderer = ListRenderer()
        # Everything table rendering needs, bundled so _render_table can
        # recurse into nested tables without a ten-argument signature.
        table_ctx = {
            "doc_part": doc.part,
            "images_dir": images_dir,
            "image_counter": image_counter,
            "theme_colors": theme_colors,
            "file_prefix": file_prefix,
            "numbering_formats": numbering_formats,
            "bullet_color": selected_bullet_color,
        }

        for block in iter_block_items(doc):
            if isinstance(block, Paragraph):
                list_info = _paragraph_list_info(block, numbering_formats)
                if list_info:
                    num_id, ilvl, fmt = list_info
                    html_body.append(list_renderer.open_or_adjust(ilvl, fmt))
                    is_ordered = fmt in ORDERED_LIST_CSS_TYPE
                    marker_html = None if is_ordered else _bullet_marker_html(selected_bullet_color)
                    html_body.append(process_paragraph(block, doc.part, images_dir, image_counter, theme_colors, file_prefix, force_tag="li", prefix_html=marker_html))
                else:
                    html_body.append(list_renderer.close_all())
                    html_body.append(process_paragraph(block, doc.part, images_dir, image_counter, theme_colors, file_prefix))
            elif isinstance(block, Table):
                html_body.append(list_renderer.close_all())
                html_body.extend(_render_table(block, table_ctx))

        html_body.append(list_renderer.close_all())

        progress.step(f"Rendering & uploading images for {filename}...")

        # Append the webinar button and any custom link buttons (each
        # independent of the other), then the signature, then the
        # disclaimer, at the end of the document body.
        webinar_html = build_webinar_html(selected_webinar, selected_bullet_color)
        link_buttons_html = build_link_buttons_html(selected_link_buttons, selected_bullet_color)
        html_body.append(build_buttons_block_html(webinar_html, link_buttons_html))

        signature_html = build_signature_html(selected_signature_key, images_dir)
        html_body.append(signature_html)

        disclaimer_html = build_disclaimer_html(selected_disclaimer_key)
        html_body.append(disclaimer_html)

        raw_html = "".join(_wrap_in_text_column(html_body))

        # The chosen accent color (selected_bullet_color) re-skins headings
        # (title text, always bold), links and the table header row to match
        # whichever gestora's factsheet is being converted. Category lines
        # faked with bold dark-blue direct formatting are repainted in the
        # same color inside process_run (RECOLOR_TITLES_TO_ACCENT).
        heading_css = "\n        ".join(
            f"h{level} {{ font-size: {size}pt !important; color: {selected_bullet_color} !important; "
            f"font-weight: {HEADING_FONT_WEIGHT} !important; }}"
            for level, size in HEADING_FONT_SIZES_PT.items()
        )
        # Every hued color in the stylesheet is derived from the accent here
        # (see COLOR POLICY at the top): nothing chromatic is hardcoded.
        table_header_tint = _lighten_hex_color(selected_bullet_color, TABLE_HEADER_TINT)
        table_border_color = _lighten_hex_color(selected_bullet_color, TABLE_BORDER_TINT)
        disclaimer_color = _mute_hex_color(selected_bullet_color, DISCLAIMER_MUTE)

        body_font_size_pt = doc_default_size_pt or BODY_FONT_SIZE_FALLBACK_PT

        container_open = _email_container_open(
            BODY_FONT_FAMILY, BODY_FONT_COLOR, body_font_size_pt,
            BODY_LINE_HEIGHT, BODY_MAX_WIDTH, BODY_PADDING, BODY_TEXT_ALIGN,
        )
        container_close = _EMAIL_CONTAINER_CLOSE

        full_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{os.path.basename(base_name)}</title>
    <style>
        body {{
            font-family: {BODY_FONT_FAMILY};
            color: {BODY_FONT_COLOR};
            font-size: {body_font_size_pt}pt;
            line-height: {BODY_LINE_HEIGHT};
            max-width: {BODY_MAX_WIDTH};
            margin: 0 auto;
            padding: {BODY_PADDING};
            text-align: {BODY_TEXT_ALIGN};
        }}
        /* Everything that is READ lives in one centered column; only
           image-only paragraphs (see LARGE_IMAGE_MIN_WIDTH_PX) sit outside
           it and may use the full page width. The column is a real
           container, so Word's inline margin-left indents are measured from
           its left edge instead of the page's. */
        .{TEXT_COLUMN_CLASS} {{
            max-width: {TEXT_MAX_WIDTH};
            margin: 0 auto;
        }}
        .{TEXT_COLUMN_CLASS} > table {{
            width: 100%;
        }}
        p.{IMAGE_BLOCK_CLASS} {{
            max-width: 100%;
            margin-left: 0;
            margin-right: 0;
        }}
        {heading_css}
        a {{
            color: {selected_bullet_color};
            text-decoration: underline;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: {TABLE_BLOCK_MARGIN};
        }}
        /* A table nested inside a cell shouldn't get the full between-tables
           gap - that spacing is for separating blocks of the email. */
        td table {{
            margin: 6px 0;
        }}
        th, td {{
            border: 1px solid {table_border_color};
            padding: {TABLE_CELL_PADDING};
            text-align: {BODY_TEXT_ALIGN};
            line-height: 1.2;
        }}
        th {{
            background-color: {table_header_tint};
        }}
        table p {{
            margin: 0;
        }}
        /* Native bar-chart rows (see BAR_CHART_TABLE_MARKER): one table per
           bar, borderless and fixed-layout so the per-cell width
           percentages emitted above are honoured exactly. The white
           hairline top/bottom is what separates consecutive bars - padding
           can't, because a cell's background is painted across its padding
           box too. */
        table.bar-chart-row {{
            table-layout: fixed;
            border-collapse: collapse;
            width: 100%;
            margin: 0;
        }}
        table.bar-chart-row td {{
            border-top: 1px solid {BAR_CHART_ROW_GAP_COLOR};
            border-bottom: 1px solid {BAR_CHART_ROW_GAP_COLOR};
            border-left: none;
            border-right: none;
            padding: {BAR_CHART_CELL_PADDING};
            height: {BAR_CHART_ROW_HEIGHT};
            line-height: 1.2;
            text-align: left;
            vertical-align: middle;
            font-size: 8pt;
            white-space: nowrap;
            overflow: hidden;
        }}
        table.bar-chart-row td p {{
            margin: 0;
            padding: {BAR_CHART_LABEL_PADDING};
        }}
        ul, ol {{
            padding-left: {LIST_PADDING_LEFT};
            margin: {LIST_MARGIN};
            text-align: left;
        }}
        li {{
            margin: {LIST_ITEM_MARGIN};
        }}
        td ul, td ol {{
            margin: 4px 0;
        }}
        img {{
            max-width: {IMAGE_MAX_WIDTH};
            height: auto;
            display: block;
            margin: {IMAGE_MARGIN};
        }}
        table img {{
            max-width: {TABLE_IMAGE_MAX_WIDTH};
        }}
        .signature-block {{
            margin-top: {SIGNATURE_BLOCK_MARGIN_TOP};
            padding-top: {SIGNATURE_BLOCK_PADDING_TOP};
            border-collapse: collapse;
        }}
        .signature-block td {{
            border: none;
            padding: 0;
            vertical-align: middle;
        }}
        .signature-block .signature-image {{
            max-width: {SIGNATURE_IMAGE_MAX_WIDTH};
            width: 100%;
            height: auto;
            margin: 0;
            display: block;
        }}
        .signature-links-cell {{
            font-size: {SIGNATURE_FONT_SIZE};
            padding-left: {SIGNATURE_CELL_SPACING};
            border-left: 1px solid {table_border_color};
        }}
        .signature-link {{
            margin: 3px 0;
        }}
        .signature-link a {{
            text-decoration: none;
        }}
        .disclaimer-block {{
            margin-top: {DISCLAIMER_BLOCK_MARGIN_TOP};
            font-size: {DISCLAIMER_FONT_SIZE};
            color: {disclaimer_color};
            line-height: {DISCLAIMER_LINE_HEIGHT};
            text-align: {BODY_TEXT_ALIGN};
        }}
        .disclaimer-line {{
            margin: 4px 0;
        }}
        .disclaimer-extra {{
            margin-top: {DISCLAIMER_GAP};
        }}
        /* Buttons carry ALL their styling inline (see _build_button_html),
           so nothing here is needed to render them. This only neutralises
           the generic table/td rules above for button tables, as a second
           line of defence where a client applies <style> but also ignores
           some inline declarations. */
        table.email-button {{
            width: auto;
            margin: 0 auto;
            border-collapse: separate;
        }}
        table.email-button td {{
            border: none;
            padding: 0;
        }}
        table.email-button a {{
            text-decoration: none;
        }}
    </style>
</head>
<body style="margin: 0; padding: 0; width: 100%;">
    {container_open}{raw_html}{container_close}
</body>
</html>"""

        # Bake the stylesheet onto the elements themselves. Everything above
        # is written for readability; this is what makes the email survive a
        # client that discards <style> (see _inline_stylesheet).
        full_html = _inline_stylesheet(full_html)

        progress.step(f"Writing HTML for {filename}...")

        with open(html_path, "w", encoding="utf-8") as html_file:
            html_file.write(full_html)

        html_by_file[os.path.basename(html_path)] = full_html

        print(f"Successfully converted to: {os.path.basename(html_path)}")

    progress.close()

    print("\nSuccess! Selected Word documents converted cleanly to HTML.")
    print("Images were auto-uploaded to your Mailchimp File Manager and linked by URL")
    print("(not embedded as base64, since Mailchimp strips those on send).")
    print(f"Local reference copies were also saved to '{os.path.join(SCRIPT_DIR, IMAGE_EXPORT_DIRNAME)}'.")

    show_html_output_window(html_by_file, selected_bullet_color)
