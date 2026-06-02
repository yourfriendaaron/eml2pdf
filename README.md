# eml2pdf — instant Outlook email → PDF

Drag an email from Outlook into **`~/Downloads`** (or the **Email-PDFs** Dock
stack). A few seconds later a PDF of that email appears in the same folder.
No print dialog, no printer loading, no spinner.

## Why this works

The slow part of "Print → Save as PDF" on a Mac is the **print system**
(loading printers + the CUPS spinner) — not making the PDF itself.

Dragging an email out of Outlook saves it instantly as a `.eml` file.
A background watcher converts that `.eml` to a real PDF and skips the print
system entirely.

```
Drag email from Outlook ─▶ ~/Downloads/foo.eml    (instant)
                                  │  launchd notices the new file
                                  ▼
                          eml2pdf (WeasyPrint → Chrome fallback)
                                  ▼
                          ~/Downloads/foo.pdf       (the .eml is moved
                                                     into .eml-processed/)
```

**Renderers (fastest first):**
1. **WeasyPrint** (`brew install weasyprint`) — pure HTML/CSS → PDF, no
   browser. Total conversion ~1s. This is the primary path.
2. **Headless Chrome** (`--print-to-pdf`) — automatic fallback for the rare
   email whose CSS WeasyPrint can't handle. ~4–6s (mostly Chrome's cold start).

The log line records which renderer was used, e.g. `OK [weasyprint] ...`.

## How Downloads works without Full Disk Access

`~/Downloads` is a macOS TCC-protected folder. A background `launchd` agent
running the *shared* system Python can't read it unless you grant Python
**Full Disk Access** — a broad grant that would apply to every Python script
on the machine.

To avoid that, the converter is packaged as a **signed `.app` bundle**
(`dist/eml2pdf.app`, built with PyInstaller, ad-hoc code-signed). Because it
has its own code identity, macOS attributes the permission to *the app* — so
on first run it shows the narrow **"eml2pdf wants to access your Downloads
folder"** prompt. Approving it grants **Downloads-only** access to this one
app. System Python keeps zero special permissions.

You can see/revoke the grant at
System Settings → Privacy & Security → **Files and Folders → eml2pdf**.

> ⚠️ **Rebuilding the app re-triggers the prompt.** The ad-hoc signature's
> identity is tied to the binary's hash, so if you rebuild `eml2pdf.app`
> you'll be asked to approve Downloads access again (and may want to
> `tccutil reset DownloadsFolder com.github.yourfriendaaron.eml2pdf` first).

## Files

| File | Purpose |
|------|---------|
| `eml2pdf.py` | Parses `.eml`, embeds inline (cid:) images, renders the PDF to `/tmp`, then writes it into the watched folder. Renders via WeasyPrint, falls back to headless Chrome. |
| `dist/eml2pdf.app` | Signed app bundle (PyInstaller) so macOS can scope the Downloads permission to this app instead of system Python. |
| `com.github.yourfriendaaron.eml2pdf.plist` | launchd agent that watches `~/Downloads` and `~/Email-PDFs` and runs the app on any change. Replace `YOUR_USERNAME` with your macOS short username before installing (launchd does not expand `~`/`$HOME`). |

Installed agent: `~/Library/LaunchAgents/com.github.yourfriendaaron.eml2pdf.plist`.
Logs: `~/Library/Logs/eml2pdf.log` (plus `.out.log` / `.err.log`).

Note: the renderer subprocesses (WeasyPrint/Chrome) don't have Downloads
access — only the app does — so the app renders to `/tmp` first and does the
final write into Downloads itself.

## Rebuilding the app (after editing `eml2pdf.py`)

```sh
cd path/to/eml2pdf
python3 -m PyInstaller --noconfirm --windowed \
  --name eml2pdf --osx-bundle-identifier com.github.yourfriendaaron.eml2pdf \
  --distpath dist --workpath /tmp/pyi-build --specpath /tmp/pyi-spec \
  eml2pdf.py
# re-apply the background-agent + usage-string keys to Info.plist, then:
codesign --force --deep --sign - dist/eml2pdf.app
```
(`Info.plist` needs `LSUIElement`, `LSBackgroundOnly`, and
`NSDownloadsFolderUsageDescription` — PyInstaller overwrites it on each build.)

## Manual use

```sh
# convert specific files (uses system python; fine for files outside Downloads)
python3 eml2pdf.py ~/Downloads/something.eml

# convert every .eml sitting in one or more folders
dist/eml2pdf.app/Contents/MacOS/eml2pdf --watch-dir ~/Downloads ~/Email-PDFs
```

## Managing the watcher

```sh
UID=$(id -u)
LABEL=com.github.yourfriendaaron.eml2pdf
launchctl list | grep eml2pdf                              # is it loaded?
launchctl kickstart -k gui/$UID/$LABEL                     # force a run
launchctl bootout gui/$UID/$LABEL                          # stop/disable
launchctl bootstrap gui/$UID ~/Library/LaunchAgents/$LABEL.plist  # re-enable
```
