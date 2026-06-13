# Rostrum for Google Docs: Invisibility Mode

Bring Rostrum's **Invisibility Mode** and **debate styles** to a Google Doc, with
no Marketplace install. You paste one script file into the doc's own Apps Script
project; a **Rostrum** menu appears after a reload.

> **What "Hide" actually does.** Hide reduces every word of body text
> except headings, cites, and highlighted text to a 1‑point size, so a long
> file reads like a speech doc and the page count collapses. It does not conceal
> anything: anyone can reveal it with Select All → font size 11 (see *Reveal
> everything* below). This matches how Rostrum behaves in Word. **Show All brings
> every word back.**

Current version: **v0.1.0** (shown in the panel footer as `Rostrum v0.1.0`).

---

## 1. Install (paste-in, ~2 minutes)

You install per‑doc, into the document's *container‑bound* script. Nothing is
published anywhere; the script lives inside the doc.

### Get the two files

You paste in two small files: `Code.gs` (the script) and `appsscript.json` (the
permissions manifest). The friendliest path is the install page, which has both as
download buttons: **<https://andrewtjin.github.io/rostrum/google-docs.html>**.

Or grab them directly:

| File | Download | Direct copy (if the download doesn't start) |
|---|---|---|
| `Code.gs` | [Code.gs](https://rostrum-downloads.rostrum.workers.dev/code.gs) | [/google-docs/Code.gs](https://andrewtjin.github.io/rostrum/google-docs/Code.gs) |
| `appsscript.json` | (none) | [/google-docs/appsscript.json](https://andrewtjin.github.io/rostrum/google-docs/appsscript.json) |

(The first `Code.gs` link counts an anonymous download; the direct copy doesn't. Either
works; see *Privacy* below.) Then:

1. Open the Google Doc you want Rostrum in.
2. **Extensions → Apps Script.** A new editor tab opens with an empty `Code.gs`.
3. Select everything in `Code.gs` and delete it, then **paste the entire contents
   of the provided `Code.gs`** in its place.
4. Tell the project which permissions it needs (one‑time manifest step):
   - In the Apps Script editor, click **Project Settings** (the gear, left rail)
     and tick **"Show 'appsscript.json' manifest file in editor."**
   - Back in the **Editor**, open `appsscript.json` and replace its contents with
     the provided `appsscript.json` (it declares the Docs service and the two
     scopes below). Alternatively, just add the **Docs API** advanced service:
     **Editor → Services (＋) → Google Docs API → Add.**
5. **Save** (Ctrl/Cmd+S).
6. **Reload the Google Doc tab.** After the reload a **Rostrum** menu appears in
   the menu bar (next to *Help*).

### Permissions it asks for, and why
| Scope | Why |
|---|---|
| See, edit, create, and delete your Google Docs documents (`.../auth/documents`) | Read the doc once and apply font‑size‑only edits to hide/show and to set debate styles. |
| Display and run third‑party web content in prompts and sidebars (`script.container.ui`) | Draw the Rostrum panel and dialogs. |

Rostrum only ever changes **font size** (to hide/show) and **named‑style
definitions** (to apply debate styles). It never inserts, deletes, or reorders any
text character.

### Privacy: no tracking inside Google Docs
The pasted script **sends nothing back** as you use it: no document content, no
personal data, no usage pings. It runs entirely inside your doc and only edits font
sizes and style definitions. The *only* anonymous number Rostrum ever counts is the
one‑time `Code.gs` download from the website (and only if you used the counted download
link above, not the direct copy), exactly mirroring the Word side. Full disclosure:
the [privacy page](https://andrewtjin.github.io/rostrum/privacy.html).

---

## 2. First run: authorize, then run again

The **first** time you pick a Rostrum action, Google interrupts to ask for
permission. Your click is consumed by the consent flow, not the action.

1. Pick **Rostrum → Hide** (or any verb). A Google authorization dialog appears.
2. Choose your account.
3. Because this is your own unverified script, Google shows a warning screen.
   Click **Advanced**, then **"Go to *(your doc)* (unsafe)."** (This is expected
   for a personal script you pasted yourself; you are granting it to *your own*
   doc.)
4. Review the permissions and click **Allow.**
5. The dialog closes **without running the action**. This is normal. **Click the
   same Rostrum action again** and it runs. (The panel's first‑run line says the
   same thing: *"First time? Google uses your first click to ask permission, just
   click again."*)

> **Which Google account does Rostrum use?** It runs as the account that
> authorized it, which is the account you are signed into **in this Chrome
> profile** when you click *Allow*. If you are signed into several Google
> accounts in the same browser, open and run the doc under that **same** account,
> or the menu actions fail to authorize. The simplest setup is a Chrome profile
> signed into exactly the one account you want Rostrum to act as. (During testing,
> authorization succeeded only for the account Chrome was logged in with. This is
> that behavior, not a bug.)

> On a **managed / school (Workspace Education)** account, admin policy may block
> the consent screen entirely. If you hit a hard block, install on a **personal
> Gmail** account instead.

---

## 3. Using it

Open **Rostrum** in the menu bar. Entries, in order:

| Menu item | What it does |
|---|---|
| **Hide** | Shrink all body text except keepers (headings, cites, highlighted text). |
| **Show All** | Restore every word to its saved size; bring any stray tiny text back to normal. |
| **Apply debate styles** | Redefine this doc's styles to the debate convention and restyle existing headings (see below). |
| **Mark cite** | Mark the selected line as a cite (bold, cite size) so Hide keeps it. |
| **Open Rostrum panel…** | The sidebar: Hide/Show All buttons, keep‑colors, spacing toggle, the shortcut cheat sheet. |
| **Help & shortcuts** | What Hide does, the reading‑seams note, how to reveal everything, and team norms. |
| **Diagnostics** | A technical report you copy into chat when reporting back. |

### Debate styles & keyboard shortcuts
**Apply debate styles** sets this doc's named styles to:

| Style | Maps to | Size |
|---|---|---|
| Pocket | Heading 1 | 26 pt bold, boxed |
| Hat | Heading 2 | 22 pt |
| Block | Heading 3 | 16 pt |
| Tag | Heading 4 | 14 pt bold |
| Normal | Normal text | 11 pt (spacing zeroed) |

After applying, style a line with the **native Docs heading shortcuts**:
**Ctrl+Alt+1…4** on Windows, **⌘+Option+1…4** on Mac (Pocket/Hat/Block/Tag);
**…+0** for Normal. (Add‑ons can't register their own keyboard shortcuts, so
Rostrum rides the native ones; the panel's cheat sheet maps them to your Verbatim
F‑keys.) If a shortcut doesn't take, the **Rostrum menu and panel do everything
the shortcuts do.**

For Verbatim muscle memory, the F‑keys map across as: **Pocket F4 · Hat F5 ·
Block F6 · Tag F7 · Normal/clear F12.** **Cite is Verbatim F8**, which isn't a
heading, so use **Rostrum → Mark cite** (there's no native Docs chord for it,
which is why it isn't in the heading cheat sheet).

### Keep colors
Hide keeps highlighted text visible. By default it keeps the standard highlight
palette; the panel lets you choose exact colors or **"Keep any highlight color."**

---

## 4. Sharing with teammates

A Google Doc's bound script **travels with a copy.** To give a teammate the styled,
Rostrum‑enabled doc: **File → Make a copy** (or share a copy). The copy carries the
bound script, so the teammate gets the **Rostrum** menu after their own first‑run
authorization. There is no central install to manage.

**Team norm:** don't edit a doc while it's hidden. **Show All first.** Moving or
copying hidden text can separate it from its saved sizes; Show All still brings it
back (with a note), but editing un‑hidden is cleaner.

---

## 5. Reveal everything (no add‑on needed)

If you (or anyone you send the doc to) ever need text back **without** Rostrum:

> **Select All** (Ctrl+A / ⌘+A), then set the **font size to 11.** Everything
> becomes visible again.

---

## 6. Limits in this version

- **Single‑tab docs only.** A doc using Google Docs *tabs* is not supported;
  Rostrum stops and explains why (move your speech to a doc without extra tabs).
- **Suggesting mode:** Hide and Apply debate styles refuse while there are
  unresolved suggestions (resolve them first). **Show All always works**, even
  with suggestions present.
- **Cut/paste of hidden text** can detach it from its saved sizes; Show All then
  normalizes it back to full size and tells you how many passages it reset.
- Condense/Shrink and Marketplace publication are **not** part of this version.
