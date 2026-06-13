# Install Rostrum for Microsoft Word

Rostrum for Word is a desktop **Office.js add-in**. You install it by registering one
manifest file in Word, **once**. The add-in itself runs from the hosted bundle, so you
host nothing and run no server. (Looking for the Google Docs version instead? See
[`google-docs/README.md`](../google-docs/README.md). Back to the project overview:
[README](../README.md).)

> **Desktop Word only.** Rostrum hides text with Word's *Hidden font* attribute
> (`<w:vanish/>`), which is `WordApiDesktop 1.2`, present on **Word for Windows and Mac**,
> absent on **Word for the web** and Office 2016–2021 perpetual. On an unsupported host the
> pane explains the manual fallback instead of failing silently. (Web Word support is on the
> roadmap; because the web host lacks the hidden-font attribute, it would behave differently
> from desktop and is not promised here yet.)

The fastest path for most debaters is the hosted install page, which links the production
`manifest.xml` and walks the per-OS steps with screenshots:

> **https://andrewtjin.github.io/rostrum/** → choose **Microsoft Word**

---

## Requirements

Word for **Windows or Mac** (desktop). Manifest floor: `WordApiDesktop 1.2` + `WordApi 1.4`
(the hide mechanism + the in-document manifest store + the Track-Changes gate). **Apply
Styles** additionally needs Word 1.5+. Node ≥ 18 only for local development.

---

## Debater install (hosted build)

You register one `manifest.xml` in Word; the add-in code is served from GitHub Pages. This
is Stage A distribution, for motivated early adopters; a future Stage B adds one-click AppSource
install.

### Windows

Rostrum installs as a **trusted add-in catalog**: a shared folder holding `manifest.xml`.
Once registered, the **Rostrum** tab appears on the ribbon of **every document
automatically** (new and existing) with no per-document re-adding, and it survives
restarting Word.

1. Save `manifest.xml` (from the [install page](https://andrewtjin.github.io/rostrum/))
   into a folder, e.g. `C:\Users\you\Documents\rostrum-addin`.
2. Right-click that folder ▸ **Properties ▸ Sharing ▸ Share…**, pick your own Windows
   account, click **Share**, then **Done**. Copy the network path Windows shows (it looks
   like `\\YOUR-PC\rostrum-addin`). Word needs this share to load Rostrum on every document.
   *(If **Share…** is greyed out (common on school- or work-managed laptops), file sharing
   may be blocked by your administrator; [open an issue](https://github.com/andrewtjin/rostrum/issues)
   and we'll help with an alternative.)*
3. In Word: **File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸ Trusted Add-in
   Catalogs**.
4. Paste the share path into *Catalog Url*, click **Add catalog**, tick **Show in Menu**,
   then **OK**.
5. **Restart Word.**
6. **Insert ▸ My Add-ins ▸ Shared Folder ▸ Rostrum ▸ Add.** The Rostrum tab is now on the
   ribbon of every document.

✓ You should now see a **Rostrum** tab on the Word ribbon. Don't see it? Re-check steps 2–4;
the share path is the usual culprit.

### Mac

1. Save `manifest.xml`. In Finder, press <kbd>⌘</kbd>+<kbd>⇧</kbd>+<kbd>G</kbd> (**Go ▸ Go to
   Folder**), since the folder below is hidden and you can't browse to it, paste the path, and
   drop `manifest.xml` in (create the `wef` folder if it doesn't exist):
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`
2. **Restart Word.**
3. **Insert ▸ My Add-ins ▸ Rostrum.** (If you don't see it, also check the *Developer
   Add-ins* group of that menu.)

✓ You should now see a **Rostrum** tab on the Word ribbon.

### Updates

Rostrum runs from the hosted site, so you **don't reinstall to get fixes**:

- **Fixes & improvements: automatic.** Behavior changes (speed, bug fixes, tweaks to
  existing buttons) are served live. Just **restart Word**, or, with the Rostrum pane open,
  click inside it and press <kbd>Ctrl</kbd>+<kbd>F5</kbd>. You re-download nothing.
- **New version: re-download `manifest.xml`.** When a release adds a *new ribbon button or
  tab*, the change lives in the manifest file you registered. Grab the latest `manifest.xml`,
  replace your saved copy, clear the Office cache (see Uninstall), and restart Word.

Version announcements: the [GitHub releases page](https://github.com/andrewtjin/rostrum/releases).
Rule of thumb: if a button you expect is missing after an update, re-download the manifest.

### Uninstall

Rostrum installs no program; removing it means dropping the manifest registration and
clearing Word's add-in cache.

- **Windows:** **Insert ▸ My Add-ins**, right-click **Rostrum ▸ Remove** (current doc). To
  remove everywhere: close Word ▸ **File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸
  Trusted Add-in Catalogs**, untick **Show in Menu** on the Rostrum row, tick **clear all
  previously-started web add-ins cache** (Word 2108+) ▸ **OK** ▸ restart Word. Older Word
  without that checkbox: with Word closed, delete everything inside
  `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` (this clears *all* sideloaded add-ins).
- **Mac:** quit Word, delete `manifest.xml` from
  `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`, restart Word.

---

## Developer sideload

```bash
npm install
npm start          # builds, starts the HTTPS dev server on :3000, and sideloads
```

`npm start` (office-addin-debugging) trusts the dev certificate, serves the bundles over
HTTPS (Office requires HTTPS), and sideloads `manifest.xml` into Word. Then use the Rostrum
ribbon tab: the Hide / Show All / Apply Styles buttons directly, or Options on any group to
open that tool's pane.

To stop: `npm stop`. To rebuild without sideloading: `npm run build` (production) or
`npm run build:dev`.

## Producing the production manifest

The committed `manifest.xml` is the **dev** manifest (`https://localhost:3000`). The
production manifest is a build artifact, generated against the hosted origin and never
committed, so the manifest drift test stays green:

```bash
npm run build                                                            # dist/ (hashed bundles, assets, landing page)
npm run gen:manifest:prod -- --origin=https://andrewtjin.github.io/rostrum   # → dist/manifest.xml
```

`gen:manifest:prod` rebases every `SourceLocation` / icon / support URL onto the origin and
stamps the production `<Id>` (distinct from dev, so both can be sideloaded on one machine).
It errors if `--origin` is missing or not `https://`. CI
(`.github/workflows/deploy-pages.yml`) runs exactly these steps on every push to `master`
and publishes `dist/` to GitHub Pages, gated on green tests + a clean typecheck. (One-time:
repo Settings ▸ Pages ▸ Source = GitHub Actions.)

---

Reading a Rostrum document **without** the add-in: select all (<kbd>Ctrl</kbd>+<kbd>A</kbd>),
open the Font dialog (<kbd>Ctrl</kbd>+<kbd>D</kbd>), and clear the **Hidden** checkbox, or
toggle **Home ▸ ¶** to view hidden text. Show All in the add-in does the same.
