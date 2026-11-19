# GCV Updater

![Release](https://img.shields.io/github/v/release/Gaza-children-village/excel-updater?filter=updater-v*&label=release)
![Excel](https://img.shields.io/badge/Excel-2016%2B%20Windows-217346)
![VBA](https://img.shields.io/badge/VBA-add--in-blue)

Excel add-in (`GCV Updater.xlam`) that installs updates and new GCV add-ins from GitHub releases, and the home of the pieces that tie the add-ins together: the **catalog**, the **loader** and the **installer**.

## Ribbon

| Button | What it does |
|---|---|
| **تحديث الأدوات** (update tools) | Compares every installed GCV add-in with its newest release, then installs updates and new tools |
| **حول الأدوات** (about) | Versions of the loader and of every loaded add-in |

## How an update works

1. Read [`catalog.json`](catalog.json) (this repository, branch `main`): every GCV add-in with its repository and whether it is required.
2. For each add-in, find the newest stable release tagged `<package>-v<version>` and compare it with the installed version.
3. One question lists the updates and new required add-ins; each new optional add-in is offered separately.
4. Download `manifest.json` and the add-in; check id, version, size and **SHA-256**.
5. Install while Excel runs: close the add-in, keep the current file in `previous\`, put the new one in place, open it and check it answers with its version. If it does not start, the previous version is put back and that version is never offered again.
6. The updater's own update, and files another Excel window holds, are installed by GCV Loader at the next Excel start.

No access to the VBA project object model is needed.

## Adding a new tool

1. Create a public repository `excel-<tool>` with an add-in built the same way (shell with the shared ribbon tab `idQ="gcv:GcvToolsTab"`, `GcvAddin.bas`, `GcvLogClient.bas`).
2. Publish its first release with the release tool.
3. Add one entry to `catalog.json`:

```json
{"id": "gcv/<tool>", "name": "GCV <Tool>", "file": "GCV <Tool>.xlam",
  "repo": "Gaza-children-village/excel-<tool>", "required": false,
  "description": "One sentence shown when the tool is offered."}
```

The next time a secretary presses **update tools**, the new tool is offered.

## Installing on a PC

[`installer/install-gcv-tools.bat`](installer/install-gcv-tools.bat), for the current Windows user, **without** administrator rights (Excel closed):

1. Copies `GCV Loader.xlam` to `%LOCALAPPDATA%\GCV\ExcelTools` and every other `GCV *.xlam` next to the script to `...\ExcelTools\addins`.
2. Makes that folder a trusted location (no macro warnings).
3. Registers GCV Loader, so Excel opens all GCV add-ins at start-up.

Put in one folder: both installer files, `GCV Loader.xlam`, `GCV Logger.xlam`, `GCV Updater.xlam` (required) and the other add-ins. Running it again is safe. `uninstall-gcv-tools.bat` removes everything except the logs.

Tip: attach that folder as `GCV-Tools-Setup.zip` to a release of this repository, so there is always one download for a new PC.

## GCV Loader

[`loader/`](loader): the small add-in Excel opens at start-up. It applies waiting updates, opens every add-in in `addins\` (logger first, updater last), and falls back to the previous version of an add-in that does not start. It changes rarely and is updated by hand.

```
%LOCALAPPDATA%\GCV\ExcelTools\GCV Loader.xlam
%LOCALAPPDATA%\GCV\ExcelTools\addins\     installed add-ins (names never change)
%LOCALAPPDATA%\GCV\ExcelTools\previous\   version before each update
%LOCALAPPDATA%\GCV\ExcelTools\pending\    updates waiting for the next Excel start
```

## Building from source

Excel cannot load a VBA project created outside Excel, so the add-in is assembled in Excel from the files in this repository:

1. Copy `GCV Updater - shell.xlam` and rename the copy to `GCV Updater.xlam` (the shell already contains the add-in file type and the ribbon).
2. Open it, press **Alt+F11** and select its VBAProject.
3. **File › Import File**: every file in `src/`.
4. Double-click **ThisWorkbook** and paste the contents of `ThisWorkbook.code`.
5. **Debug › Compile VBAProject**, then save with **Ctrl+S**. Saving records the add-in's id, version and file name in the file.

Source files are ASCII with CRLF line endings. Arabic text is stored as Unicode code points (`FromCodePoints("0627 ...")`), because the VBA editor cannot keep Arabic characters.

## Tests

Run `RunUpdaterTests` (and `RunLoaderTests` in the loader) from the VBA editor (click inside the procedure and press **F5**). Results appear in a message box and in the Immediate window (**Ctrl+G**). The tests use fakes and temporary files: no network, no real data.

Before a release, remove the development modules (`TestUpdater`, `UpdFakeHttp`), compile and save.

## Releasing

```
python make_addin_release.py "GCV Updater.xlam" --changelog CHANGELOG.md --repo Gaza-children-village/excel-updater
```

The release tool (in [excel-updater/release-tools](https://github.com/Gaza-children-village/excel-updater/tree/main/release-tools)) checks the file (add-in type, ribbon, no test modules, version listed in `CHANGELOG.md`), writes `manifest.json` with the SHA-256 and size, and prints the `gh release create updater-v<version> ...` command. Tags are `updater-v<version>`. Drafts and pre-releases are ignored by the updater.

The release tool itself is tested with `pytest` (mypy strict, ruff): `cd release-tools && python -m pytest`.

## Trust model

Checksums come from the same GitHub release as the files: they catch damaged or incomplete downloads, not a compromised account. Protect the organisation with two-factor authentication and limit who can publish releases.

## Files

| Path | Purpose |
|---|---|
| `catalog.json` | The list of GCV add-ins the updater offers |
| `src/Updater.bas` | Ribbon, catalog, plan, install and fallback |
| `src/UpdCore.bas`, `src/UpdJson.bas` | Versions, SHA-256, releases, manifests, downloads; JSON |
| `src/UpdWinHttp.cls`, `src/IUpdHttp.cls`, `src/UpdHttpResult.cls` | HTTPS through WinHTTP behind an interface |
| `src/GcvLogClient.bas`, `src/GcvAddin.bas` | Logging through GCV Logger; identity of the add-in |
| `src/TestUpdater.bas`, `src/UpdFakeHttp.cls` | Tests (development only) |
| `loader/` | GCV Loader source and shell |
| `installer/` | Install and uninstall scripts |
| `release-tools/` | `make_addin_release.py` and its tests |

## Part of GCV Tools

Excel add-ins for the schools of Gaza Children's Village. They share one ribbon tab, **Gaza Children Village Tools**, and are installed and updated together.

| Add-in | Repository | What it does |
|---|---|---|
| GCV Logger | [excel-logger](https://github.com/Gaza-children-village/excel-logger) | Logs of all GCV add-ins; show logs, report an issue |
| GCV Updater | [excel-updater](https://github.com/Gaza-children-village/excel-updater) | Installs updates and new tools; home of the catalog, loader and installer |
| GCV Student Lookup | [excel-student-lookup](https://github.com/Gaza-children-village/excel-student-lookup) | Fills student rows from the Civil Registry, on demand |
| GCV English Names | [excel-english-names](https://github.com/Gaza-children-village/excel-english-names) | `=TO_ENGLISH(cell)`: English spelling of Arabic names |

See [CHANGELOG.md](CHANGELOG.md).
