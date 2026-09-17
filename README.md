# TIDP Builder (proof of concept)

Static web app that generates a Bond Bryan Task Information Delivery Plan from the practice Library and exports a macro-free Excel workbook plus an internal resourcing note. No server, no macros, nothing to install.

## Files

| File | Role |
|---|---|
| `index.html` | The page (layout and styles) |
| `app.js` | Logic: selection, storey/volume duplication, Excel and Word export |
| `library.json` | **The Library.** One record per deliverable, plus value lists. This is the only file the SWG maintains. |
| `logo.png` | Logo used on the page and the Excel cover |
| `vendor/exceljs.min.js`, `vendor/docx.umd.js` | Bundled open-source libraries (MIT) for Excel and Word generation. Kept locally so the page works without third-party CDNs. |

## Publish on GitHub Pages

1. Create a repository (e.g. `bondbryan/TIDP`), copy this folder's contents to its root.
2. Settings > Pages > Source: *Deploy from a branch*, branch `main`, folder `/ (root)`.
3. The app is live at `https://bondbryan.github.io/TIDP/` within a minute.

Opening `index.html` by double-clicking will not load the Library (browsers block `fetch` from `file://`). Use the Pages address, or run a local server (`python3 -m http.server`) when testing.

## Maintaining the Library

Edit `library.json` in GitHub (pencil icon, then *Commit changes*). Each document record looks like:

```json
{
 "id": "Drawings-D-2001",
 "tab": "Drawings",
 "series": "2000 Series - Design GA Plans",
 "originator": "BBA", "volume": "ZZ", "level": "XX", "type": "D",
 "number": "2001",
 "description": "Design GA Plans",
 "format": "PDF/DWG", "scale": "1:200 @ A1", "workPackage": "ZZ: Multiple",
 "stages": {"1": null, "2": "YYYY-MM-DD", "3": "Ongoing", "4": "Ongoing", "5": "Ongoing", "6": null, "7": null},
 "tender": {"3+": true, "4a": true, "4b": true},
 "status": "Internal Team Item - Project Lead to allocate resource acordingly",
 "comments": "",
 "sectors": {"Education": true, "Residential": true, "...": true},
 "floorDup": true,
 "volDup": true,
 "buildType": "any"
}
```

Rules (same as the spreadsheet Library):

- Keep records grouped by `tab` then `series` in numeric order; the app writes them in file order.
- `id` must be unique: `<tab>-<type>-<number>`.
- `stages`: `"YYYY-MM-DD"` marks the stage a document is first issued (the app replaces it with the project's stage date); `"Ongoing"` for later stages; `null` when not required.
- `floorDup` / `volDup`: duplicate the row per ticked storey / volume. Storey rows get the level code and the description becomes "Ground Floor <description>" ("Plans" becomes "Plan").
- `buildType`: `"any"` (default), `"existing"` (pre-selected only when the project involves existing buildings), `"new"` (new build only).
- Bump `meta.version` when you publish a change; saved project files record which Library version they were made with.

A JSON validator (VS Code, or https://jsonlint.com) catches typos before committing. Any malformed file shows "Library not loaded" on the page rather than a broken TIDP.

## Project files

*Save project file* downloads `<code> - TIDP - <rev>.tidp.json`: the project set-up and the manual ticks. Store it in the project's TIDP folder next to the exported Excel; *Open a saved TIDP* restores it for the next revision. The browser also keeps a draft locally, for convenience only.

## Excel export

Sheets: `Cover`, one sheet per document type with rows (`Drawings`, `Images`, `Lists`, `Models`, `Text`, `Video`), then `Data`. Internal items, Status and Outsourcing Comments are never exported. `Data` is one row per document, every cell a formula linking to the content sheets, so edits made in Excel flow through to the raw list. No macros, no external links, no data validation to break.

## Resourcing note

Available when *Outsourcing* is ticked on step 1. A Word document for the Project Lead listing Outsourcing Partner items, items needing scope confirmation and internal items, with the Library comments. Internal only.
