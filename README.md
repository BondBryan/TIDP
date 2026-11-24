# TIDP Builder

Static web app that generates a Bond Bryan Task Information Delivery Plan from the practice Library and exports a macro-free Excel workbook and, when outsourcing is ticked, the project Outsourcing Tracker. No server, no macros, nothing to install.

## Files

| File | Role |
|---|---|
| `index.html` | The page (layout and styles) |
| `app.js` | Logic: selection, storey/volume duplication, Excel export, password |
| `library.json` | **The Library.** One record per deliverable, plus value lists. This is the only file the SWG maintains. |
| `logo.png` | BondBryan logo: the page and the Excel Cover |
| `logo-bbf.png`, `logo-interiors.png` | Cover logos for BondBryan:Fairhursts / Fairhursts and for Interior Design (export only) |
| `tracker-template.xlsx` | Copy of the practice Outsourcing Tracker template (must contain the `Suitable Drawings` sheet). Replace it here whenever the template changes. |
| `vendor/exceljs.min.js`, `vendor/jszip.min.js` | Bundled open-source libraries (MIT) for Excel generation and for writing into the tracker template. Kept locally so the page works without third-party CDNs. |

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
 "stages": {"0": null, "1": null, "2": "YYYY-MM-DD", "3": "Ongoing", "4": "Ongoing", "5": "Ongoing", "6": null, "7": null},
 "tender": {"3+": true, "4a": true, "4b": true},
 "gateway": {"G1": true, "G2": true, "CR": true},
 "status": "Internal Team Item - Project Lead to allocate resource accordingly",
 "comments": "",
 "sectors": {"Education": true, "Residential": true, "...": true},
 "floorDup": true,
 "volDup": true,
 "buildType": "any"
}
```

Rules (same as the spreadsheet Library):

- Keep records grouped by `tab` then `series` in numeric order; the app writes them in file order.
- `id` must be unique: `<tab>-<type>-<number>`. When the same deliverable needs a different record per discipline (different default or name), or the plain id is already used by an architecture record, add a suffix: `-L` (all three Landscape disciplines), `-LE`, `-LR`, `-LO`, `-LRO` (Landscape Education, Residential, Other, or Residential + Other) and `-I` (Interior Design). Ids are internal only: they never show on screen or in the export.
- `stages`: `"YYYY-MM-DD"` marks the stage a document is first issued; `"Ongoing"` marks the later stages it is issued in; `null` when not required. In the export every issued stage shows the project's date for that stage (first issue in bold); `YYYY-MM-DD` in red until the date is entered on step 1.
- `gateway`: ticks for the Gateway 1, Gateway 2 and Construction Record columns of the export (all sectors).
- `Berkeley_Naming_Type` / `Berkeley_Naming_Code` (Residential records only): the Berkeley Group document type (e.g. `DR`) and code (Berkeley series + Bond Bryan number, e.g. `005-2001`). Used instead of `type` / `number` in the export when the project is a Berkeley project. Records without them keep the Bond Bryan type and number.
- `sectorOverrides` (optional): stages, tender ticks or status that replace the record's own values when the project is set to that sector, e.g. `"sectorOverrides": {"Residential": {"stages": {...}, "tender": {...}}}`. Used for the Residential Deliverables Tracker alignment; leave it out when the sector follows the record's values.
- `visibility` (optional): list of the sectors or disciplines in which the item is shown and can be ticked, e.g. `"visibility": ["Landscape_Residential", "Landscape_Other"]`. Without it, the item follows the general rules (discipline items hidden for the other sectors). Used by the scope items.
- **Scope items**: records with the status "Confirmation of scope required" are off for every sector (`sectors` all `false`) and carry a `visibility` list. The Project Lead ticks them in the scope review, which always opens once before the export. An item that is part of the standard offer for a sector is not a scope item: give it the normal status and switch it on in `sectors`. When it is standard offer for one sector only and a scope item elsewhere, keep the scope status, switch that sector on in `sectors` and add the normal status under `sectorOverrides`, e.g. `"sectorOverrides": {"Residential": {"status": "Internal Team Item - Project Lead to allocate resource accordingly"}}`.
- `floorDup` / `volDup`: duplicate the row per ticked storey / volume. Storey rows get the level code and the description becomes "Ground Floor <description>" ("Plans" becomes "Plan").
- `buildType`: `"any"` (default), `"existing"` (pre-selected only when the project involves existing buildings), `"new"` (new build only).
- Bump `meta.version` when you publish a change.
- `disciplineBySector` gives the discipline name used in the exported file names (`default` is `Architect`), alongside `roleBySector` for the role code.

A JSON validator (VS Code, or https://jsonlint.com) catches typos before committing. Any malformed file shows "Library not loaded" on the page rather than a broken TIDP.

## Password

The first pop-up asks for the password published on the TIDP HUB page; nothing works until it is entered (once per browser tab). `app.js` holds only the SHA-256 hash of the password (`PW_HASH`). To change the password, run `printf 'NewPassword' | shasum -a 256` in Terminal, paste the result into `PW_HASH`, commit, and update the HUB page. This keeps casual visitors out; it is not strong security, as the page and the Library stay publicly downloadable from GitHub Pages.

## Project set-up

*Issuing Company* (BondBryan, BondBryan:Fairhursts or Fairhursts) sets the originator code in every file reference (`BBA`, or `BBF` for both Fairhursts options) and the logo on the exported Cover. Interior Design exports carry the Interiors logo whichever company issues. The *Internal* Library items are kept in `library.json` but are not displayed or exported.


Nothing is saved: the page holds no project data, and each TIDP is issued as revision P01 dated on the day of export. *Reset* on step 1 clears the set-up and the ticks and shows the guidance pop-ups again. On step 1, untick the RIBA stages outside the appointment: a Library item whose stages all fall outside the ticked ones is unticked by default on step 2 (items with no stage information are unaffected), and unticked stages are left blank in the export. The tender stages (3+, 4a, 4b) ticked on step 1 are the only tender columns exported.

## Berkeley projects

Choosing *Residential* on step 1 asks "Is this a Berkeley project?". *Yes* switches the export to the Berkeley document types and codes (Type and Number columns, Data sheet, Outsourcing Tracker and the TIDP's own file reference, e.g. `PR` `010-1001`). The choice shows on step 1 as a tick box and can be changed at any time; choosing another sector switches it off. Project code, originator and revision stay as set by the Project Lead.

## Export

One button exports the TIDP and, when *Outsourcing may be used on this project* is ticked, the Outsourcing Tracker:

- `<code>-<office>-XX-XX-L-<role>-1001 Task Information Delivery Plan_<discipline>.xlsx`
- `<code>-<office>-XX-XX-L-<role>-7001 Outsourcing Tracker_P01.xlsx`

### Excel

Sheets: `Cover`, one sheet per document type with rows (`Drawings`, `Images`, `Lists`, `Models`, `Text`, `Video`), `Data`, then `Data fields` (hidden). Gridlines are off everywhere except `Data`.

- `Cover`: logo, Project, Project code, File reference, Date, Revision, Issued by, Issuing Company. The file reference follows the Revision cell.
- Rows 2 to 6 of each content sheet (Issuing Company, Project Name, File Ref, Date, Revision) are formulas linked to the Cover. A data validation rule refuses anything typed over them with the message "Change this information in the Cover sheet." (Excel does not run validation when a cell is cleared with Delete or pasted over.)
- Level (D), Format (I) and Work Package (K) are drop-downs fed by the hidden `Data fields` sheet, over the table and 200 rows below it. To accept a new value, unhide `Data fields` and add it inside the list.
- Each RIBA stage column is followed by a narrow `Progress` column, set to 0% wherever the document is issued at that stage: staff overtype 0 to 100, shown as a percentage with a progress bar (other values are refused). Each series header row shows the average of the Progress values entered in its series at that stage, whether or not the stage has a date.
- Series header rows carry a marker in the stage, tender and gateway columns so they stay visible when a column is filtered; it is the same colour as the cell and cannot be seen.
- After the tender columns: `Gateway 1`, `Gateway 2` and `Construction Record` tick columns.
- Within a series, rows are ordered by volume / system (column C), then level (column D), then number (column G). Internal items, Status and Outsourcing Comments are never exported. `Data` is a live list of every document on the content sheets: one Microsoft 365 formula in A2 (LET, VSTACK, FILTER, XLOOKUP) reads rows 9 to 2000 of each content sheet, so rows renamed, inserted, added or deleted in Excel flow through to the raw list. The sheet is locked with no password: it can be selected, copied, filtered and its columns resized, but not edited (Review > Unprotect Sheet if ever needed). Needs Microsoft 365 Excel. `DATA_MODE` in `app.js` switches back to the previous one-link-per-cell sheet (`'links'`, unlocked). No macros and no external links.

### Outsourcing Tracker

Exported only when outsourcing is ticked on step 1. It is `tracker-template.xlsx` unchanged, with the project name and code written on `Summary` and the `Suitable Drawings` sheet filled with every ticked document the Library flags as suitable for outsourcing: TIDP reference, sheet, description, format, first stage and the Library outsourcing note. Column A (*For outsourcing*) is a YES / NO drop-down left blank for the Project Lead; YES rows turn green, NO rows grey. The other Library status notes (scope confirmation, internal team) stay in `library.json` for later use.
