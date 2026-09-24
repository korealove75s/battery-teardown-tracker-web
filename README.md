# Battery Teardown Tracker (web)

https://sangnew.github.io/battery-teardown-tracker-web/

The UI is in English. Data is stored in two ways:

- **Shared database mode (Firebase)**: once `firebase-config.js` is filled in, everyone sees the same data on any PC or IP after signing in with the team password. Changes appear on other screens in real time.
- **Browser-only mode**: while `firebase-config.js` is empty, data is saved only in this browser's IndexedDB (same as the earlier version).

## Firebase setup (once, about 5 minutes)

1. Go to https://console.firebase.google.com → **Add project** (Google Analytics not needed).
2. **Build → Firestore Database → Create database** → choose a location (e.g. `asia-northeast3` Seoul) → start in **production mode**.
3. On the Firestore **Rules** tab, paste the contents of `firestore.rules` from this folder → **Publish**.
4. **Build → Authentication → Get started → Sign-in method → Email/Password → Enable**.
5. **Authentication → Users → Add user**
   - Email: `team@example.com` (must match `BTT_TEAM_EMAIL` in `firebase-config.js` and the email in `firestore.rules`)
   - Password: the team password you want to use (at least 6 characters)
6. **Project settings (gear icon) → General → Your apps → Web (`</>`)** → register an app (Hosting not needed) → copy the `firebaseConfig` values into `firebase-config.js`.
7. `git add . && git commit -m "Configure Firebase" && git push`

> Note: `apiKey` in `firebase-config.js` is not a secret; Firebase web apps always expose it. Access is controlled by the team password and `firestore.rules`.

To change the team password: Firebase console → Authentication → Users → the account's ⋮ menu → Reset password. Browsers that are already signed in stay signed in, so to cut access, choose **Disable account** and then enable it again.

## Moving existing data (from browser-only mode)

After sign-in, if the browser still has data saved by the old version, a yellow banner appears. Click **Upload to shared database** to upload records and images. Do this once on each PC that has data.

## Features

- Two tabs: **Tear Down** / **Frozen IR · Spot Analysis**
- New records, inline cell editing, add/delete rows, paste from Excel (Ctrl+V, header rows matched by column name)
- CSV/Excel import and export (for the current tab)
- Cell ID search, filters, duplicate Cell ID warnings
- Multiple VHX/Image attachments per row (in shared mode, images are auto-compressed to under about 1 MB each)
- **Reset**: delete the current tab or all tabs. You must type `RESET` to confirm. In shared mode this deletes the data for the whole team.
- TD Dates without a year (e.g. `1/26`) ask for the year when saved.

## Local testing

```
npx serve web
```

## Importing from the OCV tracking workbook (Frozen IR · Spot Analysis tab)

The former "OCV Report Builder" site is now part of this tab
(https://sangnew.github.io/ocv-report-builder-web/ redirects here).

1. Open the **Frozen IR · Spot Analysis** tab and click **Import OCV workbook**.
2. Choose the *Mass Production E&L Grade OCV Tracking Sheet* .xlsx. It is read in the browser and never uploaded.
3. Pick LOTs and click **Import**. Rows are added or updated by Cell ID:
   - New cells are added with the values found in the workbook.
   - For existing cells, blank fields are filled, and fields that still hold the previous import's value are refreshed. **Values the team edited are kept.**
4. Cells that still need manual input are highlighted **yellow**. **Light blue** means the tracking-sheet analysis disagrees with Master E & L, so please verify it. Hover a cell to see where its value came from. Use **Needs input only** (or the ⚠ chip) to list the rows left to fill in.
5. **Export Excel** on this tab writes the report layout (headers on row 2, No. in column B), keeps the highlights, and adds a Legend sheet.

Voltage drop rule: for each layer (column B) of the cell's own tracking sheet, dOCV = the largest fall between the tracking dates in C/D/E. A cell is **Drop** when one inner layer is more than 2.6σ above the others (the sheet's R6 formula) **and** that fall is at least 1.5 mV (adjustable in the import dialog). Otherwise it is **NTF**. The conversion rules are in `convert.js`.

## Weekly low-voltage PowerPoint (report.html)

https://sangnew.github.io/battery-teardown-tracker-web/report.html (also linked from the tracker header as **Weekly PPT report**)

The page builds the weekly "E81C 양산랏 저전압 현황" deck: 9 A4 slides with native, editable charts. All files are read in the browser and nothing is uploaded.

| Input | Example | Used for |
|---|---|---|
| OCV tracking workbook (required) | Mass Production E&L Grade OCV Tracking Sheet | Cells, drop analysis from each tracking sheet, EDS, top/back, x/y, layer |
| Analysis report (recommended) | test1.xlsx or the Analysis tab's *Export Excel* | Location (inside/outside), Shape, reviewed values |
| Lot summary (required) | test2.xlsx | Monthly trend, production, E/L rates |
| Genealogy CSV (optional) | 9.241.csv | AZS stacker / DNC (slide 5) and electrode lots (slides 6–7) |

Slides: 1 monthly trend with improvement notes, 2 five-analysis results, 3–4 foreign-material locations, 5 coating-top AZS/DNC, 6–7 cathode/anode electrode lots, 8 cumulative results, 9 C4~SOP history (static data from the original deck).
Improvement notes and headline sentences can be edited on the page. The code is in `report-model.js` (numbers) and `report-ppt.js` (layout).
