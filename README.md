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
