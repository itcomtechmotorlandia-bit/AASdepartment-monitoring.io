# Service & Facility Dashboard

A static website. No server, build step or install is needed.

## How to open
Open `login.html` (or `index.html`, which sends you to the login page). Only the users added in Firebase Authentication can log in.
The page opens in a browser in a browser (Chrome, Edge, Firefox or Safari).
An internet connection is needed the first time for the Inter font and the Excel/PDF export libraries (loaded from a CDN).

## Files
- `login.html` – login page
- `auth.js` – Firebase settings and login functions
- `guard.js` – checks the login, loads/saves the shared data, then starts `app.js`
- `index.html` – page structure (sidebar, pages, tables)
- `styles.css` – all styling, themes and hover/animation effects
- `app.js` – all logic: dashboards, formulas, forms, saving and export

## Firebase (login + shared data)
- Project: `aasdepartmentmonitoring`. Login = Firebase Authentication (Email/Password). Data = Firestore collection `appdata`.
- Add or remove users in Firebase console > Authentication > Users. The rules below also list the allowed emails.
- "Forgot password?" sends a real reset email through Firebase.
- All users share the same data. Changes save automatically (the name at the top shows "saving..." then "saved"). If someone else changes the data, a "Refresh now" box appears.
- First login ever: the data already in that browser is uploaded to Firestore. After that, Firestore is the master copy and the browser copy is replaced on every login. Logging out clears the data from the browser.
