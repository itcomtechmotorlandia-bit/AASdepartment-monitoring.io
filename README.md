# Service & Facility Dashboard

A static website. No server, build step or install is needed.

## How to open
Open `login.html` (or `index.html`, which sends you to the login page). Only three authorized emails can log in (see `auth.js`). Passwords are stored only as salted hashes.
The page opens in a browser in a browser (Chrome, Edge, Firefox or Safari).
An internet connection is needed the first time for the Inter font and the Excel/PDF export libraries (loaded from a CDN).

## Files
- `login.html` – login page
- `auth.js` – the list of users (passwords are stored as hashes)
- `guard.js` – sends people who are not signed in back to the login page
- `index.html` – page structure (sidebar, pages, tables)
- `styles.css` – all styling, themes and hover/animation effects
- `app.js` – all logic: dashboards, formulas, forms, saving and export

## Where the data is saved
Everything you type is saved in the browser (localStorage) on the computer that opened the site.
It is not shared between people or computers. Clearing the browser data erases it.

## To put it online
Upload all files (all in one folder, next to index.html) to any static host, such as GitHub Pages, Netlify or a company web server.

## Logins
Three emails have access: motoassap@gmail.com, aas.dept01@gmail.com and itcomtechmotorlandia@gmail.com.

**Forgot password:** click "Forgot password?" on the login page, enter your email, the recovery code (kept by the administrator), and a new password.
The new password is saved in that browser only, so on another computer or after clearing browser data the original password applies again.

**Important:** this login runs only in the browser. It keeps casual visitors out, but anyone who knows web development can get around it,
and it cannot send reset emails. For real security and email password resets, use a server login such as Firebase Authentication.
