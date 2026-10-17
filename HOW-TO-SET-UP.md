# Kwentoink – How to set up 🐃

You have 5 files:

| File | Where it goes |
|---|---|
| `Code.gs` | Google Apps Script (the back end) |
| `index.html` | GitHub repo |
| `sw.js` | GitHub repo |
| `manifest.json` | GitHub repo |
| `logo.png` | GitHub repo |

The 4 GitHub files all go **together in the same repo, side by side**. No folders.

---

## Part 1 – Back end (Apps Script)

1. Go to **script.google.com** → **New project**. Name it `Kwentoink`.
2. Delete everything in the editor, paste all of **Code.gs**, and save.
3. At the top of Code.gs, fill in:
   - `APP_URL` – your GitHub Pages link for Kwentoink, ending with `/`
     (e.g. `https://yourname.github.io/kwentoink/`). You'll get this in Part 2, so you can come back and fill it in.
   - `SUPPORT_CONTACT` – what customers see when their trial ends (your GCash/Messenger).
4. In the function dropdown, choose **setupKwentoink** → **Run** → allow permissions.
   This creates a folder **Kwentoink (customers)** in your Drive, with your **Developer Tracker** sheet inside.
5. Click **Deploy → New deployment → Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy** and **copy the Web app URL** (ends in `/exec`).
6. Your **admin page** is that URL + `?admin=YOUR_KEY`.
   To see your key: run **showAdminKey** and check the log (View → Logs / Execution log).
   Bookmark the admin page on your phone and keep it private.

## Part 2 – App (GitHub Pages)

1. In GitHub, create a repo named `kwentoink` (Public).
2. Open **index.html**, find this line near the top of the script, and paste your Web app URL:
   ```
   const API_URL = 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
   ```
3. Upload **index.html, sw.js, manifest.json, logo.png** to the repo.
4. Repo **Settings → Pages** → Source: **Deploy from a branch** → `main` / root → Save.
5. Wait about a minute. Your app link is `https://yourname.github.io/kwentoink/`.
   Put this in `APP_URL` in Code.gs if you haven't yet, then **Deploy → Manage deployments → Edit → New version → Deploy**.

## Part 3 – Add a customer (from your phone)

1. Open your admin page → type the business name → **Create & get link**.
2. Tap **📋 Copy link** and paste it in Messenger, or **💬 Text it**.
3. The owner opens the link → install guide → setup questions → name → tutorial.
4. Staff can open the **same link** on their phones. They skip setup, enter their name, and get a short tutorial.

Admin buttons:
- **✅ Paid +1 month** – after they pay you
- **⏸ Pause / ▶️ Resume** – stop or restart access (data is kept)
- **🔁 New link** – the old link stops working on all phones; the data stays

New customers get a **14-day free trial** (change `TRIAL_DAYS` in Code.gs).

## Try it before deploying

Open `index.html?id=demo` (or any link while `API_URL` isn't filled in yet). Kwentoink runs in **DEMO** mode, with data saved only on that phone. That's handy for showing the business owner.

## When you update the app later

1. Upload the new **index.html** to GitHub.
2. In **sw.js**, change `kwentoink-v1` to `kwentoink-v2` (then v3, and so on) and upload it too, so phones pick up the new version.
3. If you change **Code.gs**: **Deploy → Manage deployments → Edit → New version**. The URL stays the same.
