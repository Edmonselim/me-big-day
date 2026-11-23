# M&E BIG DAY: one-time setup (about 15 minutes)

The app lives at **https://edmonselim.github.io/me-big-day/** and reads and writes
`_Wedding_Budget_Planner.xlsx` in Ed's OneDrive (`Edmon Drive/TOUTA/_BIG DAY/`).
No wedding data is stored in this repo. The code is public, but your numbers stay in OneDrive.

---

## 1. Turn on GitHub Pages (2 min)
1. Open the repo **Settings → Pages**.
2. Under **Build and deployment** choose **Deploy from a branch**. Pick branch `main`, folder `/ (root)`, and click **Save**.
3. After about a minute the site is live at `https://edmonselim.github.io/me-big-day/`.
   It will say *"Almost there! Paste your client ID…"* until step 2 is finished.

## 2. Register the app with Microsoft (8 min)
This gives the app permission to sign you in and open the workbook. It's free.

1. Go to **https://entra.microsoft.com** and sign in with your **Outlook** account.
   - If it says you need a directory or tenant, first sign up for a free account at
     **https://azure.microsoft.com/free** with the same Outlook address, then come back.
2. Go to **Applications → App registrations → New registration**.
   - **Name:** `M&E BIG DAY`
   - **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts*
   - **Redirect URI:** choose **Single-page application (SPA)** and enter
     `https://edmonselim.github.io/me-big-day/`, including the trailing slash.
   - Click **Register**.
3. On the overview page, copy the **Application (client) ID**. It looks like `1a2b3c4d-…`.
4. Open **API permissions → Add a permission → Microsoft Graph → Delegated permissions** and tick:
   `Files.ReadWrite.All`, `User.Read` and `offline_access`. Then click **Add permissions**.
   No admin consent is needed. You'll each approve the permissions the first time you sign in.

## 3. Paste the client ID (1 min)
Edit **`config.js`** in this repo on GitHub: click the pencil icon, replace
`PASTE-YOUR-CLIENT-ID-HERE` with your ID, and commit. Pages redeploys in about a minute.

## 4. Install on Ed's iPhone
1. Open **Safari** and go to `https://edmonselim.github.io/me-big-day/`.
2. Tap **Share → Add to Home Screen → Add**. The M&E BIG DAY icon appears.
3. Open it from the home screen, tap **Sign in with Microsoft**, sign in, and accept the permissions.
4. Tap **"It's in my OneDrive (Ed)"**. The app finds the workbook and loads it.

## 5. Share with Touta
1. **Touta:** if your Gmail address isn't a Microsoft account yet, create one at
   **https://account.microsoft.com** and choose "use your email address" to use your Gmail.
2. **Ed:** in OneDrive (web or Windows), right-click `_Wedding_Budget_Planner.xlsx`,
   choose **Share**, enter Touta's Gmail and set it to **Can edit**. Then copy the share link,
   or the **partner code** from the app's **More → Settings** screen, and send it to her.
3. **Touta:** install the app the same way as step 4, then sign in with your Microsoft account
   that uses your Gmail address. On the Connect screen, paste the link or partner code and tap **Connect shared file**.

---

## How syncing works
- The app downloads the latest workbook when it opens, when you tap ↻, and every minute while it's open.
- Your edits show instantly and are saved to OneDrive a second or two later. Offline edits wait
  on the phone (the pill says *"N unsaved"*) and are saved automatically once you're back online.
- Saving writes only the cells you changed, so your formatting, dropdowns, colors and formulas stay as they are.
  Excel recalculates everything the next time the file is opened.
- If you both edit at the same moment, the app merges the changes. If you both edit the *same cell*, the last save wins.
  New rows never overwrite each other.
- **Close the file in desktop Excel** before saving from the app. While Excel has it open, OneDrive may lock it,
  and the app will say so and retry automatically.

## Troubleshooting
| Message | Fix |
|---|---|
| "Sign in needed" | Tap **Sign in**. Microsoft sign-ins for web apps expire after about a day of inactivity. |
| "Open in Excel" | Close the workbook in desktop Excel. The app retries every 20 seconds. |
| "Can't find the workbook" | It was moved or renamed. Put it back, or update `ownerFilePath` in `config.js`. |
| Touta can't connect | Check the file is shared with the **same** Microsoft account she signed in with. |

To undo a mistake, the original file is saved next to the workbook as
`_Wedding_Budget_Planner (backup 2026-11-23).xlsx`, and OneDrive keeps **Version history** (right-click the file).
