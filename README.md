# stoffbergtroufees.com

Save the date for Ané & Dirk, 18 December 2027, Joubertsdal.

- Every invite has its own link: `https://stoffbergtroufees.com/g/<code>`
- On opening it, the guest types their name and surname. It must match a name on that invite in the guest list.
- Every name entered, every visit and every "Add to calendar" click lands in a Google Sheet. Download it as Excel any time (File > Download > Microsoft Excel).

## How it fits together

| Part | Where | What it does |
|---|---|---|
| The website | this repo, on GitHub Pages | `index.html`, `chapel.webp`, `config.js` |
| Guest list + log | Google Sheet "Stoffberg troufees" | tabs **Guests**, **Log**, **Directory** |
| The link between them | `apps-script/Code.gs`, pasted into the sheet | checks names, writes the log |

The guest list never goes into this repo, so the repo can stay public.

## 1. The Google Sheet (once)

1. Open the sheet "Stoffberg troufees" in your Drive (or create a new Google Sheet).
2. **Extensions > Apps Script**. Delete what's there, paste all of `apps-script/Code.gs`, click **Save**.
3. Reload the sheet. A **Troufees** menu appears. Click **Troufees > Set up tabs** and allow the permissions it asks for (it's your own script).
4. Fill in **Guests**, one row per invite:

   | Code | Invite name | Names (separate with ;) | Language | Personal link |
   |---|---|---|---|---|
   | *(leave empty)* | Jan & Marie | Jan van der Merwe; Marie van der Merwe | af | *(leave empty)* |

   Want one link per person? Put one name per row.

   **First names only for now?** That works too: with the personal link, the right first name is enough
   (a title like "Oom" or "Tannie" in front is ignored). The full name the guest types is saved in the
   Directory, so you collect surnames as people open their link. Once you put a full name in column C,
   that guest has to type the full name.
5. **Troufees > Make personal links**. Codes and links appear in columns A and E. Existing codes never change, so you can add guests later.

## 2. Connect the sheet to the website (once)

1. In Apps Script: **Deploy > New deployment > Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
2. Copy the **Web app URL** (ends in `/exec`).
3. In this repo, open `config.js` and paste it between the quotes:
   `window.TROUFEES_API = 'https://script.google.com/macros/s/.../exec';`

If you later change `Code.gs`, use **Deploy > Manage deployments > Edit > Version: New version**, so the URL stays the same.

## 3. GitHub Pages (once)

Repo **Settings > Pages**: Source **Deploy from a branch**, branch **main**, folder **/ (root)**.

## 4. The domain stoffbergtroufees.com (once)

1. Buy the domain (TransIP, Hostnet, Namecheap…).
2. At the domain provider, add these DNS records:

   | Type | Name | Value |
   |---|---|---|
   | A | @ | 185.199.108.153 |
   | A | @ | 185.199.109.153 |
   | A | @ | 185.199.110.153 |
   | A | @ | 185.199.111.153 |
   | CNAME | www | dirk461.github.io |

3. Repo **Settings > Pages > Custom domain**: `stoffbergtroufees.com` (the `CNAME` file already says this). When the check turns green, tick **Enforce HTTPS**. DNS can take up to a day.

## What you'll see in the sheet

- **Directory**: one row per invite that has been opened: who typed their name, first and last visit, number of visits, and whether they clicked **Add to calendar** (and when).
- **Scores**: the game leaderboard, best score per player. The page shows the top 10 as first name + initial.
- **Log**: every event, newest at the bottom: *Name entered*, *Name not matched*, *Opened* (a returning guest on the same phone), *Added to calendar*.

## Notes

- A guest types their name once per phone; after that the phone remembers it and later visits are logged as *Opened*. "Not you?" at the bottom lets someone else on the same phone enter their name.
- Names are matched without caring about capitals or accents. A middle name may be left out ("Anna Botha" opens "Anna Maria Botha").
- If someone opens the wrong link, or the plain homepage, and types a name from the guest list, they're sent to their own invite.
- On iPhone/Mac "Add to calendar" downloads a calendar file; on Android/Windows it opens Google Calendar.
- With `config.js` left empty, the site works in try-out mode: any full name is accepted and nothing is logged.
