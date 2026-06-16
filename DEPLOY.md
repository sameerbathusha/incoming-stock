# Putting your app online (about 10 minutes, no installs)

Your app is the four files in this `webapp` folder:
`index.html`, `app.js`, `mapping.js`, `config.js`. You only ever edit `config.js`.

## Step 1 — Get your two Supabase values
In Supabase: **Project Settings** (gear, bottom-left) → **Data API** for the
**Project URL**, and **API Keys** for the **anon / public** key (sometimes
called "publishable"). The anon key is safe to ship — your login and the
database rules do the protecting.

## Step 2 — Put them in config.js
Open `config.js` in any text editor (Notepad/TextEdit are fine). Replace the two
placeholder values with your Project URL and anon key. Save.

## Step 3 — Create your login
A database has no users yet, so make one:
1. Supabase → **Authentication** → **Users** → **Add user** → enter your email
   and a password, and tick **Auto Confirm User**. Create it.
2. Make yourself an admin: Supabase → **SQL Editor**, run
   (replace with your email):
   ```sql
   update public.profiles set role = 'admin' where email = 'you@company.com';
   ```
   (Roles: `admin` = everything, `operations` = view + upload, `sales` = view only.)

## Step 4 — Put it online (drag and drop)
1. Go to **app.netlify.com/drop** (sign up free if asked).
2. Drag this whole `webapp` folder onto the page.
3. In a few seconds you get a live link like `https://something.netlify.app`.
   That's your app. Bookmark it; share it with your team later.

## Step 5 — Use it
Open the link, sign in with the user from Step 3. Go to **Upload sheet**, pick
the matching template (AT / Momax / Tangem), choose the Excel file, **Read
file**, check the preview, then **Import**. Open **Dashboard** to see it light up.

## Adding teammates later
For each person: Supabase → Authentication → Add user (email + password), then
set their role in SQL (`operations` for the ops team, `sales` for sales). They
sign in at the same link.

## Notes
- Test using the Netlify link, not by double-clicking the file — browsers block
  parts of the app when opened directly from your computer.
- To change anything later, edit the file and drag the folder onto Netlify Drop
  again (or set up a Netlify account to update in place).
- Updating `config.js` after deploy means re-dragging the folder.
