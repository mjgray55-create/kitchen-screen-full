# Kitchen screen: Sheet edition

Turns an old tablet into a kitchen screen: the week's menu, a meals list, the shopping list and kitchen timers,
plus fridge and freezer stock, a recipe book, and seasonal and holiday looks, depending on the package.

In the **Sheet edition**, everything is kept in the household's own **Google Sheet**, using their own Google account.
That means:

- the shopping list (and everything else) can be checked and changed **from a phone** in the Google Sheets app,
- it can be edited on a computer,
- Google keeps the version history, so there's nothing to back up by hand.

(The other edition, **Lite**, needs no Google account and keeps everything on the tablet itself. It lives in its own
repository, `kitchen-screen-lite`.)

## Packages

The package is picked on the tablet during first-time setup (Settings → *What this kitchen has*). The choice is
saved on the sheet's **Settings** tab.

| Package  | What's included |
|----------|-----------------|
| Basic    | Weekly menu (lunch, dinner, who's cooking, notes), meals list the menu picks from, shopping list, kitchen timers |
| Expanded | Basic, plus fridge and freezer stock with leftovers ("Anything left from last night?", portions, and "Eaten" on the menu counts down the fridge or freezer), and the recipe book with tap-to-start timers |
| Primo    | Expanded, plus looks: seasons (autumn leaves, winter snow, spring blossom, summer fireflies), US holidays, birthdays (people, dogs, cats) and anniversaries with a two-week heads-up and confetti on the day. Falling decorations can be set to all year, only on special days, or off. Tap the clock 5 times to preview every look. |

US holidays: Halloween (Oct 20–31), Thanksgiving (Thursday–Sunday), Christmas (Dec 1–25, with lights), New Year,
Valentine's Day, St Patrick's Day, Easter (flowers), Fourth of July.

## What's in this repository

| File | What it is |
|------|------------|
| `Code.gs` | The Apps Script code that reads and writes the sheet. It also creates any missing tabs by itself. |
| `Index.html` | The screen itself. |
| `test/` | A way to try the screen on a computer without Google (see *Testing* at the bottom). |

---

## Part 1: Make the template (once)

You only do this once. The result is a blank "Kitchen" sheet with the screen built in, which customers copy.

1. Go to **sheets.new** (signed in to your business Google account). Name the sheet **Kitchen**.
2. In the sheet: **Extensions → Apps Script**. A code editor opens in a new tab.
3. At the top left, click **Untitled project** and rename it **Kitchen screen**.
4. **Code.gs**: select everything in the editor, delete it, and paste in the whole of `Code.gs` from this repository.
5. **Index.html**: click the **+** next to *Files* → **HTML**, type `Index` (capital I, no `.html`), press Enter.
   Select everything in the new file, delete it, and paste in the whole of `Index.html`.
6. Click the **Save** icon (the floppy disk).
7. Make the tabs: in the toolbar, choose **setUpTabs** in the function dropdown, then click **Run**.
   - Google asks for permission the first time. Click **Review permissions**, pick your account, then
     **Advanced → Go to Kitchen screen (unsafe) → Allow**. ("Unsafe" only means Google hasn't reviewed this
     small private script. It's your own code.)
8. Go back to the sheet tab. It now has these tabs: Menu, Meals, Shopping, Fridge, Freezer, Consumed, Recipes,
   Birthdays and Settings. They're all empty except for headings and the starting settings.
9. **Time zone**: in Apps Script, click the **gear (Project Settings)** and set the time zone your customers are in.
   Copies keep this setting.
10. **Share the template**: in the sheet, click **Share**. Under *General access*, choose **Anyone with the link**,
    set to **Viewer**, then **Copy link**.
11. Turn that link into a "make a copy" link: in the link, replace everything from `/edit` onward with `/copy`.

    `https://docs.google.com/spreadsheets/d/LONG-ID/edit?usp=sharing` becomes
    `https://docs.google.com/spreadsheets/d/LONG-ID/copy`

    Anyone who opens the `/copy` link gets their own private copy, with the screen code included.

Don't type anything into the template itself. Keep it blank.

## Part 2: Set up a customer (each time)

Do this with the customer, signed in to **their** Google account (on a computer or on the tablet).

1. **Make their copy.** Open your `/copy` link and click **Make a copy**. Google mentions that the attached
   Apps Script will be copied too. That's the screen. Rename the copy if you like (for example "Kitchen").
2. **Open the script.** In their copy: **Extensions → Apps Script**.
3. **Deploy the screen.**
   1. Click **Deploy → New deployment**.
   2. Click the **gear** next to *Select type* and choose **Web app**.
   3. Description: `Kitchen screen`.
   4. *Execute as*: **Me**.
   5. *Who has access*: **Only myself**.
   6. Click **Deploy**. When asked, **Authorize access**: pick their account, then
      **Advanced → Go to Kitchen screen (unsafe) → Allow**.
   7. Copy the **Web app URL** (it ends in `/exec`). This is the screen's address.
4. **Open it on the tablet.** On the tablet, sign in to Chrome with the **same Google account**, open the Web app URL,
   then **⋮ → Add to Home screen** so it opens with one tap.
5. **First-time setup** appears on the screen. Pick the package they bought, then fill in who cooks, their freezers
   and (for Primo) birthdays. Tap **Done — start using it**.
6. **Their phone.** Install the **Google Sheets** app, sign in, and open the Kitchen sheet. The **Shopping** tab is the
   shopping list. Anything added there shows on the tablet within a minute. To give someone else in the household
   the shopping list, **Share** the sheet with them as an **Editor**.

### Things worth telling the customer

- Counts only change when someone taps (for example *Eaten* on the menu, or *Take out* on a freezer item).
  Nothing counts down by itself.
- Everything is in their sheet. **File → Version history** in the sheet can undo mistakes.
- On the **Freezer** tab, each heading is on its own row with its items underneath. If they add things by hand,
  keep that layout.
- Birthdays can be added on the tablet (Settings) or on the **Birthdays** tab: name, date like `March 4`, and
  Person, Dog, Cat or Anniversary.
- The **Settings** tab can be changed on the sheet too. The tablet picks up changes within a minute.

## Part 3: Updating a customer's screen later

When there's a new version of `Code.gs` or `Index.html`:

1. In their sheet: **Extensions → Apps Script**. Paste the new file(s) over the old ones and **Save**.
2. **Deploy → Manage deployments** → the pencil icon ✏️ → *Version*: **New version** → **Deploy**.

Never use **New deployment** for an update. That makes a new address and the tablet would keep showing the old one.

If the update adds a new tab, it's created automatically the next time the screen opens. In the sheet there's also
a **Kitchen screen → Set up or repair the tabs** menu. It only adds missing tabs and never changes existing ones.

---

## Testing (for development)

`test/` runs the real `Code.gs` and `Index.html` in a browser against a pretend Google Sheet, so changes can be
checked without deploying. In PowerShell, from this folder:

```
powershell -ExecutionPolicy Bypass -File test\build.ps1
powershell -ExecutionPolicy Bypass -File test\serve.ps1
```

Then open `http://localhost:8765/test.html`, and check at 1334×800 (landscape) and 800×1334 (portrait) with no
console errors. In the browser console, `__BOOK__.getSheetByName('Freezer').dump()` shows what a tab would contain.
