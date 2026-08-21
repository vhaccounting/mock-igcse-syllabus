# Student Register — private dashboard

The full teacher view: names, marks, grades, mock trajectory, and WhatsApp
click-to-chat links for student, father and mother. Not published anywhere.

## Why it is built this way

The dashboard used to read the spreadsheet from the browser. That only works
if the spreadsheet is shared with "anyone with the link" — and once it is,
anyone who knows the spreadsheet ID can download every tab, including the
contact columns, with no sign-in at all.

Here the browser never contacts Google Sheets. The Apps Script server reads
the sheet with its own credentials and returns rows only to a signed-in
account on `ALLOWLIST`. That is what allows the spreadsheet's own sharing to
be set to **Restricted**.

## Setup

1. Open the marksheet spreadsheet → **Extensions → Apps Script**.
2. Create two files in the script project:
   - `Code.gs` ← paste from `Code.gs`
   - `Index.html` ← paste from `Index.html` (File → New → HTML, name it `Index`)
3. Edit `ALLOWLIST` in `Code.gs` to list the Google accounts allowed in.
4. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone with a Google account**
5. Open the `/exec` URL, authorise once, and bookmark it.
6. Go back to the spreadsheet → **Share → Restricted**.

Step 6 is the one that closes the hole. Until then the sheet stays readable
by anyone holding its ID.

"Anyone with a Google account" only decides who Google will sign in. The
`ALLOWLIST` check decides who receives data; everyone else gets a refusal
page and `getRows` throws before touching the spreadsheet.

## Adding a teacher

Add their address to `ALLOWLIST`, save, then **Deploy → Manage deployments →
Edit → Version: New version**. Without a new version the running deployment
keeps the old code.

## Access log

Every successful read appends to an `AccessLog` tab: timestamp, account, tab.
Delete the tab to reset it; it is recreated on the next read.
