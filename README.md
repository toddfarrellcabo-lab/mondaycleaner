# Monday Cleaner

A browser-based utility that turns a Monday.com **Creative Requests** Excel export into the working workbook used by the Creative Services team.

## What the browser version does

The app reads the selected `.xlsx` file **locally in the user's browser** and generates a new workbook containing:

- **Status** — prioritized dashboard with worksheet navigation.
- **Updates** — Monday updates/replies mapped back to the request, with technical IDs hidden.
- **Attachments** — clickable references to Monday attachment URLs exported in the source workbook.
- **One task worksheet per parent request** — with linked management fields, request details, and navigation back to Status.
- Status dropdown: `New`, `Working`, `Stuck`, `On Hold`, `Out for Review`, `Complete`.
- Assigned dropdown: `Todd`, `Travis`, `Brady`.
- Arial 8 pt body text; Arial 9 pt bold headers.
- Current-date output filename: `MM-DD-YY_CR_ToDos.xlsx`.
- Light tab colors and internal workbook navigation links.
- No conditional formatting.

The file is not posted to an application server. JavaScript reads it using the browser's File API and generates the output workbook in browser memory.

## Repository files

```text
Monday Cleaner/
├── index.html
├── app.js
├── styles.css
├── README.md
├── .nojekyll
└── assets/
    ├── monday-logo.png
    └── Exp-Mon.gif
```

The Excel processing library is **ExcelJS 4.4.0**, loaded from jsDelivr by `index.html`.

## Publish with GitHub Pages

1. Create a repository for **Monday Cleaner**.
2. Upload the files from this package to the root of the repository.
3. Commit/push them to your default branch.
4. In the repository's GitHub Pages settings, publish the site from the root of that branch.
5. GitHub will provide the site URL. Share that URL with the team.

If your organization restricts GitHub Pages, use the approved internal/static hosting option instead. This project does not require a backend server.

## Normal use

1. Export the Creative Requests board from Monday.com as Excel.
2. Open Monday Cleaner site.
3. Drag the `.xlsx` export onto the page or choose it.
4. Click **Clean Monday Export**.
5. Download the generated `MM-DD-YY_CR_ToDos.xlsx`.

For the **Updates** worksheet to populate, the Monday export needs to include the `updates` tab.

## Privacy / security model

The workbook content is processed locally in the browser. The application does not contain code that uploads the selected workbook to GitHub or to an application server.

The page does make a normal network request to jsDelivr to load ExcelJS. The workbook itself is not passed to that request.

If company policy requires all dependencies to be self-hosted, download the approved ExcelJS browser bundle into the repository and change the script reference in `index.html` from the jsDelivr URL to the local file.

## Attachments

The app preserves `protected_static` attachment URLs found in Monday's exported **Assets** and subitem **Files** fields and makes them clickable in the generated workbook.

It does **not** automatically download those files. Protected Monday assets may require the user's authenticated Monday session and can also be subject to browser CORS restrictions. That should be treated as a separate feature.

## Updating the tool

Once GitHub Pages is live, updates are centralized: modify the repository and publish the changes. Team members do not need to replace BAT/Python files or install anything.

## Browser support

Use a current desktop version of Microsoft Edge or Google Chrome. Large Monday exports may take several seconds to process because the entire Excel workbook is handled in-browser.


## Page branding / QuickStart

The page uses the included `assets/monday-logo.png` in the header and `assets/Exp-Mon.gif`
for the looping QuickStart area. The top-right **Jump to QuickStart Guide** link scrolls
directly to that section.

The small Sparklight logo in the footer currently uses Sparklight's official website image URL.
If you prefer a fully self-contained repo, add an approved Sparklight logo to `assets/` and
change the footer image source in `index.html` to that local file.

Footer line: `a Problematic solution 2026.`


## Header / responsive QuickStart

The page header now follows the compact utility pattern used by the Creative Services DCO Builder:
Sparklight is the owning brand at left, Monday Cleaner is the main title, and a small Monday.com
mark sits beside the tool name to identify the source platform.

The title uses `Effra` when it is installed on the user's machine and automatically falls back to Arial.

On desktop, `assets/Exp-Mon.gif` renders normally and loops as an animated GIF.
On screens 760px wide or smaller, the embedded GIF is hidden by default and the page shows an
**Open QuickStart animation** link that opens the GIF directly in a new tab.


## QuickStart GIF file

The packaged source image that had been named `Exp-Mon.gif` was detected as a single-frame PNG,
so it could not animate. It is now stored as:

`assets/Exp-Mon-fallback.png`

To enable the looping guide, place the real animated GIF in the repo as:

`assets/Exp-Mon.gif`

No HTML change is required. The page attempts to load the GIF first and falls back to the static
preview only if the GIF is missing.

The top-right **Jump to QuickStart Guide** link now targets the animation container directly.


## Header mascot

The current header uses:

`assets/monday-cleaner-mascot.png`

The title, Monday.com logo, Sparklight logo, QuickStart link, and tool controls remain live HTML/CSS.
The browser/page title is `Monday Cleaner`.
