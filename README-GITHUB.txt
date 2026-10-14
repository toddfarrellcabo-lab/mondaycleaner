MONDAY CLEANER - UNIVERSAL UPDATE

GITHUB UPLOAD
1. Back up your existing repository or make a new branch.
2. Upload index.html and the entire "Monday Cleaner_files" folder to the ROOT of your GitHub Pages repository, replacing matching files.
3. Commit the changes. GitHub Pages should serve index.html automatically.
4. Hard-refresh your website (Ctrl+Shift+R).

MODES
Auto: Uses the original specialized Creative Requests converter if its expected headers are present; otherwise uses Universal mode.
Creative Requests: Forces the original specialized converter and shows an error if the columns are missing.
Universal: Detects header rows on each worksheet, keeps all nonempty data rows, formats headers, freezes top row, enables filters, and allows worksheet preview.

LIMITATIONS
Universal mode is a clean flat-table export, not a perfect reproduction of Monday subitem hierarchy or attachments. It preserves cell text/links represented as values but does not preserve original Excel formatting, merged cells, images, or formulas as formulas. It detects one main header per worksheet and may need adjustments for unusual multi-table layouts. Preview shows first 100 rows per sheet.

TESTING
JavaScript syntax was checked. Automated browser test was blocked by the runtime browser security policy. Test the supplied Q426 example in GitHub Pages before relying on the output.

All processing occurs in the browser. No server-side upload.
