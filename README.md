# huanyang
Incorporating physical climate risks into banks' credit risk models based on BIS working papers No 1274
# Physical Climate Risk Credit Capital Calculator

This folder is a static GitHub Pages version of the calculator.

## Files

- `index.html` - page structure
- `app.css` - styling
- `app.js` - all model logic and browser-side Excel processing
- `.nojekyll` - tells GitHub Pages to serve files as-is

## Deployment

Option A: publish this folder as the repository root.

1. Create a GitHub repository.
2. Upload the contents of `github_pages_app`.
3. Go to `Settings -> Pages`.
4. Select `Deploy from a branch`, choose `main` and `/root`.
5. Open the generated GitHub Pages URL.

Option B: use a `docs` folder.

1. Rename or copy `github_pages_app` to `docs`.
2. Push the repository.
3. Go to `Settings -> Pages`.
4. Select `Deploy from a branch`, choose `main` and `/docs`.

## Excel support

Excel upload/download is handled in the user's browser through SheetJS loaded from:

```text
https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
```

If a user's network blocks CDN access, single-record calculation still works, but Excel upload/download will not.

## Privacy

All calculations are performed locally in the user's browser. Uploaded Excel files are not sent to a server by this static version.

