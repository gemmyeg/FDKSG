# Future Designers - Management System

Custom management system for "Future Designers" (Egypt-based zipper/textile trading business), built entirely on Google Apps Script.

## Structure

- **`Code.gs`** - Backend: all server-side logic (`doGet`, `api_*` functions), bound to the "Future Designers - Master Database" Google Sheet.
- **`Index.html`** - Frontend: the entire single-page web app UI (HTML/CSS/JS), served by `doGet()` via `HtmlService`.

## Live deployment

The app is deployed as a Google Apps Script Web App. The live URL is documented as a comment at the top of `Index.html`, and is also:

```
https://script.google.com/macros/s/AKfycby80hznYPi8YiTkAFsbBqvY1F9c3TEf11A5GG5h9y3211WepB7csvKVGkZOUDEbpknp/exec
```

## Notes

- This is a snapshot of the code as of the last save in the Apps Script editor (2026-09-15). It is **not** connected to a build/deploy pipeline - to publish changes, copy the files back into the Apps Script editor (script.google.com) and deploy a new version from there.
- The Apps Script project manifest (`appsscript.json`) is hidden by default in the editor and is not included here. It can be added later via Project Settings → "Show `appsscript.json` manifest file in editor".
- The backend is bound to a specific Google Sheet (`MASTER_DB_ID` constant in `Code.gs`) which holds all business data (Customers, Suppliers, Orders, Owners, Assets, Stock, etc.) - this repo contains only the application code, no business data.
