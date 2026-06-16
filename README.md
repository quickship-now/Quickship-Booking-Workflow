# Quickship Booking Workflow

GitHub Pages version of the Quickship booking workflow.

## Files

- `index.html` is the GitHub Pages frontend.
- `config.js` stores the deployed Apps Script web-app URL used by the frontend.
- `Code.gs` is the Apps Script backend that reads/writes Google Sheets and sends email.
- `1index.html` and `1code.gs` are the original source files kept for backup.

## Deploy The Backend

1. Open your Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Replace the script code with the contents of `Code.gs`.
4. If your Apps Script project uses an HTML file, keep your existing `Index.html` there from the original project.
5. Click **Deploy > New deployment**.
6. Choose **Web app**.
7. Set **Execute as** to **Me**.
8. Set **Who has access** to **Anyone**.
9. Deploy and copy the web app URL.

## Connect GitHub Pages

Edit `config.js` and paste your Apps Script web app URL:

```js
window.WORKFLOW_API_URL = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec";
```

## Publish On GitHub Pages

1. Create a GitHub repository.
2. Upload these files to the repository root:
   - `index.html`
   - `config.js`
   - `Code.gs`
   - `README.md`
3. In GitHub, open **Settings > Pages**.
4. Under **Build and deployment**, choose **Deploy from a branch**.
5. Select your main branch and `/root`.
6. Save.

Your live site will be available at:

```text
https://YOUR-GITHUB-USERNAME.github.io/YOUR-REPOSITORY-NAME/
```

## Notes

Until `config.js` has a real Apps Script web app URL, the frontend runs in demo mode with sample data.

The office credentials are stored in the Google Sheet by the Apps Script backend. Change the default users in `Code.gs` or update the `WorkflowUsers` sheet after the first run.
