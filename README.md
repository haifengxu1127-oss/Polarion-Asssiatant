# Polarion Requirement Assistant — static web edition

This folder is a standalone static website for GitHub Pages. It has no server, build step, account credentials, or SharePoint API connection.

## What it does

1. Import a SharePoint link-list `.xlsx` or `.csv` with `File Title` and `Hyperlink` columns. A clickable hyperlink on the title cell is also supported.
2. Select a locally synced folder containing PDF, DOCX, TXT, or Markdown documents. The website reads file **names** for matching; it does not read document contents.
3. Run AI analysis in the existing VS Code extension, export its Excel result, and import that workbook into this page. Priority and Scope values are matched by file title.
4. Review and edit the rows, then export a four-column `.xlsx`: `File Title`, `Hyperlink`, `Priority`, `Scope`.

Inputs are handled in browser memory. This website contains no code that uploads selected files to a server. Original SharePoint links must come from your own verified link list. Rows with the same filename in different folders require manual review.

## Publish to GitHub Pages

Upload **only the contents of this folder** to the root of a new GitHub repository: `index.html`, `app.js`, `styles.css`, `icon.png`, and `vendor/`. Do not upload the parent project folder, VSIX packages, customer documents, or real link lists.

In the repository, open **Settings → Pages → Build and deployment**. Choose **Deploy from a branch**, select `main` and `/(root)`, then save. GitHub will provide a URL similar to `https://USERNAME.github.io/REPOSITORY/`.

The website uses relative asset paths, so it works at a GitHub Pages project URL. It can also be served locally with any static HTTP server. GitHub Pages websites are publicly reachable, including some sites whose source repository is private. Check your company's publishing policy before using the hosted page with company data.

## Important limit

GitHub Pages cannot call VS Code's GitHub Copilot Language Model API. This website imports the **results** produced by the existing VS Code extension. A standalone browser-based AI analyzer would require a separate company-approved model integration. Likewise, direct SharePoint reading would require company-approved Microsoft application registration and permissions; this static edition uses an imported link list.

The exported workbook's Polarion field mapping still needs to be checked in your company's Polarion environment.

## Third-party code

`vendor/xlsx.full.min.js` is SheetJS Community Edition 0.18.5, licensed under Apache 2.0. The license text is in `vendor/LICENSE-Apache-2.0.txt`.
