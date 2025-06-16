# Everyone's Answer Board

This project contains the source code for a Google Apps Script (GAS) application.
It uses [clasp](https://github.com/google/clasp) to push and pull code to and from
your GAS project.

## Setup

1. Install Node.js (v20 or later).
2. Install dependencies:
   ```bash
   npm install
   ```
3. Authenticate clasp:
   ```bash
   npx clasp login
   ```
4. Set your Apps Script project ID in `.clasp.json` by replacing `ENTER_YOUR_SCRIPT_ID`.

## Usage

- `npm run push` – Upload code in `src/` to the GAS project.
- `npm run pull` – Download the latest code from GAS.
- `npm run open` – Open the GAS project in your browser.
- `npm run update-url` – Save the latest web app URL to script properties.

## Running tests
Run `npm install` first to install Jest and other dev dependencies, then run `npm test`.

## Project structure

All script files live directly under the `src` folder. **Do not create any
subdirectories under `src`.**

## Highlighting answers

Add a column named `Highlight` to your spreadsheet (using a checkbox or TRUE/FALSE values).
Rows marked as `TRUE` will appear with a yellow border on the board.
Administrators see a star button on each answer card allowing them to toggle this flag.

## Admin access

1. Open the sheet selector sidebar from the spreadsheet menu.
2. Enter comma-separated administrator emails in the **管理者メールアドレス** field and click **保存**.
3. Users listed here automatically see admin features (reaction counts, names and highlight controls) when viewing the board.

### Student view for administrators

Administrators normally see the same student interface. To display admin
features on demand, use the **管理者として開く** button in the sheet selector
sidebar. The button opens the board with `mode=admin` automatically added to the
URL. You can also manually append `mode=admin` when sharing the link or
projecting the board.

## Continuous Integration

A GitHub Actions workflow automatically pushes code to the Apps Script project whenever files in `src/` change. To enable it, add your service account credentials JSON as the `CLASP_CREDENTIALS` secret in the repository settings.
