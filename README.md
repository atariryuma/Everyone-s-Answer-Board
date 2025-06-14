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
After installing dependencies (`npm install`), run `npm test` to execute Jest tests.

## Project structure

All script files live directly under the `src` folder. **Do not create any
subdirectories under `src`.**

## Highlighting answers

Add a column named `Highlight` to your spreadsheet (using a checkbox or TRUE/FALSE values).
Rows marked as `TRUE` will appear with a yellow border on the board.
When the board is opened with the `?admin=1` query parameter,
a star button is shown on each answer card to toggle this flag.
Use the "Open as administrator" link in the sheet selector sidebar to launch the board in this mode.

## Admin access

1. Open the sheet selector sidebar from the spreadsheet menu.
2. Enter comma-separated administrator emails in the **管理者メールアドレス** field and click **保存**.
3. Only these users will see the "管理者として開く" link and can append `?admin=1` to access admin features.

## Additional admin tools

- **Opinion groups** – The *意見グループ* link groups similar answers together. The `groupSimilarOpinions` function uses Gemini API to analyze all opinions and reasons and returns a summary of majority and minority views.


## Continuous Integration

A GitHub Actions workflow automatically pushes code to the Apps Script project whenever files in `src/` change. To enable it, add your service account credentials JSON as the `CLASP_CREDENTIALS` secret in the repository settings.
