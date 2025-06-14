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

## Running tests
After installing dependencies (`npm install`), run `npm test` to execute Jest tests.

## Project structure

All script files live directly under the `src` folder. **Do not create any
subdirectories under `src`.**

## Highlighting answers

Add a column named `Highlight` to your spreadsheet (using a checkbox or TRUE/FALSE values).
Rows marked as `TRUE` will appear with a yellow border on the board.
When viewing the board as an administrator (your Google account name contains `t`),
a star button is shown on each answer card to toggle this flag.


## Continuous Integration

A GitHub Actions workflow automatically pushes code to the Apps Script project whenever files in `src/` change. To enable it, add your service account credentials JSON as the `CLASP_CREDENTIALS` secret in the repository settings.
