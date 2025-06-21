# StudyQuest - みんなの回答ボード

This repository contains the source code of a Google Apps Script project. It publishes a web app that displays students' answers stored in a spreadsheet. Users can like answers and filter the board by class. Teachers can manage which sheet is currently published from a custom menu within the spreadsheet.

## Setup

1. Install **Node.js 20** or later.
2. Install dependencies:
   ```bash
   npm install
   ```
3. Sign in to Google with clasp:
   ```bash
   npx clasp login
   ```
4. Edit `.clasp.json` and set your Apps Script project ID.

## Usage

- `npm run push` – Upload the contents of `src/` to your Apps Script project.
- `npm run pull` – Download the latest code from Apps Script.
- `npm run open` – Open the Apps Script project in a browser.
- `npm run update-url` – Store the latest Web App URL in script properties.

## Spreadsheet structure

The sheet you publish must contain the following columns:

- **メールアドレス** – address of the user who submitted the answer
- **クラスを選択してください。** – class name used for filtering
- **これまでの学んだことや、経験したことから、根からとり入れた水は、植物のからだのどこを通るのか予想しましょう。** – question text displayed on the board
- **予想したわけを書きましょう。** – explanation shown when opening a card
- **いいね！** – comma separated list of users who liked the answer

A separate sheet named `roster` should contain roster information with the columns `学年`, `組`, `番号`, `姓`, `名`, `Googleアカウント`, `ニックネーム`. These names are shown on the board instead of raw email addresses. If your roster sheet has a different name, set the `ROSTER_SHEET_NAME` script property to override the default.

## Managing the board

Opening the spreadsheet adds an **アプリ管理** menu. From here you can:

1. **管理パネルを開く** – choose which sheet to publish or unpublish.
   Use the "Config" section to map the question, answer and name columns. The
   selected settings are saved automatically when publishing a sheet for the
   first time.
2. **名簿キャッシュをリセット** – refresh the cached roster information.

When unpublished, visiting the Web App URL shows a message that the board is closed. Once published, the board is available and updates automatically every 15 seconds.

### Admin interface

Administrators see a toggle button that allows switching between viewing and management modes. The list of administrator email addresses is set in the `ADMIN_EMAILS` script property.

### Setting `ADMIN_EMAILS`

1. Open the Apps Script **Project Settings** and expand **Script Properties**.
2. Click **Add script property** and set the name to `ADMIN_EMAILS`.
3. Enter a comma‑separated list of addresses and save.

You can also set the value from the console:

```javascript
PropertiesService.getScriptProperties()
  .setProperty('ADMIN_EMAILS', 'teacher1@example.com,teacher2@example.com');
```

### Setting `ROSTER_SHEET_NAME`

1. Open the Apps Script **Project Settings** and expand **Script Properties**.
2. Add a property named `ROSTER_SHEET_NAME` and set its value to your roster sheet's name.
3. Leave it blank or remove it to use the default `roster`.

## Front‑end features

- Answers are displayed in a responsive grid. A slider allows changing the number of columns.
- You can filter answers by class.
- Clicking a card opens a modal with the full text and a button to like the answer.
- Like counts update instantly and contribute to the sorting order.

## Project structure

All Apps Script and HTML files must live directly in the `src` directory. **Do not create subdirectories under `src`.**

## Testing

Install dependencies first and then run:
```bash
npm test
```

