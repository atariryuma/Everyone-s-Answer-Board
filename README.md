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

A separate sheet named `sheet 1` should contain roster information with the columns `姓`, `名`, `ニックネーム` and `Googleアカウント`. These names are shown on the board instead of raw email addresses.

## Managing the board

Opening the spreadsheet adds an **アプリ管理** menu. From here you can:

1. **管理パネルを開く** – choose which sheet to publish or unpublish.
2. **名簿キャッシュをリセット** – refresh the cached roster information.

When unpublished, visiting the Web App URL shows a message that the board is closed. Once published, the board is available and updates automatically every 15 seconds.

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

