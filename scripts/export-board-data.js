/**
 * 一回限りのデータエクスポート用スクリプト
 *
 * 現在ボードにひもづく3つのスプレッドシートを Sheets API 経由で取得し、
 * 個人情報保護のため `name` 列を除外したうえで CSV / JSON として書き出す。
 *
 * Usage:
 *   node scripts/export-board-data.js <outputDir>
 */
const fs = require('fs');
const path = require('path');
const https = require('https');
const { getAccessToken } = require('./lib/gas-auth');

const SHEETS = [
  {
    profile: '本時の議論',
    formTitle: 'ポスターをどうする？（道徳5/14 本時）',
    spreadsheetId: '1S0L6724xlKq9-br1EaGoxMfgU-Rj753HbU4sDeAPfhY',
    sheetName: 'フォームの回答 1',
    nameColumnIndex1Based: 3
  },
  {
    profile: '導入アンケート',
    formTitle: 'AIの答えそのまま使ったことある？',
    spreadsheetId: '1STTRp3N4pvZ86JUjyp04yBwlrintWchc8oJdvyh0yqY',
    sheetName: 'フォームの回答 1',
    nameColumnIndex1Based: 3
  },
  {
    profile: '振り返り',
    formTitle: '【振り返り】今日の道徳の学習で学んだこと（道徳5/14 振り返り）',
    spreadsheetId: '15WpPQpGB6LDkJhZpLUp2Bdx2cuI2U40V5ttH_XHVt38',
    sheetName: 'フォームの回答 1',
    nameColumnIndex1Based: 3
  }
];

function getRaw(url, token) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers: { Authorization: `Bearer ${token}` } }, (res) => {
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => {
        const body = Buffer.concat(chunks).toString('utf8');
        if (res.statusCode >= 400) {
          reject(new Error(`HTTP ${res.statusCode}: ${body.slice(0, 400)}`));
          return;
        }
        resolve(body);
      });
    }).on('error', reject);
  });
}

function getJSON(url, token) {
  return getRaw(url, token).then((body) => {
    try { return JSON.parse(body); }
    catch (e) { throw new Error(`Invalid JSON: ${body.slice(0, 400)}`); }
  });
}

function parseCSV(text) {
  const rows = [];
  let row = [];
  let cell = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { cell += '"'; i++; }
        else { inQuotes = false; }
      } else { cell += c; }
    } else {
      if (c === '"') { inQuotes = true; }
      else if (c === ',') { row.push(cell); cell = ''; }
      else if (c === '\n') { row.push(cell); cell = ''; rows.push(row); row = []; }
      else if (c === '\r') { /* skip */ }
      else { cell += c; }
    }
  }
  if (cell.length > 0 || row.length > 0) { row.push(cell); rows.push(row); }
  return rows;
}

function toCSV(rows) {
  const escape = (v) => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    if (s.includes(',') || s.includes('"') || s.includes('\n')) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };
  return rows.map((r) => r.map(escape).join(',')).join('\n');
}

async function main() {
  const outDir = process.argv[2] || '/tmp/board-export';
  fs.mkdirSync(outDir, { recursive: true });

  console.log('Getting access token...');
  const token = getAccessToken();
  console.log('Token OK\n');

  const summary = [];

  for (const s of SHEETS) {
    console.log(`=== ${s.profile} (${s.formTitle}) ===`);
    const url = `https://www.googleapis.com/drive/v3/files/${s.spreadsheetId}/export?mimeType=text/csv`;
    const csvText = await getRaw(url, token);
    const values = parseCSV(csvText);
    if (values.length === 0) {
      console.log('  (empty)\n');
      summary.push({ profile: s.profile, rowCount: 0, columnCount: 0 });
      continue;
    }

    const dropIdx0 = s.nameColumnIndex1Based - 1;
    const stripped = values.map((row) => row.filter((_, i) => i !== dropIdx0));

    const dataRowCount = stripped.length - 1;
    const colCount = stripped[0].length;
    console.log(`  rows (incl header): ${stripped.length} / data rows: ${dataRowCount} / cols (after strip): ${colCount}`);
    console.log(`  original headers: ${values[0].join(' | ')}`);
    console.log(`  stripped headers: ${stripped[0].join(' | ')}\n`);

    const safe = s.profile.replace(/[\\/:*?"<>|\s]/g, '_');
    const csvPath = path.join(outDir, `${safe}.csv`);
    const jsonPath = path.join(outDir, `${safe}.json`);
    fs.writeFileSync(csvPath, toCSV(stripped), 'utf8');
    fs.writeFileSync(jsonPath, JSON.stringify({
      profile: s.profile,
      formTitle: s.formTitle,
      spreadsheetId: s.spreadsheetId,
      sheetName: s.sheetName,
      strippedColumn: { originalIndex1Based: s.nameColumnIndex1Based, reason: 'personal information (name)' },
      headers: stripped[0],
      rows: stripped.slice(1)
    }, null, 2), 'utf8');

    summary.push({
      profile: s.profile,
      formTitle: s.formTitle,
      spreadsheetId: s.spreadsheetId,
      rowCount: dataRowCount,
      columnCount: colCount,
      headers: stripped[0],
      files: { csv: csvPath, json: jsonPath }
    });
  }

  fs.writeFileSync(path.join(outDir, 'summary.json'), JSON.stringify(summary, null, 2), 'utf8');
  console.log('=== summary written to summary.json ===');
}

main().catch((e) => {
  console.error('FATAL:', e.message);
  process.exit(1);
});
