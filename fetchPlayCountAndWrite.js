// fetchPlayCountAndWrite.js — C列URL → E列以降に再生回数を記録
const { GoogleSpreadsheet } = require('google-spreadsheet');
const axios = require('axios');

// ===== 設定 =====
const SHEET_ID   = '1wVFefWuElsq7krWpZjTVcerYOHX7SeBTQujVXI7bdXk';
const SHEET_NAME = '投稿再生回数データ';
const CHUNK_SIZE = 100;        // 100行ごとに処理
// =================

function columnToLetter(col) {
  let temp = '', letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

function getJstTodayStrings() {
  const now = new Date();
  const utc = now.getTime() + now.getTimezoneOffset() * 60000;
  const jst = new Date(utc + 9 * 3600000);
  const y = jst.getFullYear();
  const m = jst.getMonth() + 1;
  const d = jst.getDate();

  const md  = `${m}/${d}`;
  const ymd = `${y}/${String(m).padStart(2,'0')}/${String(d).padStart(2,'0')}`;
  const iso = `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
  return { md, ymd, iso };
}

async function fetchPlayCount(url) {
  try {
    const res = await axios.get(url, {
      headers: { 'User-Agent': 'Mozilla/5.0' },
      timeout: 15000,
      maxContentLength: 20 * 1024 * 1024,
    });
    const html = res.data;
    const match = html.match(/["']?playCount["']?\s*[:=]\s*(\d+)/i);
    const n = match ? Number(match[1]) : 0;
    return Number.isFinite(n) ? n : 0;
  } catch (err) {
    console.error(`❌ ${url}: ${err.message}`);
    return 0;
  }
}

(async () => {
  const creds = JSON.parse(
    Buffer.from(process.env.GOOGLE_CREDS_BASE64, 'base64').toString('utf-8')
  );

  const doc = new GoogleSpreadsheet(SHEET_ID);
  await doc.useServiceAccountAuth(creds);
  await doc.loadInfo();

  const sheet = doc.sheetsByTitle[SHEET_NAME];
  if (!sheet) {
    console.error(`❌ シート「${SHEET_NAME}」が見つかりません`);
    process.exit(1);
  }

  const rowCount = sheet.rowCount;
  const colCount = sheet.columnCount;

  await sheet.loadCells(`A1:${columnToLetter(colCount)}1`);

  const { md, ymd, iso } = getJstTodayStrings();

  let targetCol = null;
  for (let col = 4; col < colCount; col++) {
    const c = sheet.getCell(0, col);
    const raw  = (c.value ?? '').toString().trim();
    const disp = (c.displayValue ?? '').toString().trim();
    if ([raw, disp].some(v => v === md || v === ymd || v === iso)) {
      targetCol = col;
      break;
    }
  }

  if (targetCol === null) {
    for (let col = 4; col < colCount; col++) {
      const c = sheet.getCell(0, col);
      const hasVal = c.value !== null && c.value !== undefined && c.value !== '';
      if (!hasVal) {
        c.value = md;
        targetCol = col;
        break;
      }
    }
    if (targetCol === null) {
      console.error('❌ 空き列がありません（E列以降を増やしてください）');
      process.exit(1);
    }
    await sheet.saveUpdatedCells();
  }

  const targetColLetter = columnToLetter(targetCol + 1);
  console.log(`🗓 書き込み先ヘッダー列: ${targetColLetter} (index=${targetCol})`);

  // 4) C5:C を対象に100行ずつ処理
  for (let startRow = 4; startRow < rowCount; startRow += CHUNK_SIZE) {
    const endRow = Math.min(rowCount - 1, startRow + CHUNK_SIZE - 1);

    const aStart = startRow + 1;
    const aEnd   = endRow + 1;

    const urlRange = `C${aStart}:C${aEnd}`;
    const outRange = `${targetColLetter}${aStart}:${targetColLetter}${aEnd}`;

    await sheet.loadCells(urlRange);
    await sheet.loadCells(outRange);

    let wrote = 0;

    for (let r = startRow; r <= endRow; r++) {
      const urlCell = sheet.getCell(r, 2);         // C列 (index=2)
      const outCell = sheet.getCell(r, targetCol);
      const url     = (urlCell.value || '').toString().trim();

      let playCount = 0;
      if (url && url.startsWith('http') && url.includes('tiktok.com')) {
        playCount = await fetchPlayCount(url);
      }

      if (!Number.isFinite(playCount)) playCount = 0;

      outCell.value = playCount;
      outCell.numberFormat = { type: 'NUMBER', pattern: '0' };
      wrote++;
      console.log(`✅ 行${r + 1} → ${playCount}`);
    }

    await sheet.saveUpdatedCells();
    console.log(`💾 保存: 行${aStart}-${aEnd}（${wrote}件更新）`);
  }

  console.log('🎉 完了');
})().catch(err => {
  console.error('❌ Fatal:', err?.stack || err);
  process.exit(1);
});
