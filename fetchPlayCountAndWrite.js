// 2) 今日の列を探す（E列= index 4 以降のみ）
let targetCol = null;
for (let col = 4; col < colCount; col++) {   // ★ ここを 1 → 4 に変更
  const c = sheet.getCell(0, col);
  const raw  = (c.value ?? '').toString().trim();
  const disp = (c.displayValue ?? '').toString().trim();
  if ([raw, disp].some(v => v === md || v === ymd || v === iso)) {
    targetCol = col;
    break;
  }
}

// 3) なければ空き列を E 列以降で探す
if (targetCol === null) {
  for (let col = 4; col < colCount; col++) {   // ★ ここも 1 → 4 に変更
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
