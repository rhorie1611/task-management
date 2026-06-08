function autoCheckTD_200(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  const START_ROW = 200; // テスト開始行
  const COL_G = 7;       // G列
  const COL_K = 11;      // K列（チェック）
  const COL_L = 12;      // L列（自動付与フラグ：TRUEなら自動で付けた）

  // G列以外の編集は無視
  if (range.getColumn() !== COL_G) return;

  // 200行目未満は無視（複数行編集にも対応）
  const rowStart = range.getRow();
  const numRows = range.getNumRows();
  const rowEnd = rowStart + numRows - 1;
  if (rowEnd < START_ROW) return;

  const gValues = range.getValues(); // 1列
  for (let i = 0; i < gValues.length; i++) {
    const r = rowStart + i;
    if (r < START_ROW) continue;

    const g = String(gValues[i][0] ?? "");
    const hasTD = g.includes("【TD】");

    const kCell = sheet.getRange(r, COL_K);
    const lCell = sheet.getRange(r, COL_L);
    const autoFlag = lCell.getValue() === true;

    if (hasTD) {
      // TDあり → 強制ON（自動印も付ける）
      kCell.setValue(true);
      lCell.setValue(true);
    } else {
      // TDなし → 「自動で付けたもの」だけOFFに戻す
      if (autoFlag) {
        kCell.setValue(false);
        lCell.setValue(false);
      }
    }
  }
}
