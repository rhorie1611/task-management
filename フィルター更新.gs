/** 設定ここだけ確認してね */
const CONFIG = {
  sheetName: '管理シート', // ←あなたのシート名に変更
  headerRow: 6,            // 見出し行
  colDone: 2,              // 完了列 = B列(2)
  colDateTime: 5,          // 日時 = E列(5)
  colTime: 6,              // 締切(時間) = F列(6)
  triggerCellA1: 'H4'      // 更新ボタン代わりのチェックボックスセル
};

/**
 * ★プロジェクト内で onEdit(e) はこれ1つだけ★
 * - H4 が TRUE になったら フィルタ/ソート更新 → H4 を FALSE に戻す
 * - それ以外の編集は カレンダー処理へ渡す（calendarOnEdit_(e)）
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const range = e.range;
    const sheet = range.getSheet();

    if (sheet.getName() !== CONFIG.sheetName) return;

    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(5000)) return;

    try {
      // ① フィルター更新（H4）
      if (range.getA1Notation() === CONFIG.triggerCellA1) {
        if (e.value !== 'TRUE') return;

        refreshTodoView_(sheet);
        range.setValue(false); // ボタン戻し
        return;
      }

      // ② それ以外はカレンダー処理へ
      if (typeof calendarOnEdit_ === 'function') {
        calendarOnEdit_(e);
      }

    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error(err);
  }
}


/**
 * 手動実行用（PCから実行したい時に便利）
 */
function refreshTodoView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${CONFIG.sheetName}`);
  refreshTodoView_(sheet);
}

/**
 * フィルタ：完了=FALSE のみ表示
 * ソート：日時(E)昇順 → 締切(F)昇順
 * ※見出し行は並べ替え対象から除外（見出しが消える事故防止）
 */
function refreshTodoView_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.headerRow) return;

  const lastCol = sheet.getLastColumn();

  // フィルタ範囲（見出し行を含める）
  const filterRange = sheet.getRange(
    CONFIG.headerRow,
    1,
    lastRow - CONFIG.headerRow + 1,
    lastCol
  );

  // フィルタを作成 or 範囲がズレてたら作り直す
  let filter = sheet.getFilter();
  if (!filter) {
    filterRange.createFilter();
    filter = sheet.getFilter();
  } else {
    const fr = filter.getRange();
    const needRecreate =
      fr.getRow() !== CONFIG.headerRow ||
      fr.getNumRows() !== filterRange.getNumRows() ||
      fr.getNumColumns() !== filterRange.getNumColumns();

    if (needRecreate) {
      filter.remove();
      filterRange.createFilter();
      filter = sheet.getFilter();
    }
  }

  // 完了列：FALSE のみ表示（チェックボックス列想定）
  const firstDataRow = CONFIG.headerRow + 1;
  const doneColLetter = columnToLetter_(CONFIG.colDone);
  const formula = `=$${doneColLetter}${firstDataRow}=FALSE`;

  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenFormulaSatisfied(formula)
    .build();

  filter.setColumnFilterCriteria(CONFIG.colDone, criteria);

  // ★ソート範囲：データ行のみ（見出しを除外）
  const dataRange = sheet.getRange(
    CONFIG.headerRow + 1,
    1,
    lastRow - CONFIG.headerRow,
    lastCol
  );

  // あなたの希望：同じ日付の中でも時間順 → 日時(E) → 締切(F)
  dataRange.sort([
    { column: CONFIG.colDateTime, ascending: true },
    { column: CONFIG.colTime, ascending: true }
  ]);
}

/** 1→A, 2→B... の変換 */
function columnToLetter_(column) {
  let temp = '';
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
