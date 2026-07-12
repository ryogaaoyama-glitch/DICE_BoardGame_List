// ===== 設定 =====
const SHEET_NAME = 'ボドゲ一覧';

// 列インデックス（0始まり）
const COL_SHELF   = 0; // A: 棚番
const COL_NAME    = 1; // B: ボドゲ名
const COL_WEIGHT  = 2; // C: 区分け（重/中/軽）
const COL_TIME    = 3; // D: 時間
const COL_PLAYERS = 4; // E: プレイ人数
const COL_LANG    = 5; // F: 言語
// G列(6): 備考 → 読むだけ
const COL_LENDING = 7; // H: 貸出状況
// I列(8): よみがな
const COL_IMAGE   = 9; // J: 画像URL
// K列(10): タグ

// ===== CORSヘッダー付きレスポンス =====
function createResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== OPTIONSリクエスト（CORS preflight）への対応 =====
function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ===== Webアプリ エントリーポイント =====
// すべてのリクエストをGETで処理（JSONP対応）
function doGet(e) {
  const action   = e.parameter.action;
  const callback = e.parameter.callback; // JSONPコールバック名

  try {
    let result;
    switch (action) {
      case 'getGames':
        result = getGames();
        break;
      case 'getPlayed':
        result = { played: getPlayedGames(e.parameter.email) };
        break;
      case 'lendGame':
        result = lendGame(Number(e.parameter.rowIndex), e.parameter.borrowerName);
        break;
      case 'returnGame':
        result = returnGame(Number(e.parameter.rowIndex));
        break;
      case 'togglePlayed':
        result = togglePlayedGame(e.parameter.email, e.parameter.rowIndex);
        break;
      case 'updateGame':
        result = updateGame(
          e.parameter.rowIndex,
          e.parameter.shelf,
          e.parameter.tags,
          e.parameter.image,
          e.parameter.yomi,
          e.parameter.note
        );
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
    return createJsonpResponse(result, callback);
  } catch(err) {
    return createJsonpResponse({ error: err.message }, callback);
  }
}

// JSONP形式でレスポンスを返す
function createJsonpResponse(data, callback) {
  const json = JSON.stringify(data);
  const output = callback
    ? `${callback}(${json})`
    : json;
  // text/javascriptではなくtext/plainで返すことで
  // スマホブラウザのセキュリティブロックを回避
  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.TEXT);
}

// ===== ゲーム一覧の取得 =====
function getGames() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { error: `シート「${SHEET_NAME}」が見つかりません。` };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

  return values
    .map((row, i) => ({
      rowIndex: i + 2,
      shelf:    String(row[COL_SHELF]).trim(),
      name:     String(row[COL_NAME]).trim(),
      weight:   String(row[COL_WEIGHT]).trim(),
      time:     String(row[COL_TIME]).trim(),
      players:  String(row[COL_PLAYERS]).trim(),
      lang:     String(row[COL_LANG]).trim(),
      note:     String(row[6]).trim(),
      borrower: String(row[COL_LENDING]).trim(),
      yomi:     String(row[8]).trim(),
      image:    String(row[COL_IMAGE]).trim(),
      tags:     String(row[10]).trim(),
    }))
    .filter(g => g.name !== '');
}

// ===== 貸出処理 =====
function lendGame(rowIndex, borrowerName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, message: `シート「${SHEET_NAME}」が見つかりません` };

  const cell    = sheet.getRange(rowIndex, COL_LENDING + 1);
  const current = String(cell.getValue()).trim();
  if (current !== '') return { success: false, message: `「${current}」さんがすでに借りています` };

  cell.setValue(borrowerName);
  return { success: true, message: `「${borrowerName}」さんへの貸し出しを記録しました` };
}

// ===== 返却処理 =====
function returnGame(rowIndex) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, message: `シート「${SHEET_NAME}」が見つかりません` };

  const cell    = sheet.getRange(rowIndex, COL_LENDING + 1);
  const current = String(cell.getValue()).trim();
  if (current === '') return { success: false, message: 'このゲームは貸出中ではありません' };

  cell.clearContent();
  return { success: true, message: `「${current}」さんからの返却を記録しました` };
}

// ===== プレイ済み機能（rowIndexで管理） =====
const SHEET_PLAYED = 'プレイ済み';

function initPlayedSheet() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PLAYED);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PLAYED);
    // ヘッダー行
    sheet.getRange(1, 1).setValue('rowIndex');
    sheet.getRange(1, 2).setValue('ゲーム名');
    syncGameRows(sheet);
  }
  return sheet;
}

// ボドゲ一覧のrowIndexとゲーム名をプレイ済みシートに同期
function syncGameRows(playedSheet) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEET_NAME);
  if (!gamesSheet) return;

  const lastRow = gamesSheet.getLastRow();
  if (lastRow < 2) return;

  // ボドゲ一覧の全ゲームを取得（rowIndex, name）
  const masterGames = gamesSheet.getRange(2, 1, lastRow - 1, 2).getValues()
    .map((r, i) => ({ rowIndex: i + 2, name: String(r[1]).trim() }))
    .filter(g => g.name);

  const playedLastRow = playedSheet.getLastRow();
  const lastCol       = Math.max(playedSheet.getLastColumn(), 2);

  // 現在のプレイ済みシートのデータを取得
  const currentData = playedLastRow > 1
    ? playedSheet.getRange(2, 1, playedLastRow - 1, lastCol).getValues()
    : [];

  // rowIndexをキーにしたマップを作成
  const existingMap = {};
  currentData.forEach(row => {
    const ri = String(row[0]).trim();
    if (ri) existingMap[ri] = row.slice(2); // ユーザーデータ部分
  });

  // 新しいデータを構築
  const newData = masterGames.map(g => {
    const played = existingMap[String(g.rowIndex)] || [];
    return [g.rowIndex, g.name, ...played];
  });

  // 書き直す
  if (playedLastRow > 1) {
    playedSheet.getRange(2, 1, playedLastRow - 1, lastCol).clearContent();
  }

  if (newData.length > 0) {
    const newLastCol = Math.max(...newData.map(r => r.length));
    playedSheet.getRange(2, 1, newData.length, newLastCol).setValues(
      newData.map(r => { while (r.length < newLastCol) r.push(''); return r; })
    );
  }
}

// ユーザーのプレイ済みrowIndex一覧を返す
function getPlayedGames(email) {
  const sheet = initPlayedSheet();
  syncGameRows(sheet);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 3) return [];

  const headers = sheet.getRange(1, 3, 1, lastCol - 2).getValues()[0];
  const userColIndex = headers.indexOf(email);
  if (userColIndex === -1) return [];

  const userCol    = userColIndex + 3;
  const rowIndices = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => Number(r[0]));
  const playedVals = sheet.getRange(2, userCol, lastRow - 1, 1).getValues().map(r => String(r[0]).trim());

  return rowIndices.filter((_, i) => playedVals[i] === '✅');
}

// rowIndexでプレイ済みをトグル
function togglePlayedGame(email, rowIndex) {
  const sheet   = initPlayedSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 2);

  // ユーザー列を探す or 追加
  const headers = lastCol > 2
    ? sheet.getRange(1, 3, 1, lastCol - 2).getValues()[0]
    : [];
  let userColIndex = headers.indexOf(email);
  let userCol;

  if (userColIndex === -1) {
    userCol = lastCol + 1;
    sheet.getRange(1, userCol).setValue(email);
  } else {
    userCol = userColIndex + 3;
  }

  // rowIndexの行を探す
  if (lastRow < 2) return { success: false, message: 'データがありません' };
  const rowIndices = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => Number(r[0]));
  const gameRowIndex = rowIndices.indexOf(Number(rowIndex));

  if (gameRowIndex === -1) return { success: false, message: `rowIndex ${rowIndex} が見つかりません` };

  const sheetRow = gameRowIndex + 2;
  const cell     = sheet.getRange(sheetRow, userCol);
  const current  = String(cell.getValue()).trim();

  if (current === '✅') {
    cell.clearContent();
    return { success: true, played: false };
  } else {
    cell.setValue('✅');
    return { success: true, played: true };
  }
}

// ===== BGG画像URL取得バッチ =====
const BATCH_SIZE = 50;

function fetchBggImages() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('データがありません'); return; }

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  let fetched = 0, skipped = 0, failed = 0;

  for (let i = 0; i < values.length; i++) {
    const name     = String(values[i][COL_NAME]).trim();
    const imageUrl = String(values[i][COL_IMAGE]).trim();
    if (!name || imageUrl !== '') { skipped++; continue; }
    if (fetched >= BATCH_SIZE) { Logger.log(`${BATCH_SIZE}件処理しました。続きは再実行してください。`); break; }

    const url = fetchBggImageUrl(name);
    if (url) { sheet.getRange(i + 2, COL_IMAGE + 1).setValue(url); fetched++; Logger.log(`✅ ${name}`); }
    else      { failed++; Logger.log(`❌ ${name}`); }
    Utilities.sleep(500);
  }

  Logger.log(`完了: 取得=${fetched}, スキップ=${skipped}, 失敗=${failed}`);
  SpreadsheetApp.getUi().alert(`完了！\n取得: ${fetched}件\nスキップ: ${skipped}件\n失敗: ${failed}件`);
}

function fetchBggImageUrl(gameName) {
  try {
    const searchUrl = `https://boardgamegeek.com/xmlapi2/search?query=${encodeURIComponent(gameName)}&type=boardgame&exact=1`;
    const searchRes = UrlFetchApp.fetch(searchUrl, { muteHttpExceptions: true });
    if (searchRes.getResponseCode() !== 200) return null;

    const searchXml = XmlService.parse(searchRes.getContentText());
    const items     = searchXml.getRootElement().getChildren('item');
    let gameId      = null;

    if (items.length > 0) {
      gameId = items[0].getAttribute('id').getValue();
    } else {
      const searchUrl2 = `https://boardgamegeek.com/xmlapi2/search?query=${encodeURIComponent(gameName)}&type=boardgame`;
      const searchRes2 = UrlFetchApp.fetch(searchUrl2, { muteHttpExceptions: true });
      if (searchRes2.getResponseCode() !== 200) return null;
      const items2 = XmlService.parse(searchRes2.getContentText()).getRootElement().getChildren('item');
      if (items2.length === 0) return null;
      gameId = items2[0].getAttribute('id').getValue();
    }

    Utilities.sleep(300);
    const detailRes = UrlFetchApp.fetch(`https://boardgamegeek.com/xmlapi2/thing?id=${gameId}`, { muteHttpExceptions: true });
    if (detailRes.getResponseCode() !== 200) return null;

    const imageEl = XmlService.parse(detailRes.getContentText()).getRootElement().getChild('item').getChild('image');
    if (!imageEl) return null;

    const imageUrl = imageEl.getText().trim();
    return imageUrl.startsWith('//') ? 'https:' + imageUrl : imageUrl;
  } catch(e) {
    Logger.log(`エラー (${gameName}): ${e.message}`);
    return null;
  }
}

// ===== ウォームアップ用トリガー =====
// GASを常にウォームな状態に保つ
// Apps Scriptのトリガーで5分ごとに実行するよう設定してください
function keepWarm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getName(); // 軽い処理でウォームアップ
  Logger.log('keepWarm: ' + new Date().toISOString());
}

// 初回実行時にトリガーを自動設定する関数（手動で1回だけ実行）
function setupWarmupTrigger() {
  // 既存のkeepWarmトリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'keepWarm') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 5分ごとに実行するトリガーを作成
  ScriptApp.newTrigger('keepWarm')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('ウォームアップトリガーを設定しました');
}

// ===== ゲーム編集機能 =====
function updateGame(rowIndex, shelf, tags, image, yomi, note) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, message: 'シートが見つかりません' };

  const row = Number(rowIndex);
  if (!row || row < 2) return { success: false, message: '無効な行番号です: ' + rowIndex };

  // 列番号は1始まり
  // A=1(棚番), B=2(名前), C=3(重さ), D=4(時間), E=5(人数), F=6(言語), G=7(備考), H=8(貸出), I=9(よみがな), J=10(画像URL), K=11(タグ)
  sheet.getRange(row, 1).setValue(shelf || '');  // A列: 棚番
  sheet.getRange(row, 7).setValue(note  || '');  // G列: 備考
  sheet.getRange(row, 9).setValue(yomi  || '');  // I列: よみがな
  sheet.getRange(row, 10).setValue(image || ''); // J列: 画像URL
  sheet.getRange(row, 11).setValue(tags  || ''); // K列: タグ

  return { success: true, message: '更新しました' };
}
