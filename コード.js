/*** === Webアプリ入口 === ***/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('デジタル・ライティング・ポートフォリオ')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}

/*** === APIエントリーポイント === ***/
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    let data;

    switch (action) {
      case 'getInitialData':
        data = getInitialData();
        break;
      case 'listProblems':
        data = listProblems(params.spreadsheetId, params.force);
        break;
      case 'getProblem':
        data = getProblem(params.spreadsheetId, params.rowIndex);
        break;
      case 'refreshGradesCache':
        data = refreshGradesCache();
        break;
      case 'saveBackupToCloud':
        data = saveBackupToCloud(params.payload);
        break;
      case 'loadBackupFromCloud':
        data = loadBackupFromCloud(params.id4);
        break;
      case 'deleteBackupFromCloud':
        data = deleteBackupFromCloud(params.id4);
        break;
      case 'enqueueEssay':
        data = enqueueEssay(params.payload);
        break;
      case 'triggerBatchProcess':
        data = triggerBatchProcess();
        break;
      default:
        throw new Error('不明なアクションです: ' + action);
    }

    // CORSのプリフライト回避のため、フロントからは text/plain で送られてくるが、
    // 戻り値はJSONとして返す
    const output = ContentService.createTextOutput(JSON.stringify({ ok: true, data: data }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;

  } catch (err) {
    const output = ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  }
}

/*** === 共通ユーティリティ === ***/
function _s(v){ return v==null ? '' : String(v).trim(); }

/** rgb(...) / 3桁hex / 色名 などを #rrggbb に寄せる。解釈不能なら null（DocumentApp の「色の値が無効」を防ぐ） */
function normalizeColorForGAS_(raw) {
  if (raw == null) return null;
  var c = String(raw).trim();
  if (!c) return null;
  var lc = c.toLowerCase();
  if (lc === 'transparent' || lc === 'inherit') return null;

  var hex6 = /^#([0-9a-f]{6})$/i.exec(c);
  if (hex6) return '#' + hex6[1].toLowerCase();
  var hex3 = /^#([0-9a-f]{3})$/i.exec(c);
  if (hex3) {
    var h = hex3[1].toLowerCase();
    return '#' + h.charAt(0) + h.charAt(0) + h.charAt(1) + h.charAt(1) + h.charAt(2) + h.charAt(2);
  }
  var rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(c);
  if (rgb) {
    function toHexByte(n) {
      var x = Math.max(0, Math.min(255, parseInt(n, 10)));
      var t = x.toString(16);
      return t.length === 1 ? '0' + t : t;
    }
    return '#' + toHexByte(rgb[1]) + toHexByte(rgb[2]) + toHexByte(rgb[3]);
  }
  var named = {
    red: '#ff0000', blue: '#0000ff', green: '#008000', purple: '#800080', gray: '#808080', grey: '#808080',
    black: '#000000', white: '#ffffff', orange: '#ffa500', yellow: '#ffff00', pink: '#ffc0cb',
    lightgreen: '#90ee90', lightblue: '#add8e6', brown: '#a52a2a', navy: '#000080', maroon: '#800000',
    lime: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', silver: '#c0c0c0', gold: '#ffd700', indigo: '#4b0082',
    violet: '#ee82ee', coral: '#ff7f50', salmon: '#fa8072', khaki: '#f0e68c', lavender: '#e6e6fa'
  };
  if (named[lc]) return named[lc];
  return null;
}

function setTextForegroundSafe_(textElem, color) {
  var hex = normalizeColorForGAS_(color);
  if (!hex) return;
  try {
    textElem.setForegroundColor(hex);
  } catch (e) {}
}

function setTextBackgroundSafe_(textElem, color) {
  var hex = normalizeColorForGAS_(color);
  if (!hex) return;
  try {
    textElem.setBackgroundColor(hex);
  } catch (e) {}
}

function setTextFontSizeSafe_(textElem, pt) {
  var n = parseFloat(pt);
  if (isNaN(n) || n <= 0) return;
  n = Math.min(36, Math.max(6, n));
  try {
    textElem.setFontSize(n);
  } catch (e) {}
}

function setTextFontFamilySafe_(textElem, family) {
  if (!family) return;
  var name = String(family).trim();
  if (!name) return;
  try {
    textElem.setFontFamily(name);
  } catch (e) {
    try {
      textElem.setFontFamily('Arial');
    } catch (e2) {}
  }
}

function setTextDecorationsSafe_(textElem, run) {
  var bold = !!(run && run.bold);
  var italic = !!(run && run.italic);
  var underline = !!(run && run.underline);
  try { textElem.setBold(bold); } catch (e) {}
  try { textElem.setItalic(italic); } catch (e) {}
  try { textElem.setUnderline(underline); } catch (e) {}
}

function getParentFolderId_() {
  const scriptId = ScriptApp.getScriptId();
  const cache = CacheService.getScriptCache();
  const key = 'PARENT_ID:' + scriptId;
  let parentId = cache.get(key);
  if (parentId) return parentId;

  const thisFile = DriveApp.getFileById(scriptId);
  const parents = thisFile.getParents();
  if (!parents.hasNext()) throw new Error('親フォルダが見つかりませんでした。');
  parentId = parents.next().getId();
  cache.put(key, parentId, 3600);
  return parentId;
}

function getSubfolderId_(parentId, name, {createIfMissing=false}={}) {
  const parent = DriveApp.getFolderById(parentId);
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next().getId();
  if (!createIfMissing) throw new Error(`フォルダが見つかりません: ${name}`);
  return parent.createFolder(name).getId();
}

function countWords_(text) {
  const m = (_s(text)).match(/[A-Za-z]+(?:[-'][A-Za-z]+)*/g);
  return m ? m.length : 0;
}

function formatTimestampJST_(dateObj) {
  return Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function sanitizeFilename_(name) {
  return _s(name).replace(/[\\\/:*?"<>|\u0000-\u001F]/g, ' ').replace(/\s+/g, ' ').trim();
}

/** ===================== アプリ内キャッシュ ===================== **/
function _getAppCache_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('APP_CACHE_JSON');
  return raw ? JSON.parse(raw) : { updatedAt: null, grades: [], problems: {} };
}
function _setAppCache_(obj) {
  PropertiesService.getScriptProperties().setProperty('APP_CACHE_JSON', JSON.stringify(obj));
}
function _setCachedGrades_(grades) {
  const c = _getAppCache_();
  c.grades = grades || [];
  c.updatedAt = new Date().toISOString();
  _setAppCache_(c);
}
function _setCachedProblems_(spreadsheetId, list) {
  const c = _getAppCache_();
  c.problems = c.problems || {};
  c.problems[spreadsheetId] = { list: list || [], updatedAt: new Date().toISOString() };
  _setAppCache_(c);
}
function _getCachedGrades_() {
  const c = _getAppCache_();
  return c.grades || [];
}
function _getCachedProblems_(spreadsheetId) {
  const c = _getAppCache_();
  return (c.problems && c.problems[spreadsheetId] && c.problems[spreadsheetId].list) || [];
}

function getInitialData() {
  ensureAppStructure_();

  // バッチ処理のトリガー確認
  const props = PropertiesService.getScriptProperties();
  const lastBatchTimeStr = props.getProperty('LAST_BATCH_PROCESS_TIME');
  const now = new Date().getTime();
  const lastTime = lastBatchTimeStr ? parseInt(lastBatchTimeStr, 10) : 0;
  
  if (now - lastTime > 10 * 60 * 1000) { // 10分以上
    try {
      processBatchSubmissions();
    } catch(e) {
      console.error('Initial batch process failed:', e);
    }
  }

  return {
    grades: listGradeSheetsInAssignments(false),
    gemUrl: PropertiesService.getScriptProperties().getProperty('GEM_URL') || ''
  };
}

function listGradeSheetsInAssignments(force) {
  if (!force) {
    const cached = _getCachedGrades_();
    if (cached.length) return cached;
  }
  const parentId = getParentFolderId_();
  const assignmentsFolderId = getSubfolderId_(parentId, 'Assignments', {createIfMissing:true});
  const folder = DriveApp.getFolderById(assignmentsFolderId);
  const it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  const result = [];
  while (it.hasNext()) {
    const f = it.next();
    result.push({ id: f.getId(), name: f.getName() });
  }
  result.sort((a,b)=>a.name.localeCompare(b.name,'ja'));
  _setCachedGrades_(result);
  return result;
}

function listProblems(spreadsheetId, force) {
  if (!force) {
    const cached = _getCachedProblems_(spreadsheetId);
    if (cached.length) return cached;
  }
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sh = ss.getSheetByName('Assignments');
  if (!sh) throw new Error('「Assignments」シートが見つかりません');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const header = values[0].map(_s);
  const idx = { serial: header.indexOf('通し番号'), title: header.indexOf('見出し'), body: header.indexOf('問題文'), note: header.indexOf('備考') };

  const out = [];
  for (let r=1; r<values.length; r++) {
    const row = values[r];
    const serial = _s(row[idx.serial]), title=_s(row[idx.title]), body=_s(row[idx.body]), note=_s(row[idx.note]);
    if (!serial && !title && !body && !note) continue;
    out.push({ rowIndex: r+1, serial, title, body, note, label: `${serial} ${title}`, serialNum: Number(serial) || 999 });
  }
  out.sort((a,b)=>a.serialNum - b.serialNum);
  _setCachedProblems_(spreadsheetId, out);
  return out;
}

function getProblem(spreadsheetId, rowIndex) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sh = ss.getSheetByName('Assignments');
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(_s);
  const col = { serial: header.indexOf('通し番号')+1, title: header.indexOf('見出し')+1, body: header.indexOf('問題文')+1, note: header.indexOf('備考')+1 };
  const row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
  return { serial: _s(row[col.serial-1]), title: _s(row[col.title-1]), body: _s(row[col.body-1]), note: _s(row[col.note-1]), label: `${_s(row[col.serial-1])} ${_s(row[col.title-1])}` };
}

function refreshGradesCache() { return listGradeSheetsInAssignments(true); }

/*** === アプリ初期化とバッチ処理 === ***/

function ensureAppStructure_() {
  const parentId = getParentFolderId_();
  const assignmentsFolderId = getSubfolderId_(parentId, 'Assignments', {createIfMissing:true});
  getSubfolderId_(parentId, 'inbox_submissions', {createIfMissing:true});
  getSubfolderId_(parentId, 'processed_submissions', {createIfMissing:true});
  getSubfolderId_(parentId, 'Portfolios', {createIfMissing:true});
  
  const assignmentsFolder = DriveApp.getFolderById(assignmentsFolderId);
  const files = assignmentsFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  if (!files.hasNext()) {
    const ss = SpreadsheetApp.create('01_Sample Subject');
    const ssFile = DriveApp.getFileById(ss.getId());
    // Move to assignments folder
    try {
      ssFile.moveTo(assignmentsFolder);
    } catch(e) {
      assignmentsFolder.addFile(ssFile);
      DriveApp.getRootFolder().removeFile(ssFile);
    }
    setupWorkbookSheets_(ss);
  } else {
    while (files.hasNext()) {
      const ss = SpreadsheetApp.open(files.next());
      setupWorkbookSheets_(ss);
    }
  }
}

function setupWorkbookSheets_(ss) {
  let ash = ss.getSheetByName('Assignments');
  if (!ash) {
    ash = ss.insertSheet('Assignments');
    ash.appendRow(['通し番号', '見出し', '問題文', '備考']);
    ash.appendRow([1, 'Sample 1', 'Write a self-introduction.', '']);
  }
  let ssh = ss.getSheetByName('Submissions');
  if (!ssh) {
    ssh = ss.insertSheet('Submissions');
    ssh.appendRow(['Timestamp', 'ID', 'Name', 'Subject', 'Serial', 'Title', 'Word Count', 'Original Draft', 'Corrected Text', 'Highlight Data', 'Reflection', 'AI Feedback']);
    ssh.setFrozenRows(1);
  }
}

function triggerBatchProcess() {
  try {
    return processBatchSubmissions();
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

function processBatchSubmissions() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('現在他の処理が実行中です。しばらくしてから再度お試しください。');
  }

  try {
    const parentId = getParentFolderId_();
    const inboxId = getSubfolderId_(parentId, 'inbox_submissions', {createIfMissing:true});
    const processedId = getSubfolderId_(parentId, 'processed_submissions', {createIfMissing:true});
    
    const inboxFolder = DriveApp.getFolderById(inboxId);
    const processedFolder = DriveApp.getFolderById(processedId);
    const files = inboxFolder.getFilesByType(MimeType.JSON);
    
    if (!files.hasNext()) {
      return { ok: true, message: '未処理の提出ファイルは存在しません。' };
    }
    
    const recordsBySpreadsheet = {};
    const processedFiles = [];
    
    while (files.hasNext()) {
      const file = files.next();
      const content = file.getBlob().getDataAsString();
      try {
        const record = JSON.parse(content);
        if (record.spreadsheetId) {
          if (!recordsBySpreadsheet[record.spreadsheetId]) {
            recordsBySpreadsheet[record.spreadsheetId] = [];
          }
          recordsBySpreadsheet[record.spreadsheetId].push(record);
        }
        processedFiles.push(file);
      } catch (e) {}
    }
    
    for (const ssId in recordsBySpreadsheet) {
      try {
        const ss = SpreadsheetApp.openById(ssId);
        setupWorkbookSheets_(ss);
        const sh = ss.getSheetByName('Submissions');
        
        const rows = recordsBySpreadsheet[ssId].map(r => [
          r.timestamp || '',
          r.id4 || '',
          r.name || '',
          r.gradeSubject || '',
          r.serial || '',
          r.title || '',
          r.wordCount || 0,
          r.message || '',
          r.correctedText || '',
          r.highlightData ? JSON.stringify(r.highlightData) : '',
          r.reflection || '',
          r.feedback || ''
        ]);
        
        if (rows.length > 0) {
          sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
        }
      } catch (e) {
        throw new Error(`スプレッドシートへの書き込みに失敗しました: ${e.message}`);
      }
    }
    
    processedFiles.forEach(file => {
      try {
        file.moveTo(processedFolder);
      } catch(e) {
        processedFolder.addFile(file);
        inboxFolder.removeFile(file);
      }
    });
    
    PropertiesService.getScriptProperties().setProperty('LAST_BATCH_PROCESS_TIME', new Date().getTime().toString());
    
    return { ok: true, message: `${processedFiles.length} 件のファイルを一斉書き込みしました。` };
  } finally {
    lock.releaseLock();
  }
}

/*** === バックアップ（一時保存）機能 === ***/

function saveBackupToCloud(payload) {
  const id4 = _s(payload.id4);
  if (!id4) throw new Error('4桁IDが入力されていません。');
  
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  
  // 1. 古いデータの削除（2週間以上前のデータを消去して容量確保 = ガベージコレクション）
  const allProps = props.getProperties();
  const twoWeeksAgo = new Date(now.getTime() - 14 * 24 * 60 * 60 * 1000);
  for (const key in allProps) {
    if (key.startsWith('DWP_BACKUP_')) {
      try {
        const data = JSON.parse(allProps[key]);
        if (new Date(data.savedAt) < twoWeeksAgo) {
          props.deleteProperty(key);
        }
      } catch(e) {
        // 破損データは念のため削除
        props.deleteProperty(key);
      }
    }
  }

  // 2. 保存処理
  const keyToSave = 'DWP_BACKUP_' + id4;
  const dataToSave = {
    id4: payload.id4,
    name: payload.name,
    message: payload.message,
    correctedText: payload.correctedText,
    highlightData: payload.highlightData,
    feedback: payload.feedback,
    reflection: payload.reflection,
    gradeSelect: payload.gradeSelect,
    problemSelect: payload.problemSelect,
    savedAt: now.toISOString(),
    savedAtJST: formatTimestampJST_(now)
  };
  
  const jsonStr = JSON.stringify(dataToSave);
  const byteSize = Utilities.newBlob(jsonStr).getBytes().length;
  // PropertiesService は 1プロパティあたり 9KB (9216 bytes) 制限
  if (byteSize > 9000) {
    throw new Error('データ容量が大きすぎます(最大約3000文字)。少し文章を削ってから再度保存してください。');
  }
  
  props.setProperty(keyToSave, jsonStr);
  return { ok: true, savedAt: dataToSave.savedAtJST };
}

function loadBackupFromCloud(id4) {
  if (!id4) throw new Error('4桁IDが入力されていません。');
  const props = PropertiesService.getScriptProperties();
  const jsonStr = props.getProperty('DWP_BACKUP_' + id4);
  if (!jsonStr) return null;
  try {
    return JSON.parse(jsonStr);
  } catch(e) {
    return null;
  }
}

function deleteBackupFromCloud(id4) {
  if (!id4) throw new Error('4桁IDが入力されていません。');
  PropertiesService.getScriptProperties().deleteProperty('DWP_BACKUP_' + id4);
  return { ok: true };
}

/*** === 受付：PDF生成 ＆ 保存 === ***/
function enqueueEssay(payload) {
  const id4 = _s(payload.id4);
  const name = _s(payload.name);
  const message = _s(payload.message);
  const correctedTextHTML = _s(payload.correctedText);
  const correctedTextPlain = _s(payload.correctedTextPlain || payload.correctedText).replace(/<[^>]+>/g, '');
  const highlightData = payload.highlightData;
  const feedbackHTML = _s(payload.feedback);
  const feedbackPlain = _s(payload.feedbackPlain || payload.feedback).replace(/<[^>]+>/g, '');
  const reflection = _s(payload.reflection);
  const feedbackData = payload.feedbackData;
  const inheritPdfStyles = payload.inheritPdfStyles;
  const spreadsheetId = _s(payload.spreadsheetId);
  const serial = _s(payload.serial);
  const title = _s(payload.title);
  const imageB64 = payload.imageB64;

  const gradeSubject = spreadsheetId ? DriveApp.getFileById(spreadsheetId).getName() : 'Unknown';
  const wordCount = countWords_(message);
  const timestamp = formatTimestampJST_(new Date());

  // 先頭に【ID】を付与し、その後に各項目をアンダーバーで連結
  const fnameBase = `【${id4}】` + [sanitizeFilename_(gradeSubject), sanitizeFilename_(serial), sanitizeFilename_(id4), String(wordCount)].join('_');
  const parentId = getParentFolderId_();

  // 1. JSON保存
  const record = { timestamp, spreadsheetId, gradeSubject, serial, title, id4, name, message, correctedText: correctedTextHTML, correctedTextPlain, highlightData, wordCount, feedback: feedbackHTML, feedbackPlain, feedbackData, reflection };
  const inboxId = getSubfolderId_(parentId, 'inbox_submissions', {createIfMissing:true});
  DriveApp.getFolderById(inboxId).createFile(Utilities.newBlob(JSON.stringify(record, null, 2), 'application/json', fnameBase + '.json'));

  // 2. PDF生成 (Portfolios)
  const portfolioFolderId = getSubfolderId_(parentId, 'Portfolios', {createIfMissing:true});
  const doc = DocumentApp.create(fnameBase + '_temp');
  const body = doc.getBody();

  // 余白設定（狭く：36pt = 約12.7mm）
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);

  // タイトル
  body.appendParagraph("Digital Writing Portfolio").setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  const info = body.appendParagraph(`提出日: ${timestamp}  /  学年: ${gradeSubject}\n問題: ${serial} ${title}  /  ID: ${id4}  /  氏名: ${name}\n語数: ${wordCount} words`);
  info.setFontSize(11);

  // 1. 原文
  body.appendParagraph("１．【原文 (Original Draft)】").setHeading(DocumentApp.ParagraphHeading.HEADING2).setFontSize(18);
  body.appendParagraph(message + "\n").setFontSize(14).setLineSpacing(1.5);

  // 2. 添削後文
  if (correctedTextPlain || correctedTextHTML) {
    body.appendParagraph("２．【添削後文 (Corrected Text)】").setHeading(DocumentApp.ParagraphHeading.HEADING2).setFontSize(18);
    if (highlightData && highlightData.length > 0) {
      const p = body.appendParagraph("");
      p.setLineSpacing(1.5);
      p.setFontSize(14);
      highlightData.forEach(run => {
        const textElem = p.appendText(run.text);
        setTextForegroundSafe_(textElem, run.color);
        setTextBackgroundSafe_(textElem, run.bg);
        setTextDecorationsSafe_(textElem, run);
        if (run.fontSizePt != null) setTextFontSizeSafe_(textElem, run.fontSizePt);
        if (run.fontFamily) setTextFontFamilySafe_(textElem, run.fontFamily);
      });
      body.appendParagraph("\n");
    } else {
      body.appendParagraph(correctedTextPlain + "\n").setFontSize(14).setLineSpacing(1.5);
    }
  }

  // 3. 振り返り
  body.appendParagraph("３．【振り返り (Reflection)】").setHeading(DocumentApp.ParagraphHeading.HEADING2).setFontSize(18);
  body.appendParagraph((reflection || "（入力なし）") + "\n").setFontSize(14);

  // 画像
  if (imageB64) {
    try {
      body.appendParagraph("【添付画像】").setHeading(DocumentApp.ParagraphHeading.HEADING2).setFontSize(18);
      const b64Data = imageB64.split(',')[1];
      const imgBlob = Utilities.newBlob(Utilities.base64Decode(b64Data), "image/jpeg", "handwriting");
      const img = body.appendImage(imgBlob);
      
      // 画像サイズ調整（余白ギリギリいっぱいまで横幅を合わせる）
      const availableWidth = body.getPageWidth() - body.getMarginLeft() - body.getMarginRight();
      const ratio = availableWidth / img.getWidth();
      img.setWidth(availableWidth).setHeight(img.getHeight() * ratio);
    } catch(e) {}
  }

  // 4. AIフィードバック
  body.appendParagraph("４．【AI フィードバック (AI Feedback)】").setHeading(DocumentApp.ParagraphHeading.HEADING2).setFontSize(18);
  if (inheritPdfStyles && feedbackData && feedbackData.length > 0) {
    const p = body.appendParagraph("");
    p.setFontSize(13);
    feedbackData.forEach(run => {
      const textElem = p.appendText(run.text);
      setTextForegroundSafe_(textElem, run.color);
      setTextBackgroundSafe_(textElem, run.bg);
      setTextDecorationsSafe_(textElem, run);
    });
    body.appendParagraph("\n");
  } else {
    body.appendParagraph((feedbackPlain || "（入力なし）") + "\n").setFontSize(13);
  }

  doc.saveAndClose();
  const pdfBlob = doc.getAs('application/pdf').setName(fnameBase + ".pdf");
  const pdfFile = DriveApp.getFolderById(portfolioFolderId).createFile(pdfBlob);
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  // 学校のセキュリティ制限（外部公開禁止）を回避しつつ、生徒がダウンロードできるように権限を設定
  try {
    // 優先：学校・組織のドメイン内ユーザーのみ閲覧可
    pdfFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    try {
      // 個人のアカウント等で実行している場合はこちら
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err) {
      // どちらも許可されていない極めて厳しいセキュリティ設定の場合は何もしない（ドライブには保存されます）
    }
  }

  return { ok:true, fileUrl: pdfFile.getUrl() };
}