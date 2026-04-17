/*** === Webアプリ入口 === ***/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('デジタル・ライティング・ポートフォリオ')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}

/*** === 共通ユーティリティ === ***/
function _s(v){ return v==null ? '' : String(v).trim(); }

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
  const assignmentsFolderId = getSubfolderId_(parentId, 'Assignments', {createIfMissing:false});
  const folder = DriveApp.getFolderById(assignmentsFolderId);
  const it = folder.getFiles();
  const result = [];
  while (it.hasNext()) {
    const f = it.next();
    if (f.getMimeType() === MimeType.GOOGLE_SHEETS) {
      result.push({ id: f.getId(), name: f.getName() });
    }
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

/*** === 受付：PDF生成 ＆ 保存 === ***/
function enqueueEssay(payload) {
  const id4 = _s(payload.id4);
  const name = _s(payload.name);
  const message = _s(payload.message);
  const feedback = _s(payload.feedback);
  const reflection = _s(payload.reflection);
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
  const record = { timestamp, gradeSubject, serial, title, id4, name, message, wordCount, feedback, reflection };
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

  // 2. 振り返り
  body.appendParagraph("２．【振り返り (Reflection)】").setHeading(DocumentApp.ParagraphHeading.HEADING2).setFontSize(18);
  body.appendParagraph((reflection || "（入力なし）") + "\n").setFontSize(14);

  // 3. AIフィードバック
  body.appendParagraph("３．【AI フィードバック (AI Feedback)】").setHeading(DocumentApp.ParagraphHeading.HEADING2).setFontSize(18);
  body.appendParagraph((feedback || "（入力なし）") + "\n").setFontSize(13);

  // 画像
  if (imageB64) {
    try {
      body.appendParagraph("【添付画像】").setHeading(DocumentApp.ParagraphHeading.HEADING2).setFontSize(18);
      const b64Data = imageB64.split(',')[1];
      const imgBlob = Utilities.newBlob(Utilities.base64Decode(b64Data), "image/jpeg", "handwriting");
      const img = body.appendImage(imgBlob);
      // 画像サイズ調整（幅一杯まで）
      const maxWidth = 450; 
      if (img.getWidth() > maxWidth) {
        const ratio = maxWidth / img.getWidth();
        img.setWidth(maxWidth).setHeight(img.getHeight() * ratio);
      }
    } catch(e) {}
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