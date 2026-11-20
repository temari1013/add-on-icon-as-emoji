/** @OnlyCurrentDoc */
function onOpen() {
  DocumentApp.getUi().createMenu('発言アイコン')
    .addItem('自分のアイコンを挿入', 'insertMyIcon')
    .addToUi();
}

function insertMyIcon() {
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  const ui = DocumentApp.getUi();
  if (!cursor) {
    ui.alert('文字の選択を解除して、挿入したい位置をクリックしてください。');
    return;
  }

  const avatar = getMyAvatar_();
  if (!avatar) return;
  const size = iconHeight_(cursor, doc);
  const image = cursor.insertInlineImage(avatar.blob);
  if (!image) {
    ui.alert('この位置には画像を挿入できません。本文内にカーソルを置いてください。');
    return;
  }

  const width = Math.max(1, Math.round(size * image.getWidth() / image.getHeight()));
  image.setWidth(width).setHeight(size).setAltTitle(avatar.name);
  const parent = image.getParent();
  doc.setCursor(doc.newPosition(parent, parent.getChildIndex(image) + 1));
}

function getMyAvatar_() {
  const cache = CacheService.getUserCache();
  const key = 'speaker-avatar-v1';
  const saved = cache.get(key);
  if (saved) {
    const {data, type, name} = JSON.parse(saved);
    return {blob: Utilities.newBlob(Utilities.base64Decode(data), type), name};
  }

  const token = ScriptApp.getIdentityToken();
  if (!token) {
    DocumentApp.getUi().alert('プロフィールを取得できません。設定とプロフィールへのアクセス許可を確認してください。');
    return;
  }
  const payload = Utilities.base64DecodeWebSafe(token.split('.')[1]);
  const profile = JSON.parse(Utilities.newBlob(payload).getDataAsString('UTF-8'));
  if (!profile.picture) {
    DocumentApp.getUi().alert('プロフィール画像を取得できません。Google アカウントの画像とプロフィールへのアクセス許可を確認してください。');
    return;
  }

  const response = UrlFetchApp.fetch(profile.picture, {muteHttpExceptions: true});
  if (response.getResponseCode() !== 200) {
    DocumentApp.getUi().alert('プロフィール画像を読み込めませんでした。時間をおいて再度お試しください。');
    return;
  }
  const blob = response.getBlob();
  const name = profile.name || '発言者';
  const bytes = blob.getBytes();
  // Base64 化した画像とメタデータが 100 KB の上限に収まる場合だけ保存する。
  if (bytes.length <= 74000) {
    const value = JSON.stringify({data: Utilities.base64Encode(bytes), type: blob.getContentType(), name});
    if (Utilities.newBlob(value).getBytes().length <= 100000) {
      cache.put(key, value, 21600);
    }
  }
  return {blob, name};
}

function iconHeight_(cursor, doc) {
  const text = cursor.getSurroundingText();
  let points = null;
  if (text && text.getText().length) {
    // 文頭では右側、それ以外では直前の文字に合わせる。
    const index = Math.max(0, Math.min(cursor.getSurroundingTextOffset() - 1, text.getText().length - 1));
    points = text.getFontSize(index);
  }
  if (!points) {
    let paragraph = cursor.getElement();
    while (paragraph && typeof paragraph.getHeading !== 'function') {
      paragraph = paragraph.getParent();
    }
    const heading = paragraph ? paragraph.getHeading() : DocumentApp.ParagraphHeading.NORMAL;
    points = doc.getActiveTab().asDocumentTab().getBody().getHeadingAttributes(heading)[DocumentApp.Attribute.FONT_SIZE];
  }
  return Math.max(1, Math.round((points || 11) * 96 / 72));
}
