function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  const tag = e.parameter.tag;
  const id = e.parameter.id;
  if (!tag || !id) return sendTransparentGif(); // tagまたはidが無ければ無視

  const sheet = sheet.getSheetByName(tag);
  if (!sheet) return sendTransparentGif(); // シートが存在しなければ無視

  const expectedId = sheet.getRange("E1").getValue();
  if (expectedId != id) return sendTransparentGif(); // E1の値とidが一致しなければ無視

  const email = Session.getActiveUser().getEmail() || "不明(外部ユーザー)";
  const time = new Date();

  sheet.appendRow([time, email]);

  return sendTransparentGif();
}

// 透明1px GIF を返す共通関数（ここは割と適当）
function sendTransparentGif() {
  const imageData = Utilities.base64Decode("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");
  return ContentService.createBinaryOutput(imageData)
    .setMimeType(ContentService.MimeType.GIF);
}