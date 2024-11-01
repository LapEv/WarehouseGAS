function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const dir = DriveApp.getFolderById(params.folderUpload);
  const urlsArr = params.files.map((item) => {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(item.data),
      item.mimeType,
      item.fileName
    );
    return dir.createFile(blob).getUrl();
  });
  const urls = urlsArr.join(', ');
  return ContentService.createTextOutput(urls);
}
