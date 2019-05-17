function createImageFolder(body){
  // 出力先のフォルダを生成
  Logger.log("出力先のフォルダを生成");
  
  var docId = DocumentApp.getActiveDocument().getId();
  Logger.log("ドキュメントID：" + docId);
  var file = DriveApp.getFileById(docId);
  var thisFolder  = file.getParents().next();
  var folderName = "Image";
  
  while(thisFolder.getFoldersByName(folderName).hasNext()){
    var child = thisFolder.getFoldersByName(folderName).next();
    return child;
  }
  
  var newFolder = thisFolder.createFolder(folderName);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  Logger.log("生成しました：" + folderName);
  
  return newFolder;
}
    