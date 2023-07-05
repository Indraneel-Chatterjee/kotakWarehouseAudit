const PDF_FOLDER_ID = "1eb2PDOWZpJ6lg5rnC10UU7gzUn6B2hbV"; //June
// 1F9AvZ1zeaKkfwATMGjcTwoARtd1FSs6o -> MAY
let pdfFileLink;
let newDocLink;

// This function saves the document file as pdf.
async function saveDocAsPDF(fileId) {
  try {
    const doc = DocumentApp.openById(fileId);
    const folderId = PDF_FOLDER_ID;

    const pdfBlob = doc.getAs(PDF_MIME_TYPE);

    const folder = DriveApp.getFolderById(folderId);
    await deleteFilesIfAlreadyPresent(folder, doc.getName() + ".pdf");
    const file = folder.createFile(pdfBlob);

    pdfFileLink = file.getUrl();
    console.log("PDF " + file.getName() + " saved.");

    doc.saveAndClose();
    newDocLink = doc.getUrl();
  } catch (error) {
    const e = "Error saving doc as pdf. " + error;
    console.log(e);
    sendErrorEmail(e);
  }
}

// This function deletes any file(s) already present to avoid conflicts.
async function deleteFilesIfAlreadyPresent(folder, fileName) {
  try {
    const files = folder.getFilesByName(fileName);

    while (files.hasNext()) {
      const file = files.next();
      file.setTrashed(true);
      console.log("Deleted duplicate file: " + file.getName());
    }
  } catch (error) {
    const e = "Error deleting old files. " + error;
    console.log(e);
    sendErrorEmail(e);
  }
}

async function deleteAllFilesInFolder(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    DriveApp.getFileById(file.getId()).setTrashed(true);
  }
}
