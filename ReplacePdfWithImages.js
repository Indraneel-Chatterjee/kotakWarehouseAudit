const INSPECTION_REPORT_TABLE_NUMBER = 8;
const INSPECTION_REPORT_COLUMN_NUMBER = 0;
const INSPECTION_REPORT_ROW_NUMBER = 0;
let imgBlobs = []; // For image blobs for each page in pdfs.

// This function replaces pdf links with images of individual pages in pdf.
async function replacePdfWithImages(body, inspRepoUrls) {
  const table = body.getTables()[INSPECTION_REPORT_TABLE_NUMBER]; // Getting table containing inspection report.
  const cell = table.getCell(
    INSPECTION_REPORT_ROW_NUMBER,
    INSPECTION_REPORT_COLUMN_NUMBER
  ); // Getting cell containing inspection report.
  const links = inspRepoUrls.split(",");

  console.log("Clearing inspection cell.");
  cell.clear();

  try {
    //Getting fileIds from links
    const fileIds = [];
    for (const link of links) {
      fileIds.push(extractFileIdFromFileUrlInArray(link));
    }

    //Getting only pdf files using fileIds
    for (const fileId of fileIds) {
      const file = DriveApp.getFileById(fileId);
      if (file.getMimeType() == PDF_MIME_TYPE) {
        await convertPDFToPNG_(file.getBlob());
      } else if (
        file.getMimeType() == "image/heic" ||
        file.getMimeType() == "image/heif"
      ) {
        const blob = UrlFetchApp.fetch(
          Drive.Files.get(fileId).thumbnailLink.replace(/=s.+/, "=s600")
        ).getBlob();
        imgBlobs.push(blob);
      } else {
        // Deprecated
        try {
          imgBlobs.push(
            UrlFetchApp.fetch(
              Drive.Files.get(fileId).thumbnailLink.replace(/=s.+/, "=s600")
            ).getBlob()
          );
        } catch (error) {
          const e =
            "Error getting blob for image. Image id: " +
            fileId +
            "\n" +
            error.stack;
          console.log(e);
          sendErrorEmail(error.stack);
        }
      }
    }
  } catch (error) {
    console.error(
      "An error occurred while getting links for pdfs/images for Inspection Report Cell from the doc file:",
      error.stack
    );
    sendErrorEmail(error.stack);
  }

  try {
    // Add all images to the cell.
    console.log("No of images -> " + imgBlobs.length);
    for (const imageBlob of imgBlobs) {
      const image = cell.appendImage(imageBlob);

      if (image.getHeight() > 680) {
        image.setHeight(680);
      }

      if (image.getWidth() > 605) {
        image.setWidth(605);
      }
    }
  } catch (error) {
    console.log("Error inserting image in doc.", error.stack);
    sendErrorEmail(error.stack);
    return;
  }
}

// This function creates PDF for each page in the pdf and stores it in a temporary folder. The image for these individual pdfs is used as image for making image blobs.
async function convertPDFToPNG_(blob) {
  const data = new Uint8Array(blob.getBytes());
  const pdfData = await PDFLib.PDFDocument.load(data);
  const pageLength = pdfData.getPageCount();
  console.log(`Total pages: ${pageLength}`);
  const obj = { imageBlobs: [], fileIds: [] }; // Object for storing image blobs and file ids for temporary pdfs.

  //Loop for converting all of pages in pdf to individual pdfs and storing the image from those pdfs to "obj".
  for (let i = 0; i < pageLength; i++) {
    console.log(`Processing page: ${i + 1}`);
    const pdfDoc = await PDFLib.PDFDocument.create();
    const [page] = await pdfDoc.copyPages(pdfData, [i]);
    pdfDoc.addPage(page);
    const bytes = await pdfDoc.save();
    const blob = Utilities.newBlob(
      [...new Int8Array(bytes)],
      MimeType.PDF,
      `sample${i + 1}.pdf`
    );

    const id = DriveApp.getFolderById(TEMP_PDFS_FOLDER_ID)
      .createFile(blob)
      .getId(); // Create pdf file and store its id.
    Utilities.sleep(3200); // Allowing time for the thumbnail of the created file to be prepared.
    let link = Drive.Files.get(id, { fields: "thumbnailLink" }).thumbnailLink;

    if (!link) {
      Utilities.sleep(10000); // In case the thumbnail is not created. Wait for 10 more seconds (Worst Case).
      link = Drive.Files.get(id, { fields: "thumbnailLink" }).thumbnailLink;
      if (!link) {
        throw new Error(console.error("Image Conversion Failed", error.stack));
      }
    }

    // Create image blob and set its content type and name.
    const imageBlob = UrlFetchApp.fetch(link.replace(/\=s\d*/, "=s600")) //600 pixels size
      .getBlob()
      .setContentType("image/png")
      .setName(`page${i + 1}.png`);
    obj.imageBlobs.push(imageBlob); // Deprecated
    obj.fileIds.push(id);
    imgBlobs.push(imageBlob);
  }

  // Delete the temporary pdfs.
  obj.fileIds.forEach((id) => DriveApp.getFileById(id).setTrashed(true));
}

function extractFileIdFromFileUrlInArray(url) {
  var fileId = "";
  var match = /\/open\?id=([a-zA-Z0-9-_]+)(?:\/|$|\?|#)/.exec(url);

  if (match) {
    fileId = match[1];
  }

  return fileId;
}
