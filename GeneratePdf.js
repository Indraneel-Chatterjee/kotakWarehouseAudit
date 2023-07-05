// This is a function that takes the document's body and updates warehouse details and pdf links.
async function updateDocToPdf(doc, inspRepoUrls, rowNumber) {
  let warehouseDetails = {};
  const body = doc.getBody();
  let warehouseCode;

  try {
    console.log("Filling form's answers");
    warehouseCode = await fillFormAnswers(doc, rowNumber);
  } catch (error) {
    console.log("Error filling text data from response sheet.", error.stack);
    console.log(error.message);
    if (error.message.includes("Address unavailable")) {
      console.log("Address unavailable. Temporary Network Problem.");
      sheet
        .getRange(rowNumber + 1, pdfStatusIndexNo + 1)
        .setValue(
          "PDF Not Generated. Temporary Network Problem. Please Try Again."
        );
    }
    sendErrorEmail(error.stack);
  }

  try {
    console.log("Loading warehouse data.");
    warehouseDetails = await getWarehouseDetails(); // store warehouse details in object.
  } catch (error) {
    console.log("Error loading warehouse details", error.stack);
    sendErrorEmail(error.stack);
  }

  try {
    console.log("Updating file: " + doc.getName());
    console.log("Updating warehouse details");
    await updateWarehouseDetails(body, warehouseDetails, warehouseCode); // Update warehouse details(Name and Address) in the doc file.
  } catch (error) {
    console.log("Error updating warehouse details", error.stack);
    sendErrorEmail(error.stack);
  }

  try {
    console.log("Resizing Images");
    const images = body.getImages();
    const imageWidth = 180;
    const imageHeight = 200;

    for (let i = 1; i < images.length; i++) {
      let image = images[i];
      image.setWidth(imageWidth);
      image.setHeight(imageHeight);
    }
  } catch (error) {
    console.log("Error while resizing images" + error.stack);
    sendErrorEmail(error.stack);
  }

  try {
    if (inspRepoUrls != "" || inspRepoUrls.length != 0) {
      console.log("Replacing Pdfs with images");
      await replacePdfWithImages(body, inspRepoUrls); // Replace pdf file with its images
    }
  } catch (error) {
    console.log("Error while replacing pdfs", error.stack);
    sendErrorEmail(error.stack);
  }

  try {
    if (errorMailSent == false) {
      await deleteFilesIfAlreadyPresent(
        DriveApp.getFolderById(UPDATED_DOCS_FOLDER_ID),
        doc.getName()
      );
      doc.saveAndClose(); // Save the document.
      const docFile = DriveApp.getFileById(doc.getId());
      const upDcsFolder = DriveApp.getFolderById(UPDATED_DOCS_FOLDER_ID);
      const newDoc = docFile.makeCopy(upDcsFolder);
      newDoc.setName(doc.getName());
      const newDocId = newDoc.getId();

      console.log("Saving doc as pdf.");
      await saveDocAsPDF(newDocId); // Convert the document to pdf.
    }
  } catch (error) {
    console.log("Error saving Doc as Pdf", error.stack);
    sendErrorEmail(error.stack);
  }
}
