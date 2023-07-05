const PDF_MIME_TYPE = "application/pdf";
const TEMP_PDFS_FOLDER_ID = "1jtIwPQM6C2WKk1KJA2uJZ7l1S2q8xdoe";
const FILE_ID_PATTERN = /\/d\/([^/]+)\//;
const RESPONSES_SPREADSHEET_ID = "1hVvCH84KlDiKzyU-aTnvNId7JDAyWzyLeehzCgRNPFo";
const RESPONSES_SHEET_NAME = "Sheet1";
const REVIEWED_RESPONSES_SPREADSHEET_LINK =
  "https://docs.google.com/spreadsheets/d/1A2DQyZFzv2JjLolD37xPdFspE3RLenn9pMk544xSZeM/edit#gid=559877815";
const REVIEWED_STATUS_COLUMN_NAME = "Reviewed";
const PDF_STATUS_COLUMN_NAME = "PDF Status";
const PDF_STATUS_TEXT = "PDF Generated";
const UPDATED_DOCS_FOLDER_ID = "1T4OXwbNi-NV3vvKuw-gcAgd69hPUKgql";
const PDF_LINK_COLUMN_NAME = "PDF Link";
const DOC_LINK_COLUMN_NAME = "Document Link";
const INSPECTION_REPORT_COLUMN_NAME = "Inspection Report";
const TEMP_MOD_DOCS_FOLDER_ID = "1rbZc5n3hBXYY3Aucp1uKP_PjQsLAJu5p";
const TEMPLATE_ID = "1zesq332QOtJfLwHRQAQkRK_nKPdAnqksiN4_723eCl8";
const RECIPIENTS = ["milind.vedi@agnext.in"];
const CC_RECIPIENTS = ["milind.vedi@agnext.in"];
const BCC_RECIPIENTS = ["milind.vedi@agnext.in"];
const BODY =
  "Dear Inspections Team,<br><br>" +
  "Please find the Inspection Report for the following warehouse attached:<br>" +
  "<b>{Warehouse Code}<br>" +
  "{Warehouse Name}<br>" +
  "{Warehouse Address}<br><br></b>" +
  "Regards";
const SUBJECT = "PDF generated for Inspection of Warehouse {Warehouse Code}";

let responseName;
let errorMailSent;
let pdfStatusIndexNo;

// Update the doc file with required data.
async function Main(e) {
  try {
    const sheet = SpreadsheetApp.openById(
      RESPONSES_SPREADSHEET_ID
    ).getSheetByName(RESPONSES_SHEET_NAME);

    const headerRowValues = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const reviewStatusIndexNo = headerRowValues.indexOf(
      REVIEWED_STATUS_COLUMN_NAME
    );
    pdfStatusIndexNo = headerRowValues.indexOf(PDF_STATUS_COLUMN_NAME);
    const pdfLinkIndexNo = headerRowValues.indexOf(PDF_LINK_COLUMN_NAME);
    const modDocIndexNo = headerRowValues.indexOf(DOC_LINK_COLUMN_NAME);
    const inspectionReportIndexNo = headerRowValues.indexOf(
      INSPECTION_REPORT_COLUMN_NAME
    );
    const sheetValues = sheet.getDataRange().getValues();
    let generatedDocsCount = 0;
    let ungeneratedDocsCount = 0;
    errorMailSent = false;

    if (e.range.getColumn() - 1 != reviewStatusIndexNo) {
      console.log("No request for PDF Generation");
      return;
    } else if (
      e.range.getColumn() - 1 == reviewStatusIndexNo &&
      e.range.getValue() != "Yes"
    ) {
      console.log("No new request for PDF Generation");
      return;
    }

    for (let i = 1; i < sheetValues.length; i++) {
      errorMailSent = false;
      pdfFileLink = "";
      newDocLink = "";
      imgBlobs = [];
      if (
        sheetValues[i][reviewStatusIndexNo] == "Yes" &&
        sheetValues[i][pdfStatusIndexNo] == ""
      ) {
        const inspRepoUrls = sheetValues[i][inspectionReportIndexNo];
        let doc;
        responseName = i;

        try {
          const tempDoc = DocumentApp.openById(TEMPLATE_ID);
          const docFile = DriveApp.getFileById(tempDoc.getId());
          const tempModDocsFolder = DriveApp.getFolderById(
            TEMP_MOD_DOCS_FOLDER_ID
          );
          const modDocId = docFile.makeCopy(tempModDocsFolder).getId();
          doc = DocumentApp.openById(modDocId);
        } catch (error) {
          const e = "Invalid docUrl / Access Denied: " + error.stack;
          console.log(e);
          sendErrorEmail(e);
          sheet
            .getRange(i + 1, pdfStatusIndexNo + 1)
            .setValue("PDF Not Generated");
          ungeneratedDocsCount++;
          continue;
        }
        try {
          console.log("PDF Generation Started for data in row " + i);
          sheet
            .getRange(i + 1, pdfStatusIndexNo + 1)
            .setValue("Generating PDF");
          await updateDocToPdf(doc, inspRepoUrls, i);

          if (errorMailSent == false) {
            sheet
              .getRange(i + 1, pdfStatusIndexNo + 1)
              .setValue("PDF Generated");
            sheet.getRange(i + 1, modDocIndexNo + 1).setValue(newDocLink);
            sheet.getRange(i + 1, pdfLinkIndexNo + 1).setValue(pdfFileLink);

            const warehouseCode =
              sheetValues[i][headerRowValues.indexOf(WAREHOUSE_CODE)];
            const auditorName =
              sheetValues[i][headerRowValues.indexOf(AUDITOR_NAME)];
            const timestamp =
              sheetValues[i][headerRowValues.indexOf(TIMESTAMP)];
            const attachments = [
              DriveApp.getFileById(
                await extractFileIdFromLink(pdfFileLink)
              ).getBlob(),
            ];
            const warehouseDetails = await getWarehouseDetails();
            const warehouseName = warehouseDetails[warehouseCode].name;
            const warehouseAddress = warehouseDetails[warehouseCode].address;

            const bodyTags = {
              "{Warehouse Code}": warehouseCode,
              "{Warehouse Name}": warehouseName,
              "{Warehouse Address}": warehouseAddress,
              "{Auditor Name}": auditorName,
              "{DateTime}": timestamp.toString().substring(0, 15),
              "{Responses Sheet Link}": REVIEWED_RESPONSES_SPREADSHEET_LINK,
            };
            const subjectTags = {
              "{Warehouse Code}": warehouseCode,
            };
            await sendEmail(
              RECIPIENTS,
              CC_RECIPIENTS,
              BCC_RECIPIENTS,
              SUBJECT,
              BODY,
              attachments,
              bodyTags,
              subjectTags
            );
            console.log("PDF Generation Notification mail sent");

            generatedDocsCount++;
          } else {
            if (
              sheet
                .getRange(i + 1, pdfStatusIndexNo + 1)
                .getValue()
                .toString() !=
              "PDF Not Generated. Temporary Network Problem. Please Try Again."
            ) {
              sheet
                .getRange(i + 1, pdfStatusIndexNo + 1)
                .setValue("PDF Not Generated");
            }
            ungeneratedDocsCount++;
          }
        } catch (error) {
          const e =
            "Error updating doc to pdf: " + doc.getName() + " " + error.stack;
          console.log(e);
          sendErrorEmail(e);
        }
      }
    }

    if (generatedDocsCount == 0 && ungeneratedDocsCount == 0) {
      console.log("No new forms");
    } else {
      console.log(generatedDocsCount + " PDF(s) generated.");
      console.log(ungeneratedDocsCount + " PDF(s) unsuccessful.");
    }

    deleteAllFilesInFolder(TEMP_MOD_DOCS_FOLDER_ID);
  } catch (error) {
    const e = "Error processing sheets: " + error.stack;
    console.log(e);
    sendErrorEmail(e);
  }
}

async function extractFileIdFromLink(link) {
  var fileIdRegex = /\/file\/d\/([a-zA-Z0-9_-]+)\//;
  var match = link.match(fileIdRegex);
  if (match && match[1]) {
    var fileId = match[1];
    return fileId;
  }
  return null;
}
