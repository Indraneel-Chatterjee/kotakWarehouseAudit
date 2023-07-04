// Deprecated

// // This function replaces heic image links with jpeg images.
// async function replaceLinksWithJpegImages(doc) {
//   try {
//     const tables = doc.getTables();

//     for (let i = 0; i < tables.length - 1; i++) {
//       const table = tables[i];
//       const tableRows = table.getNumRows();
//       //const tableColumns = table.getRow(0).getNumCells();

//       for (let j = 0; j < tableRows; j++) {
//         for (let k = 0; k < table.getRow(j).getNumCells(); k++) {
//           const cell = table.getCell(j, k);
//           const cellText = cell.getText();

//           let index = -1;
//           let occurrences = [];

//           // Find all occurrences of "heic" within the cell text
//           while ((index = cellText.indexOf("heic", index + 1)) !== -1) {
//             occurrences.push(index);
//           }

//           index = -1; // Reset index to -1 for searching "heif"
//           while ((index = cellText.indexOf("heif", index + 1)) !== -1) {
//             occurrences.push(index);
//           }

//           // Log the indices of "heic" occurrences in the cell
//           if (occurrences.length > 0) {
//             console.log("Found 'heic/heif' at indices " + occurrences.join(", ") + " in cell (" + j + ", " + k + ")" + "in table " + i + ". Replacing with image.");
//             for (const occourance of occurrences) {
//               const blob = await getJpegBlobFromHeicImageUrl(cell.editAsText().getLinkUrl(occourance));
//               //cell.appendImage(blob);
//               cell.appendImage(blob).setWidth(180).setHeight(200);
//             }
//           }

//           //Remove links
//           if (occurrences.length > 0) {
//             for (let i = 0; i < cell.getNumChildren(); i++) {
//               const url = cell.getChild(i).asText().getAttributes().LINK_URL;
//               if (url != null) {
//                 cell.getChild(i).removeFromParent();
//               }
//             }
//           }
//         }
//       }
//     }
//   } catch (error) {
//     const e = "Error replacing HEIC Image Links." + error.stack;
//     console.log(e);
//     sendErrorEmail(e);
//   }

// }

// async function getJpegBlobFromHeicImageUrl(link) {
//   try {
//     const obj = Drive.Files.get(extractFileIdFromUrl(link));
//     const blob = UrlFetchApp.fetch(obj.thumbnailLink.replace(/=s.+/, "=s1000")).getBlob();
//     return blob;
//   } catch (error) {
//     const e = "Error getting HEIC Image as JPEG." + error.stack;
//     console.log(e);
//     sendErrorEmail(e);
//   }
// }

// function extractFileIdFromUrl(url) {
//   try {
//     let fileId = "";
//     const startIndex = url.indexOf("/d/") + 3;
//     const endIndex = url.indexOf("/view");

//     if (startIndex !== -1 && endIndex !== -1) {
//       fileId = url.substring(startIndex, endIndex);
//     }
//     return fileId;
//   } catch (error) {
//     const e = "Error fetching HEIC Image id from link." + error.stack;
//     console.log(e);
//     sendErrorEmail(e);
//   }
// }