async function sendErrorEmail(errorMessage) {
  const subject = "Error in Google Form Automation in: " + responseName;
  const body =
    "An error occurred in Google Form Automation's Apps Script code" +
    "\n\n" +
    errorMessage;
  //const recipients = ["milind.vedi@agnext.in", "inspection.reports@agnext.in"];
  const recipients = ["milind.vedi@agnext.in"];
  GmailApp.sendEmail(recipients.join(","), subject, body);
  errorMailSent = true;
}
