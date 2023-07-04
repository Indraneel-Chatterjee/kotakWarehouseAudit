async function sendEmail(recipients, ccRecipients, bccRecipients, subject, body, attachments, bodyTags, subjectTags) {
  for (let tag in bodyTags) {
    body = body.replace(new RegExp(tag, "g"), bodyTags[tag]);
  }
  for (let tag in subjectTags) {
    subject = subject.replace(new RegExp(tag, "g"), subjectTags[tag]);
  }
  
  MailApp.sendEmail({
    to: recipients.join(','),
    cc: ccRecipients.join(','),
    bcc: bccRecipients.join(','),
    subject,
    body,
    htmlBody: body,
    attachments
  });
}
