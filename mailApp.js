/* To automatically send emails with attachments from a Google Sheet
This script uses Google Apps Script to create a custom menu in Google Sheets
and send emails with attachments based on the data in the sheet. */

/* For Mail Menu and Mail Sending Functionality
The Mail Menu is added when you open the Google Sheet. */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mail')
    .addItem('Send Mail', 'sendMail')
    .addToUi();
};

function sendMail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("yourSheetName");
  const mailSheet = ss.getSheetByName("yourSheetName");

  const mailData = mailSheet.getDataRange().getValues();

  const data = dataSheet.getDataRange().getValues();
//   Logger.log(data)

  if(!data.length)
  {
    ss.toast("Empty Data");
    return;
  }

  // for toast notification
  ss.toast("Sending Mail...");

  // Preparation for sending mail
  for(let i=1;i<data.length;i++) 
  {
    const subject = mailData[0][1];
    let body = mailData[1][1];
    let id = data[i][0];
    let name = data[i][1];
    let email = data[i][2];
    let fileUrls = data[i][3];
 
    // Logger.log("%s\n%s\n%s\n", name, email, fileUrls);
    body = "Dear " + name + ",\n\n" + body;
    // Logger.log("%s\n%s", subject, body);

    // for fileUrls, we assume it is a comma-separated string of file URLs for multiple attachments
    // Example: "https://drive.google.com/file/d/FILE_ID/view?usp=sharing, https://drive.google.com/file/d/ANOTHER_FILE_ID/view?usp=sharing"
    let urls = fileUrls.toString().split(",");
    // Logger.log(urls);

    // Convert URLs to file blobs
    let fileBlobs = urls.map(url => {
      let fileId = url.toString().split("/")[5];
      const file = DriveApp.getFileById(fileId);
      return file.getBlob();
    })
  
    const options = {
      attachments: fileBlobs,
    };

    // Logger.log(options);

    try{
    // MailApp Api is used to send the mail
      MailApp.sendEmail(email, subject, body, options);
      ss.toast("Mail sent successfully to "+ name);
      Logger.log("Mail sent successfully");
    }
    catch(e)
    {
      Logger.log("Error", e.toString);
      ss.toast("Mail sent error to "+ name);
    }
  }
}
