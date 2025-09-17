function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('New Article')
    .addItem('Insert', 'processSubmissions')
    .addToUi();
};

const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("Pieces");

function processSubmissions() {
  const FOLDER_ID = 'FOLDER_ID';
  
  const SEARCH_QUERY = 'SEARCH_QUERY';

  const threads = GmailApp.search(SEARCH_QUERY, 0,5);
  const folder = DriveApp.getFolderById(FOLDER_ID);

  mailHandler(threads, folder);
}

function mailHandler(threads, folder)
{
  if (threads.length === 0) {
    Logger.log("No matching emails found.");
    return;
  }
  let flag = false;
  threads.forEach(thread => {
    // Get the most recent message in the thread to avoid processing duplicates.
    const message = thread.getMessages().pop();
    
    // Check if the message has already been processed.
    if (message.isUnread()) {
      Logger.log("Mail Reading and Doc Creating");
      try {
        const emailBody = message.getPlainBody();
        const attachments = message.getAttachments();
        const date = message.getDate().toLocaleDateString();
        const id = idCreator(date);

        docCreation(emailBody, attachments, id, folder, message);

      } catch (e) {
        Logger.log(`❌ Error processing email from ${message.getFrom()}: ${e.message}`);
        // message.markRead();
      }
    }
    else{
      !flag? ss.toast("No new article","Status") : "";
      flag = true;
      Logger.log("Not an unread email");
    }
  });
}
function docCreation(emailBody, attachments, id, folder, message)
{
  if (attachments.length > 0) {
          // Extract information using regular expressions
          const name = extractValue(emailBody, 'start', 'end');
          const email = extractValue(emailBody, 'start', 'end');
          const organization = extractValue(emailBody, 'start', 'end');
          const phone = extractValue(emailBody,'start', 'end');
          const title = extractValue(emailBody,'start', 'end');

          attachments.forEach(att=>Logger.log(att.getName()));

          const attachmentBlob = attachments[0];
          if(attachmentBlob.getName().endsWith("docx"))
          {

            // convert word file to google doc
            const fileName = attachmentBlob.getName();
            const newFileName = fileName.replace('.docx', '');
            const metadata = {
              title: newFileName,
              mimeType: MimeType.GOOGLE_DOCS,
            };

            const convertedFile = Drive.Files.create(metadata, attachmentBlob, { convert: true });

            DriveApp.getFileById(convertedFile.id).moveTo(folder);
            
            const doc = DocumentApp.openById(convertedFile.id);
            const docContent = doc.getBody().getText();

            const newDocName = `${id}`;

            // Make a new doc and append desire data
            const frame = DriveApp.getFileById('FILE_ID');
            const newFileUrl = frame.makeCopy(newDocName, folder).getUrl();
            const newFileBody = DocumentApp.openByUrl(newFileUrl).getBody();
            newFileBody.replaceText("<<name>>", name);
            newFileBody.replaceText("<<email>>", email);
            newFileBody.replaceText("<<organisation>>", organization);
            newFileBody.replaceText("<<phone>>", phone);
            newFileBody.replaceText("<<article>>", docContent);
            newFileBody.replaceText("<<id>>", id);

            Logger.log(`✅ Successfully processed submission for ${name}.`);
            ss.toast("Article Doc file Created","Status")

            insertIntoSheet(id, title, name, email, newFileUrl);
            
            // **Clean up:** Delete the temporary converted file to avoid clutter.
            DriveApp.getFileById(convertedFile.id).setTrashed(true);

            // Mark the email as read
            message.markRead();
            return newFileUrl;
          }
          
        } else {
          Logger.log(`Skipping email from ${message.getFrom()} - no attachment found.`);
          // message.markRead();
        }
}

function insertIntoSheet(id, article, name, email, newFileUrl)
{
  const lastRow = sheet.getDataRange().getNumRows();
  sheet.getRange(lastRow+1,1).setValue(id);
  sheet.getRange(lastRow+1,2).setValue(article);
  sheet.getRange(lastRow+1,3).setValue(name);
  sheet.getRange(lastRow+1,4).setValue(email);
  sheet.getRange(lastRow+1,6).setValue(newFileUrl);
  ss.toast("Inserted into Sheet","Status")
}

function idCreator(date)
{
  const range = sheet.getRange("A1:A").getValues();
  const index = sheet.getDataRange().getNumRows();
  const lastId = range[index-1];
  
  const [month, day, year] = date.split('/');
  const articleNumber = +lastId.toString().split('-')[3];
  let id = year.substring(2) + '-' + (month.length==1? '0'+month : month) + '-' + day;
  id = id + '-0' + (articleNumber+1);
  return id;
}

function extractValue(text, startString, endString) {
  const startIndex = text.indexOf(startString) + startString.length;
  const endIndex = text.indexOf(endString, startIndex);
  if (startIndex === -1 || endIndex === -1) {
    return 'N/A';
  }
  return text.substring(startIndex, endIndex).trim();
}