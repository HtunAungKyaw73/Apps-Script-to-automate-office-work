function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mail Reply')
    .addItem('Check', 'mailRpCheck')
    .addToUi();
};

const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("mail-list");
const SEARCH_QUERY = 'SEARCH CRITERIA HERE';

function mailRpCheck(){
  ss.toast("New Mail Reply Checking...");

  try{
    const mailInSheet = sheet.getRange("B:B").getValues();
    const threads = GmailApp.search(SEARCH_QUERY);

    if (threads.length === 0) {
      Logger.log("No matching emails found.");
      return;
    }

    threads.forEach(thread => {
      // Get the most recent message in the thread to avoid processing duplicates.
      const messages = thread.getMessages();
      messages[0].markRead();

      if (messages.length === 1) {
        Logger.log("No replies found in this thread.");
        ss.toast("No reply yet");
        return;
      }

      const replies = messages.filter((message, index) => {
          if(index!==0 && message.isUnread()) {
            return message
          }
      });

      if (!replies.length)
      {
        ss.toast("No new reply yet");
        Logger.log("No new replies found");
        return;
      }
      else{
        ss.toast(`${replies.length} new replies found`);
      }

      const reply_mails = replies.map(reply=>{
        const from = reply.getFrom();
        const mail = from.substring(from.search("<")+1, from.search(">"));
        reply.markRead();
        return mail;
      });

      reply_mails.map(reply => {
        const index = mailInSheet.findIndex(row => row.includes(reply));
        if(index)
        {
          const cell = sheet.getRange(index+1,3);
          const rule = SpreadsheetApp.newDataValidation()
            .requireCheckbox() 
            .setAllowInvalid(false) 
            .build();
          cell.setDataValidation(rule);
          cell.setValue(true); 
        }
      })
      // Logger.log("replies %s", reply_mails.length);
      ss.toast("Reply Mail Added in the sheet");
    })
  }catch(e)
  {
    Logger.log("Error: ", e.message);
    ss.toast(e.message, "Error");
  }
}