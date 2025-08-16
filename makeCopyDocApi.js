// Using DriveApp and DocumentApp to copy a file and replace placeholders
// This function is used to copy a file and replace placeholders with data from the sheet
// The function takes an array of data as input, where each element corresponds to a placeholder in the file
function fileCopy(data) {
  const file = DriveApp.getFileById("FILEID");

  const destination = DriveApp.getFolderById("FOLDERID");
  

  const newFileName = "YourFileName for "+data[0];
  const newFileUrl = file.makeCopy(newFileName, destination).getUrl();

  const newFileBody = DocumentApp.openByUrl(newFileUrl).getBody();
  newFileBody.replaceText("<<name>>", data[0]);
  newFileBody.replaceText("<<email>>", data[1])
  newFileBody.replaceText("<<age>>", data[2]);
  newFileBody.replaceText("<<address>>", data[3]);

  Logger.log(newFileUrl);
  return newFileUrl;
}