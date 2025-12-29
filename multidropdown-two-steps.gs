function onEdit(e) {
  const ss = e.source; // Get the Spreadsheet object
  const activeSheet = ss.getActiveSheet(); // Get the currently active sheet
  const editedCell = e.range; // Get the cell that was just edited
  const editedColumn = editedCell.getColumn();
  const editedRow = editedCell.getRow();

  // Configurable settings
  const dataSheetName = 'dict'; // data dictionary name
  const entrySheetName = ''; // dropdown sheet
  const stateColumn = 5; // state column number
  const townshipColumn = 6; // township column number


  if (activeSheet.getName() !== entrySheetName || editedRow === 1) {
    return; // Exit if wrong sheet or header row
  }

  if (editedColumn !== stateColumn && editedColumn !== townshipColumn) {
    return; // Exit if edit is not in a dependency column
  }

  let filterBy;
  let validationCell;
  let filterIndex;
  let mapIndex;

  if (editedCell.getColumn() == stateColumn )
  {
    filterBy = editedCell.getValue();
    validationCell = activeSheet.getRange(editedCell.getRow(), townshipColumn);
    filterIndex = 0; // first filter data (e.g. state/region data column in dict)
    mapIndex = 6; // wanna get data (e.g. township data column in dict)
    validationCell.clearContent();
  }

  
  const dataSheet = ss.getSheetByName(dataSheetName);
  if (!dataSheet) {
    Browser.msgBox("Error: Source data sheet '" + dataSheetName + "' not found.");
    return;
  }
  
  const allData = dataSheet.getRange('A2:N' + dataSheet.getLastRow()).getValues();

  const uniqueFilteredValues = [...new Set(
    allData.filter(row => row[filterIndex] === filterBy)
                                  .map(row => row[mapIndex]) 
  )];

  if (uniqueFilteredValues.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueFilteredValues)
      .setAllowInvalid(false)
      .setHelpText('Select a township from the list.')
      .build();
    validationCell.setDataValidation(rule);
  } else {
    validationCell.setDataValidation(null);
  }
}