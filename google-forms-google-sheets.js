function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var createForm = [{name: "Create Form", functionName: "CreateFormfromSheet"}];
  ss.addMenu("Create Form", createForm);
}

function CreateFormfromSheet() {
// var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
// for ( var k = 0 ; k<sheets.length ; k++) {
// var sheet = sheets[k];
 var sheet = SpreadsheetApp.getActiveSpreadsheet();  
 
 var range = sheet.getDataRange(); 
 var data = range.getValues();
 var note = range.getNotes();
 var numberRows = range.getNumRows();
 var numberColumns = range.getNumColumns();
 var firstRow = 1;
 
 var form = FormApp.create(data[2][5]);
 form.setDescription(note[2][5]);
 form.requiresLogin();

 // Toggle Menus 
 form.setAllowResponseEdits(data[8][5]);
 form.setPublishingSummary(data[6][13]);
 form.setLimitOneResponsePerUser(data[8][13]);
 form.setProgressBar(data[6][17]);
 form.setShowLinkToRespondAgain(data[8][17]);
 form.setShuffleQuestions(data[8][21]);
 form.setConfirmationMessage(data[2][35]);

/* Not working... 
***************************************************************************************************************
  // form.setCollectEmail(data[6][9]); ------> SEE REASON: https://goo.gl/lwpTpA
  // form.setRequireLogin(true);  

*/

/* To Do... 
***************************************************************************************************************
  addEditors(emailAddresses)
  
*/


 for(var i=0;i<numberRows;i++){
  var questionType = data[i][5]; 
  if (questionType==''){
     continue;
  }
  
  else if(questionType=='PAGE'){
   form.addPageBreakItem()
     .setTitle(data[i][14])
     .setHelpText(note[i][14]);
  }
  
  else if(questionType=='SECTION'){
   form.addSectionHeaderItem()
     .setTitle(data[i][14])
     .setHelpText(note[i][14]);
  }
  
  else if(questionType=='TEXT'){
   form.addTextItem()
     .setTitle(data[i][14])
     .setHelpText(note[i][14])
     .setRequired(data[i][9]);   
  }
  
  else if(questionType=='PARAGRAPH'){
   form.addParagraphTextItem()
     .setTitle(data[i][14])
     .setHelpText(note[i][14])
     .setRequired(data[i][9]);
  }

  else if(questionType=='DROPDOWN'){
  var currentRow = firstRow+i;
  var getSheetRange = sheet.getDataRange().getLastColumn();
  var range_string = 'AJ' + currentRow + ":" + getSheetRange + currentRow;
  var optionsArray = sheet.getRange(range_string).getValues();  
  var choicesForQuestion =[];
    for (var j=0;j<optionsArray[0].length;j++){
      if (optionsArray[0][j] !== "") {
        choicesForQuestion.push(optionsArray[0][j]);
        }
      }
  form.addListItem()
    .setTitle(data[i][14])
    .setHelpText(note[i][14])
    .setChoiceValues(choicesForQuestion)
    .setRequired(data[i][9]);
  }

  else if(questionType=='CHOICE'){
  var currentRow = firstRow+i;
  var getSheetRange = sheet.getDataRange().getLastColumn(); 
  var range_string = 'AJ' + currentRow + ":" + getSheetRange + currentRow;
  var optionsArray = sheet.getRange(range_string).getValues();  
  var choicesForQuestion =[];
    for (var j=0;j<optionsArray[0].length;j++){
      if (optionsArray[0][j] !== "") {
        choicesForQuestion.push(optionsArray[0][j]);
        }
      }
  form.addMultipleChoiceItem()
    .setTitle(data[i][14])
    .setHelpText(note[i][14])
    .setChoiceValues(choicesForQuestion)
    .setRequired(data[i][9]); 
  }

  else if(questionType=='CHECKBOX'){
  var currentRow = firstRow+i;
  var getSheetRange = sheet.getDataRange().getLastColumn();
  var range_string = 'AJ' + currentRow + ":" + getSheetRange + currentRow;
  var optionsArray = sheet.getRange(range_string).getValues();  
  var choicesForQuestion =[];
    for (var j=0;j<optionsArray[0].length;j++){
      if (optionsArray[0][j] !== "") {
        choicesForQuestion.push(optionsArray[0][j]);
        }
      }
  form.addCheckboxItem()
    .setTitle(data[i][14])
    .setHelpText(note[i][14])
    .setChoiceValues(choicesForQuestion)
    .setRequired(data[i][9]);
  }

  else if(questionType=='GRID'){
  var currentRow = firstRow+i;
  var getSheetRange = sheet.getDataRange().getLastColumn();
  var range_string = 'AJ' + currentRow + ":" + getSheetRange + currentRow;
  var optionsArray = sheet.getRange(range_string).getValues(); 
  var rowTitles =[];
   for (var j=0;j<optionsArray[0].length;j++){
      if (optionsArray[0][j] !== "") {
        rowTitles.push(optionsArray[0][j]);
        }
      }
  var currentRow = firstRow+i+1;
  var getSheetRange = sheet.getDataRange().getLastColumn();
  var range_string = 'AJ' + currentRow + ":" + getSheetRange + currentRow;
  var optionsArray = sheet.getRange(range_string).getValues(); 
  var columnTitles =[];
    for (var j=0;j<optionsArray[0].length;j++){
      if (optionsArray[0][j] !== "") {
        columnTitles.push(optionsArray[0][j]);
        }
      }
  form.addGridItem()
    .setTitle(data[i][14])
    .setHelpText(note[i][14])
    .setRows(rowTitles)
    .setColumns(columnTitles)
    .setRequired(data[i][9]);
  }


  else if(questionType=='DATE'){
   form.addDateItem()
     .setTitle(data[i][14])
     .setHelpText(note[i][14])
     .setRequired(data[i][9]);
  }

  else if(questionType=='TIME'){
   form.addTimeItem()
     .setTitle(data[i][14])
     .setHelpText(note[i][14])
     .setRequired(data[i][9]);
  }

  else if(questionType=='DATETIME'){
   form.addDateTimeItem()
     .setTitle(data[i][14])
     .setHelpText(note[i][14])
     .setRequired(data[i][9]);
  }

  else if(questionType=='DURATION'){
   form.addDurationItem()
     .setTitle(data[i][14])
     .setHelpText(note[i][14])
     .setRequired(data[i][9]);
  }

  else if(questionType=='SCALE'){
  form.addScaleItem()
    .setBounds(data[i][35],data[i][36])
    .setLabels(data[i][37],data[i][38])
    .setTitle(data[i][14])
    .setHelpText(note[i][14])
    .setRequired(data[i][9]);
  }
  
  else if(questionType=='IMAGE'){
   var img = UrlFetchApp.fetch(data[i][35]);
   form.addImageItem()
     .setTitle(data[i][14])
     .setHelpText(note[i][14])
     .setImage(img)
     .setAlignment(FormApp.Alignment.CENTER)
     .setWidth(data[11][25]);
  }

  else if(questionType=='VIDEO'){
   form.addVideoItem()
     .setTitle(data[i][14])
     .setHelpText(note[i][14])
     .setVideoUrl(data[i][35])
     .setAlignment(FormApp.Alignment.CENTER)
     .setWidth(data[11][25]); 
  }

  // else if(questionType=='ACCEPTANCE'){
  //   var item = form.addMultipleChoiceItem();
  //   var goSubmit = item.createChoice('YES', FormApp.PageNavigationType.SUBMIT);
  //   var goRestart = item.createChoice('NO', FormApp.PageNavigationType.RESTART);     
  //     item.setRequired(data[i][9]);
  //     item.setTitle(data[i][14]);
  //     item.setHelpText(note[i][14]);
  //     item.setChoices([goSubmit,goRestart]);   
  // }

  else{
    continue;
  }
 }
}
// }