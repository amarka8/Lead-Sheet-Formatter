function interface() {
    var ui=SpreadsheetApp.getUi();
    var leadFormatter=ui.createMenu('Lead Formatter');
    leadFormatter.addItem('Format sheets','formatSheetsPrompt');
    leadFormatter.addToUi();
  }
  
  function onOpen(){
    interface();
  }
  
  
  function formatSheetsPrompt() {
    removeDups();
    removeEmptyNames();
    modifyCompanyNames();
    emailEditor();
  }
  
  function letterToColumn(letter)
  {
    // var column = 0, length = letter.length;
    // for (var i = 0; i < length; i++)
    // {
    //   column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    // }
    // return column;
    var column=0;
    var colLetter=letter;
    for (var i = 0; i < colLetter.length; i++)
    {
      column += ((colLetter.charCodeAt(i) - 64) * Math.pow(26,i));
    }
   return column-1;
  }
  
  function containsAny(str, items){
      for(var i in items){
          var item = items[i];
          if (str.indexOf(item) > -1){
              return true;
          }
      }
      return false;
  }
  
  // Removes leads with empty first or last names
  function removeEmptyNames() {
    var app = SpreadsheetApp;
    var ui = SpreadsheetApp.getUi();
    //var sheet = SpreadsheetApp.getActiveSheet();
    var sheets = SpreadsheetApp.getActive().getSheets();
  
    var input = ui.prompt('Enter Column Letter with First Names:',ui.ButtonSet.OK_CANCEL);
    // sheets.forEach(sheet=> {
    if (input.getSelectedButton() == ui.Button.OK) {
      var spreadsheet=app.getActiveSpreadsheet().getSheets().length;
      for (i=0;i<spreadsheet;i++){
        var shts=app.getActiveSpreadsheet().getSheets()[i];
        var originalData=shts.getDataRange().getValues();
        var column=letterToColumn(input.getResponseText().toUpperCase())
        var emptyNames=originalData.filter(item => item[column]!='');
        // Logger.log(emptyNames[9]);
      if (emptyNames.length!=0){
        var deleter=originalData.length-emptyNames.length;
        // var begin=shts.getDataRange().getValues().length;
        shts.getRange(1,1,emptyNames.length,emptyNames[0].length).setValues(emptyNames);
        // Logger.log(emptyNames.length,deleter)
        shts.deleteRows(emptyNames.length+1,deleter);
      }
      // var col = letterToColumn(input.getResponseText().toUpperCase());
      // sheets.forEach(sheet=> {
      //   var data = sheet.getDataRange().getValues();
      //   for (var i = 1; i < data.length; i++) {
      //     if (sheet.getRange(i, col).isBlank()) {
      //       Logger.log(i);
      //     }
      //   }
      }
    }
  }
  
  function removeDups () {
    var app=SpreadsheetApp;
    var spreadsheet=app.getActiveSpreadsheet().getSheets().length;
    for (i=0;i<spreadsheet;i++){
      var shts=app.getActiveSpreadsheet().getSheets()[i];
      var vals=shts.getDataRange().getValues();
      for (z=0;z<vals.length;z++){
        vals[z]=JSON.stringify(vals[z]);
      }
      var noDups=vals.filter((val,idx) => vals.indexOf(val)===idx);
      for (a=0;a<noDups.length;a++){
        noDups[a]=JSON.parse(noDups[a]);
      }
      var deleter=vals.length-noDups.length;
      if (deleter!=0){
        shts.getRange(1,1,noDups.length,noDups[0].length).setValues(noDups);
        shts.deleteRows(noDups.length+1,deleter);
      }
    }
  }
  
  // Highlights leads with long company names and appends uneccesary endings
  function modifyCompanyNames() {
    var app = SpreadsheetApp;
    var ui = SpreadsheetApp.getUi();
    //var sheet = SpreadsheetApp.getActiveSheet();
    var sheets = SpreadsheetApp.getActive().getSheets();
    
  
    var input = ui.prompt('Enter Column Letter with Company Names:',ui.ButtonSet.OK_CANCEL);
    if (input.getSelectedButton() == ui.Button.OK) {
      var col = letterToColumn(input.getResponseText().toUpperCase());
      sheets.forEach(sheet=>{
        var data = sheet.getDataRange().getValues();
        var i = 1;
        while (i < data.length) {
          if(containsAny(data[i][col], ['.co', '.io', '.net', 'LLC', '.com', '.ai', '.org', ', Inc'])) {
            Logger.log(data[i][col]);
            sheet.getRange(i + 1, col + 1).setValue(data[i][col].replaceAll(/.co|.io|.net|LLC|.com|.ai|.org|, Inc/gi, ""));
          }
          i++
        }
        i = 1;
        while (i < data.length) {
          if (data[i][col].length > 21) {
            sheet.getRange(i + 1, col + 1).setBackground('red');
          }
          i++
        }
      });
    }
  }
  
  function emailEditor() {
    var ui=SpreadsheetApp.getUi();
    var input=ui.prompt('Enter Column Letter with Email data or press cancel: ',ui.ButtonSet.OK_CANCEL);
  
  
    if (input.getSelectedButton()== ui.Button.OK){
    var column=letterToColumn(input.getResponseText().toUpperCase())
  
  
    var app=SpreadsheetApp;
    var sht1=app.getActiveSpreadsheet().getSheets()[0];
    var spreadsheet=app.getActiveSpreadsheet().getSheets().length;
    var row1=[sht1.getDataRange().getValues()[0]];
    var newSht=app.getActiveSpreadsheet().insertSheet('Linkedin Only',spreadsheet+1);
    newSht.getRange(1,1,1,row1[0].length).setValues(row1);
  
  
    for (i=0;i<spreadsheet;i++){
      var shts=app.getActiveSpreadsheet().getSheets()[i];
      var originalData=shts.getDataRange().getValues();
      var newDataNoEmail=originalData.filter(item => item[column]==='');
      var newDataEmail=originalData.filter(item => item[column]!='');
      if (newDataNoEmail.length!=0){
        
      }
      if (newDataNoEmail.length!=0){
        var deleter=originalData.length-newDataEmail.length;
        var begin=newSht.getDataRange().getValues().length;
        newSht.getRange(begin+1,1,newDataNoEmail.length,newDataNoEmail[0].length).setValues(newDataNoEmail);
        shts.getRange(1,1,newDataEmail.length,newDataEmail[0].length).setValues(newDataEmail);
        shts.deleteRows(newDataEmail.length+1,deleter);
      }
    }
  
    
    var ss=app.getActiveSpreadsheet();
    var sht2=ss.getSheetByName("Linkedin Only");
    if (sht2.getDataRange().getValues().length===1){
      ss.deleteSheet(sht2);
    }
    }
  }
  