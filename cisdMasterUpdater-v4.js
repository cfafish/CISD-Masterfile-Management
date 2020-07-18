 

 

function mainFunction() {

  //SpreadsheetApp.getUi().alert('Updating...');

  var folderID = "1aYnDgl0_9NLuBThTnQBkkyvCZZRugcQN";

  var MASTERID = "1IxEMlhWVpSsB58KnicKHi5IPOxVAcGOWNNatcJXlhts";

  var MASTERNAME = "Copy of Copy of Copy of CAMPUS CONTACT TRACING";

  var fileList = [];

 

  var masterSS = SpreadsheetApp.openById(MASTERID);

  var masterSheet = masterSS.getSheetByName("MASTER");

  Logger.log("mid main");

 

  //Logger.log(fileList.length);

  //Logger.log("SS name: " + masterSS.getName());

  //Logger.log("Sheet name: " + masterSheet.getName());

  //Logger.log(typeof(getsheetName(fileList[0])));

 

  var ui = SpreadsheetApp.getUi();

  var response = ui.alert('Do you want to update the master spreadsheet?', ui.ButtonSet.YES_NO);

 

  // Process the user's response.

  if (response == ui.Button.YES) {

    fileList = getfileList(folderID, MASTERNAME);

    //Logger.log('The user clicked "Yes."');

    comparingNames(fileList, masterSS, masterSheet);

    SpreadsheetApp.getUi().alert('Done!');

 

  } 

}

 

 

function comparingNames(fileList, masterSS, masterSheet){ 

  //var emptyFiles = [];

  //Logger.log(masterSheet.getName());

  masterSheet.getRange(2, 1, masterSheet.getLastRow(), masterSheet.getMaxColumns()).clearContent();

  for(var i = 0; i < fileList.length; i++){   // !!fileLoop!!

    //Logger.log(i);

    //Logger.log(fileList[i].getName());

    //var masterNames = getNamesList(masterSheet);

    var tempSS = SpreadsheetApp.openById(fileList[i].getId());

    //Logger.log(tempSS.getName());

    var tempSheet = tempSS.getSheetByName(getsheetName(fileList[i]));

    //var tempNames = getNamesList(tempSheet);

   

    var rowN = tempSheet.getLastRow() - 1;

    var colN =  masterSheet.getMaxColumns();

    //Logger.log([rowN, colN]);

    if(tempSheet.getLastRow() > 1){

      var tempRange = tempSheet.getRange(2, 1, rowN, colN);

      var values = tempRange.getValues();

      var masterRange = masterSheet.getRange(masterSheet.getLastRow()+1, 1, rowN, colN);

      masterRange.setValues(values);

   

      //Logger.log("in file " + fileList[i].getName() + " and sheet " + tempSheet.getName() + " accessing rows/cols " + rowN + " mastersheet rwos: " + masterSheet.getLastRow());

    }

  }

   

}

//var tempSS = SpreadsheetApp.openById(fileList[i].getTargetId()).getSheetByName(getsheetName(fileList[i])).getLastRow();

 

function getsheetName(ss){

  ss = SpreadsheetApp.openById(ss.getId());

  var sheetList = [];

  ss.getSheets().forEach(function(val){

       sheetList.push(val.getName())

});

//sheetList.shift();

return sheetList[0];

}

 

function getfileList(folderID, MASTERNAME){

  var folder = DriveApp.getFolderById(folderID);

  var fileList = folder.getFiles();

  //var fileArray = [];

  //var changesCol = [];

  //var changeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("changes");

  //SpreadsheetApp.getActiveSpreadsheet().getN

  //changesCol = changeSheet.getRange(1, 1, changeSheet.getLastRow(), 2).getValues();

 

  var file;

  var data = [];

 

  while(fileList.hasNext()){

   file = fileList.next();

    if((file.getMimeType().includes("spreadsheet") || file.getMimeType().includes("shortcut")) && (file.getName() != MASTERNAME)){

     

      var emptyCheckSS = SpreadsheetApp.openById(file.getId())

      var emptyCheckSheet = emptyCheckSS.getActiveSheet();

      var isEmptyN = emptyCheckSheet.getLastRow();

      var isDifSheet = emptyCheckSheet.getMaxColumns();

     

      //SheetByName(getsheetName(file)).getLastRow();

     

      if(isEmptyN > 1 && isDifSheet > 30){

        data.push(file);

      }

      else if(isDifSheet < 31){

        SpreadsheetApp.getUi().alert(emptyCheckSS.getName() + " has an extra tab that needs to be deleted.\nIt was not added to the master file.");

      }

    }

  }

  return data;

 

  /*

  var dataCol = [];

  for(var i = 0; i < fileArray.length; i++){

    file = fileArray[i];

   

    if ((file.getMimeType().includes("spreadsheet") || file.getMimeType().includes("shortcut")) && (file.getName() != MASTERNAME)){

      //Logger.log()

     

 //     for (var j = 0; j < changeSheet.getLastRow(); j++){

        //Logger.log([file.getName()]);

//        if(changesCol[j][0] == file.getName() &&  changesCol[j][1].getTime() != file.getLastUpdated().getTime()){

          //Logger.log([file.getName(), j]);

         

          //Logger.log("True");

//         data.push(file);

          //changeSheet.getRange(data.length, 3).setValue(j);

          //Logger.log(file.getName());

//       }

       

  //    }

      if(){

        data.push(file);

      }

    }

 

  }

  //PropertiesService.getScriptProperties()

  //Logger.log("shortcut id: "+data[0].getId());

  //Logger.log("target id: " + data[0].getTargetId());

  //return data;*/

}

 

/*

function initialize(){

  var folderID = "1aYnDgl0_9NLuBThTnQBkkyvCZZRugcQN";

  var folder = DriveApp.getFolderById(folderID);

  var fileList = folder.getFiles();

  var fileArray = [];

  var changeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("changes");

  //fileList.next().getName()

  while(fileList.hasNext()){

    fileArray.push(fileList.next());

    //fileList.next().getL

  }

  var file;

  var data = [];

  var nameCol = [];

  var dateCol = [];

  for(var i = 0; i < fileArray.length; i++){

    file = fileArray[i];

    if (file.getMimeType().includes("spreadsheet") || file.getMimeType().includes("shortcut")){

      //nameCol.push(file.getName());

      //dateCol.push(file.getLastUpdated());

      //if(){

        data.push(file);

      //}

      //changeSheet.appendRow([file.getName(), (file.getLastUpdated())]);

      //Logger.log(data);

    }

  }

} */

  

  

  

  

  

  

  //=========================================================================================

//v1

/*

function comparingNames(fileList, masterSS, masterSheet){

//need master ID, master ss, master sheet, target ss, target sheet, master names, target names 

  //

  //how to get target ss:

  //LOOP THROUGH FILELIST: {targetSS = SpreadsheetApp.open(fileList[i])}

  //how to get target sheet:

  //LOOP THROUGH FILELIST: {

// targetSheet = targetSS.getSheetByName(getsheetName(fileList[i]));

 

  

  for(var i = 0; i < fileList.length; i++){   // !!fileLoop!!

    var masterNames = getNamesList(masterSheet);

    var tempSS = SpreadsheetApp.open(fileList[i]);

    var tempSheet = tempSS.getSheetByName(getsheetName(fileList[i]));

    var tempNames = getNamesList(tempSheet);

   

    for(var j = 0; j < tempNames.length; j++){   // !!fileName Loop!!

      var tempIndex = j+2;

      var masterIndex = -1;

      var inMaster = false;

      var tempSet = tempNames[j];

      //Logger.log(masterNames[0] == null);

      //Logger.log(masterNames.length > 1);

      if (masterNames[0] != null){

        for(var k = 0; k < masterNames.length; k++){   // !!masterName Loop!!

          var masterTemp = masterNames[k];

          Logger.log(masterSheet.getLastRow());

          Logger.log(tempSet);

          Logger.log(masterTemp);

          if (tempSet[0] == masterTemp[0] && tempSet[1] == masterTemp[1]){

            inMaster = true;

            masterIndex = k+2;

            k = masterNames.length; //if in master then break masterLoop

          }

        }

      }

      if (!(inMaster)){

        var otherSheetRow = tempSheet.getRange(tempIndex, 1, 1, masterSheet.getMaxColumns());

        var masterSheetRow = masterSheet.getRange(masterSheet.getLastRow()+1, 1, 1, masterSheet.getMaxColumns());

        var values = otherSheetRow.getValues();

        masterSheetRow.setValues(values);

        //otherSheetRow.copyTo(masterSheetRow);

        masterIndex = masterSheet.getLastRow();

        masterNames = getNamesList(masterSheet);

      }

     

      else{

        var otherSheetRow = tempSheet.getRange(tempIndex, 1, 1, masterSheet.getMaxColumns());

        var masterSheetRow = masterSheet.getRange(masterIndex, 1, 1, masterSheet.getMaxColumns());

        var values = otherSheetRow.getValues();

        masterSheetRow.setValues(values);

        //otherSheetRow.copyTo(masterSheetRow);

       

      }

      Logger.log("from " + tempSheet.getName() + " got " + tempNames[j]+ " and put at " + masterIndex + " row in master.");

    }  

  }

}

 

function getNamesList(tempSheet){

  var tempNames = [];

  if (tempSheet.getLastRow() > 2){

    tempNames = tempSheet.getRange(2, 3, tempSheet.getLastRow()-2, 2).getValues();

  }

  else{

    tempNames = [[" "][" "]];

  }

  return tempNames;

}

  */