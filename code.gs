function onEdit(e) {
  if (!e) return;
  
  var activeSheet = e.source.getActiveSheet();
  if (activeSheet.getName() !== "Sheet4") return;
  
  var targetCell = activeSheet.getRange("A1");

  if (targetCell.getDataValidation() === null) {
    
    var sourceRange = e.source.getSheetByName("Sheet2").getRange("A2:A");
    
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();
      
    targetCell.setDataValidation(rule);
  }
}

function search(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet4 = ss.getSheetByName("Sheet4");
  let sheet2 = ss.getSheetByName("Sheet2");
  let sheet1 = ss.getSheetByName("Sheet1");


  let sheet4Val = getTable(sheet4,1,1);
  let pendingList = getTable(sheet1);
  let studentList = getTable(sheet2);

  let outputSheet = [sheet4Val[0],["","","","","","","",""],["","","","","","","",""],["","","","","","","",""],["Name", "Course", "Whatsapp", "Price", "Status", "Prepaid", "Paying(Prepaid)", "Amount prepaid"]];

  let studentOut = studentList.filter(item => item[2] == sheet4Val[0][3]);
  for(let i = 0; i < studentOut.length; i++){
    studentOut[i].push(0,"=D" + String(i+6) + "*G" + String(i+6))
  }
  outputSheet = outputSheet.concat(studentOut)
  let studentLastRow = outputSheet.length;
  outputSheet.push(["","","","","","","",""], ["Name", "Course", "Whatsapp", "Price", "Month", "Paying(1)/No(0)","",""])

  let pendingOut = pendingList.filter(item => item[2] == sheet4Val[0][3]).sort();
  for(let i = 0; i < pendingOut.length; i++){
    pendingOut[i].push(1,"","")
  }
  outputSheet = outputSheet.concat(pendingOut)
  let pendingOutLastRow = outputSheet.length;


  let diffRowLength = sheet4Val.length-outputSheet.length
  if (diffRowLength > 0){
    for(let i = 0; i < diffRowLength; i++){
      outputSheet.push(["","","","","","","",""])
    }
  }
  outputSheet[0][3] = "=VLOOKUP(A1,Table2,3,False)"
  outputSheet[0][6] = "=SUM(H6:H" + String(studentLastRow) + ")+SUMIF(F" + String(studentLastRow+3) + ":F" + String(pendingOutLastRow) + ", \"=1\", D" + String(studentLastRow+3) + ":D" + String(pendingOutLastRow) + ")";
  sheet4Val[0][7] = "=CONCATENATE(RIGHT(YEAR(NOW()),2),\"/\",TEXT(MONTH(NOW()),\"00\"))";
  writeTable(sheet4, outputSheet,1,1);
}

function save(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet4 = ss.getSheetByName("Sheet4");
  let sheet2 = ss.getSheetByName("Sheet2");
  let sheet1 = ss.getSheetByName("Sheet1");
  let log = ss.getSheetByName("Log");

  let sheet4Val = getTable(sheet4,1,1);
  if(sheet4Val.length == 1){return}
  let pendingOutList = getTable(sheet1);
  let studentOutList = getTable(sheet2);
  let studentInList = []
  let pendingInList = []
  let temp = 0
  for(let i = 5; i < sheet4Val.length; i++){
    if (sheet4Val[i][0] == ""){
      studentInList = sheet4Val.slice(5,i)
      temp = i+2;
      break;      
    }
  }
  pendingInList = sheet4Val.slice(temp, sheet4Val.length)
  let logList = [];
  let date = new Date();
  let time = Utilities.formatDate(date, "GMT+8", "HH:mm:ss");
  let today = Utilities.formatDate(date, "GMT+8", "dd/MM/yyyy");
  for(let i = 0; i < studentOutList.length; i++){
    if(studentOutList[i][2] == sheet4Val[0][3]){
      for(let j = 0; j < studentInList.length; j++){
        if(studentOutList[i][0] == studentInList[j][0] && studentOutList[i][1] == studentInList[j][1]){
          let toBeLog = studentInList[j].slice(0,7);
          toBeLog[5] = time;
          toBeLog[6] = today;
          for(let k = 1; k <= studentInList[j][6]; k++){
            let newLog = [...toBeLog];
            let dateIn = new Date();
            dateIn.setMonth(date.getMonth()+k+studentOutList[i][5]);
            newLog[4] = Utilities.formatDate(dateIn, "GMT+8", "yy/MM");
            logList.push(newLog);
          }
          studentOutList[i][5] += studentInList[j][6]
          studentInList.splice(j,1);
          break;
        }
      }
    }
  }
  let oldPendingOutLength = pendingOutList.length;
  pendingInList = pendingInList.filter(item => item[5] == 1);
  for(let i = 0; i < pendingInList.length; i++){
    let toBeLog = pendingInList[i].slice(0,7);
    toBeLog[5] = time;
    toBeLog[6] = today; 
    logList.push(toBeLog);
  }
  pendingOutList = pendingOutList.filter((item) => {
    for (let i = 0; i < pendingInList.length; i++){
      if (item[0] == pendingInList[i][0] && item[1] == pendingInList[i][1] && item[2] == pendingInList[i][2] && item[3] == pendingInList[i][3]){
        pendingInList.splice(i,1);
        return false;
      }
    }
    return true;
  });
  let diffRowLength = oldPendingOutLength-pendingOutList.length
  if (diffRowLength > 0){
    for(let i = 0; i < diffRowLength; i++){
      pendingOutList.push(["","","","",""])
    }
  }

  for(let i = 4; i < sheet4Val.length; i++){
    sheet4Val[i] = ["","","","","","","",""]
  }
  sheet4Val[0][3] = "=VLOOKUP(A1,Table2,3,False)";
  sheet4Val[0][6] = 0;
  sheet4Val[0][7] = "=CONCATENATE(RIGHT(YEAR(NOW()),2),\"/\",TEXT(MONTH(NOW()),\"00\"))";

  writeTable(sheet2, studentOutList,2,1)
  writeTable(sheet1, pendingOutList,2,1)
  writeTable(sheet4, sheet4Val, 1, 1)
  writeTable(log,logList,undefined,1)
}

function monthlyUpdate() {
  let month = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet3").getRange("B2").getValue();
  let sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet2");
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  let tableOrigin = getTable(sheet2)
  let output = []
  for(let i = 0; i < tableOrigin.length; i++){
    if (tableOrigin[i][5] != 0){
      tableOrigin[i][5]--
    }else if(tableOrigin[i][4] == "Active"){
      let temp = tableOrigin[i].slice(0,4)
      temp.push(month)
      output.push(temp)
    }
  }
  writeTable(sheet1, output)
  writeTable(sheet2, tableOrigin, 2, 1)
}

function generateStatement(){
  let pendingList = getTable(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1"));
  let pendingObject = {};
  for (let i = 0; i < pendingList.length; i++){
    let whatsapp = pendingList[i][2];
    if (pendingObject.hasOwnProperty(whatsapp)){
      pendingObject[whatsapp].push(pendingList[i])
    }else{
      pendingObject[whatsapp] = [pendingList[i]]
    }
  }
  // Logger.log(pendingObject)
  let date = Utilities.formatDate(new Date(), "GMT+8", "HH:mm dd-MM-yyyy");
  let outputString = "===== Statement at " + String(date) + " =====\n\n";
  // Logger.log(outputString)
  let months = ["January ", "February ", "March ", "April ", "May ", "June ", "July ", "August ", "September ", "October ", "November ", "December "];
  for(let i in pendingObject){
    let mf = (i.slice(-2,-1) == 'm') ? "madam": "sir";
    outputString += "===== " + String(i) + " =====\nHi " + mf +", fees for: \n"
    let s = 0;
    for(let j = 0; j < pendingObject[i].length; j++){
      outputString += "  - " + pendingObject[i][j][0] + " (" + months[parseInt(pendingObject[i][j][4].slice(3,5))-1] + pendingObject[i][j][4].slice(0,2) + ") [RM" + pendingObject[i][j][3] + "]\n";
      s += parseFloat(pendingObject[i][j][3]);
    }
    outputString += "Total: RM" + s + "\n\nThis is just a system update, if you have paid already please ignore this message.\nKindly take note, thanks.\n===========================\n\n";
  }
  let drive = DriveApp.getFilesByName("Statement Yamaha");
  if (drive.hasNext()){
    drive.next().setContent(outputString);
  }else{
    DriveApp.createFile("Statement Yamaha", outputString, "text/plain")
  }
}

function clearTable(sheetOrigin, row = 2, col = 1){
  if(sheetOrigin.getLastRow()-row+1 < 1){return}
  sheetOrigin.getRange(row,col,sheetOrigin.getLastRow()-row+1,sheetOrigin.getLastColumn()-col+1).clearContent();
}

function getTable(sheetOrigin,row = 2, col = 1){
  if(sheetOrigin.getLastRow()-row+1 < 1){return []}
  return sheetOrigin.getRange(row,col,sheetOrigin.getLastRow()-row+1,sheetOrigin.getLastColumn()-col+1).getValues();
}

function writeTable(sheetDestination, output, row = undefined, col = 1){
  if(output.length < 1 || output[0].length < 1){return}
  if (row == undefined){
    row = sheetDestination.getLastRow()+1
  }
  sheetDestination.getRange(row,col,output.length,output[0].length).setValues(output)
}
