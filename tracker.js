function onOpen(){
  var sheet = SpreadsheetApp.getActive()
  var subMenus = [{name:'Show', functionName:'show'},
                  {name:'Meeting', functionName:'meeting'},
                  {name:'Open Mic', functionName:'openMic'},
                  {name:'Other', functionName:'other'}]
  sheet.addMenu('Attendance', subMenus)
}

function show(){
  inputGTID('Show')
}
function meeting(){
  inputGTID('Meeting')
}
function openMic(){
  inputGTID('Open Mic')
}
function other(){
  inputGTID('Other')
}
  
// Prompts user for GTID input
function inputGTID(eventType){
  var sheet = SpreadsheetApp.getActive().getSheetByName('People')
  while (true){
    var inGTID = Browser.inputBox('Input GTID')
    if (inGTID == 'cancel'){
      return
    }
    var lastrow = sheet.getLastRow()
    var GTIDlist = sheet.getRange('A1:A'+lastrow).getValues()
    var len = GTIDlist.length
    var found = false
    for (var x = 0; x < len; x++) {
      if (GTIDlist[x] == inGTID){
        found = true
        recognized(x+1, eventType)
      }
    }
    if (!found){
      notRecognized(inGTID, eventType)
    }
  }
}

// If input GTID was found in the list in People tab:
function recognized(row, eventType){
  var peopleSheet = SpreadsheetApp.getActive().getSheetByName('People')
  var firstname = peopleSheet.getRange('B'+row).getValue()
  var lastname = peopleSheet.getRange('C'+row).getValue()
  Browser.msgBox('Welcome Back, ' + firstname + ' ' + lastname)
  
  if (eventType == 'Show'){
    writeToPeople(row,'E','F')
  }
  else if (eventType == 'Meeting'){
    writeToPeople(row,'G','H')
  }
  else if (eventType == 'Open Mic'){
    writeToPeople(row,'I','J')
  }
  else if (eventType == 'Other'){
    writeToPeople(row,'K','L')
  }
  writeToOtherSheets(eventType, firstname, lastname)
}

// If input GTID was not found in the list in People tab:
function notRecognized(GTID, eventType){
  var sheet = SpreadsheetApp.getActive().getSheetByName('People')
  var names = askName()
  var firstName = names[0]
  var lastName = names[1]
  var email = Browser.inputBox("Enter your email address")
  
  var lastrow = sheet.getLastRow()+1
  sheet.getRange('A'+lastrow).setValue(GTID)
  sheet.getRange('B'+lastrow).setValue(firstName)
  sheet.getRange('C'+lastrow).setValue(lastName)
  sheet.getRange('D'+lastrow).setValue(email)
  
  if (eventType == 'Show'){
    writeToPeople(lastrow,'E','F')
  }
  else if (eventType == 'Meeting'){
    writeToPeople(lastrow,'G','H')
  }
  else if (eventType == 'Open Mic'){
    writeToPeople(lastrow,'I','J')
  }
  else if (eventType == 'Other'){
    writeToPeople(lastrow,'K','L')
  }
  writeToOtherSheets(eventType, firstName, lastName)
}

// Gets user's first and last name, with format error handling
function askName(){
  var rawName = Browser.inputBox("Enter your first and last name")
  var splitNames = rawName.split(" ")
  if (splitNames.length != 2){
    var nameGood = false
    while (!nameGood){
      var newName = Browser.inputBox("Name format not recognized. Enter only your first and last name separated by a space.")
      var splitNames = newName.split(" ")
      if (splitNames.length == 2){
        nameGood = true
      }
    }
  }
  return splitNames
}

// Increments event count by 1 and appends date to list
function writeToPeople(row, numCol, listCol){
  var peopleSheet = SpreadsheetApp.getActive().getSheetByName('People')
  var curNum = peopleSheet.getRange(numCol+row).getValue()
  peopleSheet.getRange(numCol+row).setValue(curNum+1)
  var curVal = peopleSheet.getRange(listCol+row).getValue()
  var date = getDate()
  peopleSheet.getRange(listCol+row).setValue(curVal+date+"; ")
}

// Appends attendee's name to list in correct sheet and increments count
function writeToOtherSheets(eventType, firstName, lastName){
  var thisSheet = SpreadsheetApp.getActive().getSheetByName(eventType)
  var last = thisSheet.getLastRow()
  var date = getDate()
  var found = false
  var row = 0
  
  for (var x = 1; x < last+1; x++){
    var rowDate = thisSheet.getRange("A" + x).getValue()
    rowDate = String(rowDate)
    if (rowDate == date){
      row = x
      found = true
    }
  }
  
  if (!found){
    row = last+1
    thisSheet.getRange("A"+row).setValue(date)
  } 
  
  var curNames = thisSheet.getRange("D"+row).getValue()
  thisSheet.getRange("D"+row).setValue(curNames+firstName+" "+lastName+"; ")
  var curNum = thisSheet.getRange("C"+row).getValue()
  thisSheet.getRange("C"+row).setValue(curNum+1)
}

// Returns today's date as mm/dd/yyyy. Stolen from a guy on Stack Overflow.
function getDate(){
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; //January is 0!
  var yyyy = today.getFullYear();
  if(dd<10) {
    dd='0'+dd
  } 
  if(mm<10) {
    mm='0'+mm
  } 
  var dateString = mm+'/'+dd+'/'+yyyy;
  return dateString
}
