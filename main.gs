// Email must be collected for all of the form
// ------ HARD-DEFINED-VARIABLE ------ // 
var workflow_tab_name = "Workflow"
var mailing_tab_name = "Mailing list"
var blacklisted_tab_name = "Blacklisted list"
var ongoing_workflow_tab_name = "Ongoing Workflow"

// ------ APPROVAL FORM VARIABLE ------ // 
var receipt_tab_name = "Approval Receipt"
const approval_form = FormApp.openByUrl(<THE EDITABLE FORM LINK FOR THE APPROVAL RECEIPT>); // in edit format
const approval_form_id = approval_form.getId();

// -------------- AUTOMATION --------------------- //
// holding back send email
function main(){
  convertAllTimeToID()

  var unfinishedInputRequest = findUnfinishedInput();
  for(tab = 0; tab < unfinishedInputRequest.length ; tab++){ // process sheet by sheet
    if(unfinishedInputRequest[tab].length >= 2){
      var tabName = unfinishedInputRequest[tab][0]
      target_tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName)
      for(i = 1 ; i < unfinishedInputRequest[tab].length ; i++){
        row = unfinishedInputRequest[tab][i]
        
        // check whether the requester is blacklisted
        if(isEmailBlacklisted(getCellValue(tabName,row,2))){
          console.log("Blacklisted email")
          updateRowColor(tabName,row,'#FF0000')
        }
        else{
          console.log("Non-blacklisted email")
          var ongoing_workflow_row = newWorkflow(tabName,getCellValue(tabName,row,1))
          workflow(ongoing_workflow_row,tabName,row)
        }
      }
    }
  }

  var unfinishedReceiptRequest = findUnfinishedReceipt();
  if(unfinishedReceiptRequest.length >= 2){
    tabName = unfinishedReceiptRequest[0]
    target_tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName)
      for(i = 1 ; i < unfinishedReceiptRequest.length ; i++){
        row = unfinishedReceiptRequest[i]
        
        // check whether the requester is blacklisted
        if(isEmailBlacklisted(getCellValue(tabName,row,2))){
          console.log("Blacklisted email")
          updateRowColor(tabName,row,'#FF0000')
        }
        else{
          console.log("Non-blacklisted email")
          var input_time_id = getCellValue(tabName,row,3)
          var receipt_email = getCellValue(tabName,row,2)
          var receipt_id = getCellValue(tabName,row,1)

          console.log(input_time_id,receipt_email,receipt_id)

          if(checkReceiptEmail(input_time_id,receipt_email)){
            console.log("Proper receipt")
            var ongoing_workflow_row = updateWorkflow(input_time_id,receipt_id) 
            workflow(ongoing_workflow_row,tabName,row)
          }
          else{
            console.log("faking signature or unrecognised time_id")
            updateRowColor(tabName,row,'#FF0000') // faking signature
          }
        }
      }
    }

}

function workflow(ongoing_workflow_row,original_request_tab_name,original_request_tab_row){
  var row = ongoing_workflow_row
  var summary = collateInputAndReceipt(getCellValue(ongoing_workflow_tab_name,row,3))
  var notify_email_list
  var current_status

  console.log(row,summary,notify_email_list,current_status)

  var input_id = getCellValue(ongoing_workflow_tab_name,row,3) // check input ID
  var input_type = getCellValue(ongoing_workflow_tab_name,row,1) // check stage
  var stage = getCellValue(ongoing_workflow_tab_name,row,2) // check stage

  var requester_email = getRequesterEmailOfInputID(input_id,input_type)
  
  console.log(input_id,input_type,stage,requester_email)

  if(stage == 0){
    notify_email_list = checkStageType_emailList(row,"Init");
    current_status = "INITIALISE WORKFLOW"
  }
  else{
    console.log("Searching for: ",getColumnAndIDOfLastRequest(row).ID)
    var last_receipt_time_id = getColumnAndIDOfLastRequest(row).ID
    var last_receipt_row = getRowOfReceiptID(last_receipt_time_id)
    console.log(last_receipt_row)
    var status = getCellValue(receipt_tab_name,last_receipt_row,5)
    console.log(status)

    if(status == "Approved"){
      notify_email_list = checkStageType_emailList(row,status);
      if(notify_email_list.Gmail[0]!=""){
        current_status = "Approved by previous PIC, proceeding to next PIC"
      }
      else{
        notify_email_list = type_fullEmailList(getCellValue(ongoing_workflow_tab_name,row,1))
        current_status = "FULLY APPROVED"
      }
    }
    else if(status == "Pending"){
      current_status = "PENDING"
      notify_email_list = checkStageType_emailList(row,status);
    }
    else{
      current_status = "REJECTED"
      notify_email_list = checkStageType_emailList(row,status);
    }
  }

  console.log(summary)
  console.log(notify_email_list)

  var email_count = notify_email_list.Gmail.length
  if (checkQuota(email_count)){ // CHECK QUOTA COUNT 
    if(stage == 0 || status == "Approved"){
      increaseStage(ongoing_workflow_row)
    }
    
    var html_body = generateHTMLBody(current_status,summary,requester_email,notify_email_list,input_id,input_type) 
    var cc_list = notify_email_list.Gmail.concat(notify_email_list.School_Email)
    var cc_list = cc_list.concat([requester_email])
    var cc_list = cc_list.toString()
    var email_topic = input_type + " : " + input_id
    console.log(notify_email_list.Gmail[0])
    console.log(cc_list) 
    console.log(email_topic)
    console.log(html_body)
    sendEmail(notify_email_list.Gmail[0],cc_list,email_topic,html_body)

    // colour marking
    updateRowColor(original_request_tab_name,original_request_tab_row,'#00FF00')  // update the request tab name
    
    if(current_status == "FULLY APPROVED"){ // update the ongoing tab
      updateRowColor(ongoing_workflow_tab_name,ongoing_workflow_row,'#00FF00')
    }
    else if(current_status == "REJECTED"){
      updateRowColor(ongoing_workflow_tab_name,ongoing_workflow_row,'#FF0000')
    }
    else{
      updateRowColor(ongoing_workflow_tab_name,ongoing_workflow_row,'#0000FF')
    }
  }
  else {
    console.log("EXCEED QUOTA")
    revertWorkflow(ongoing_workflow_row)
  }
}


function convertAllTimeToID(){
  console.log("CHANGE THE TIME TO ID")
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var workflowSheet = ss.getSheetByName(workflow_tab_name);
  var workflowData = workflowSheet.getRange("A2:A").getValues(); // Assuming data starts from the first row in column A
  var tabName = [receipt_tab_name]

  var temp = 0
  while(workflowData[temp]!=""){
    tabName.push(workflowData[temp][0])
    temp++
  }
  
  console.log(tabName)
  for(var tab = 0; tab < tabName.length ; tab++){

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var workflowSheet = ss.getSheetByName(tabName[tab]);
    var workflowData = workflowSheet.getRange("A2:A").getValues();
    var time = []
    var time_check = []
    var updated_time = []
    var temp = 0
    
    while(workflowData[temp]!=""){
      time.push(workflowData[temp][0])
      temp++
    }
    console.log(time)
    
    for(var time_ID = 0; time_ID < time.length ; time_ID++){
      time_check.push(isValidDate(time[time_ID]))
    }

    for(var check = 0 ; check < time_check.length ; check++){
      if(time_check[check]==true){
        updated_time.push([timeID(time[check])])
      }
      else{
        updated_time.push([time[check]])
      }
    }
    console.log(updated_time)

    workflowSheet.getRange(2,1,updated_time.length,1).setValues(updated_time)

  }
}

function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

function timeID(time){
  var temp = new Date(time).valueOf()
  temp = temp.toString()
  return temp
}


function checkStageType_emailList(ongoing_workflow_row,status){
  var row = ongoing_workflow_row
  var type = getCellValue(ongoing_workflow_tab_name,row,1)
  var stage = getCellValue(ongoing_workflow_tab_name,row,2)

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var workflowSheet = ss.getSheetByName(workflow_tab_name);
  var workflowData = workflowSheet.getDataRange(); 
  var data = workflowData.getValues()

  if(status == "Init" || status == "Approved"){
    stage += 1 ; // offsetting
  }

  var notify_positions = ""

  for (var i = 0 ; i < data.length ; i++){
    //console.log(data[i][0])
    if(data[i][0]==type){
      notify_positions = getCellValue(workflow_tab_name,i+1,stage+1) // +1 to shift right cause or the numbering system
      break
    }
  }

  return processEmailList(notify_positions)
}

function increaseStage(ongoing_workflow_row){
  var stage = getCellValue(ongoing_workflow_tab_name,ongoing_workflow_row,2)
  updateCellValue(ongoing_workflow_tab_name,ongoing_workflow_row,2,stage+1)// increment stage
}

function type_fullEmailList(type){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(workflow_tab_name);
  var data = sheet.getDataRange().getValues();

  // Search for the input in existing data
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === type) {
      var last = 2;
      while(getCellValue(workflow_tab_name,i+1,last+1)!=""){last+=1}
      return processEmailList(getCellValue(workflow_tab_name,i+1,last))
    }
  }
  return false // didn't exist
}

function findUnfinishedInput() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var workflowSheet = ss.getSheetByName(workflow_tab_name);
  var workflowData = workflowSheet.getRange("A:A").getValues(); // Assuming data starts from the first row in column A
  var allUnfinishedRow = []

  for (var i = 1; i < workflowData.length; i++) {
    var tabName = workflowData[i][0]; // Assuming each cell contains the tab name
    var targetSheet = ss.getSheetByName(tabName);

    if (targetSheet) {
      var temp = [tabName]
      var dataRange = targetSheet.getDataRange();
      var data = dataRange.getValues();
      var backgroundColors = dataRange.getBackgrounds();
      
      for (var row = 1; row < data.length; row++) {
        var isGreen = backgroundColors[row][0] === '#00ff00'; // Check if the first cell's background color is green
        var isRed = backgroundColors[row][0] === '#ff0000'; // Check if the first cell's background color is red
        if (!(isGreen || isRed)) { temp.push(row+1)}
      }
      allUnfinishedRow.push(temp)
    }
  }

  return allUnfinishedRow
}

function findUnfinishedReceipt(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  targetSheet = ss.getSheetByName(receipt_tab_name)
  if (targetSheet) {
    var temp = [receipt_tab_name]
    var dataRange = targetSheet.getDataRange();
    var data = dataRange.getValues();
    var backgroundColors = dataRange.getBackgrounds();
      
    for (var row = 1; row < data.length; row++) {
      var isGreen = backgroundColors[row][0] === '#00ff00'; // Check if the first cell's background color is green
      var isRed = backgroundColors[row][0] === '#ff0000'; // Check if the first cell's background color is red
      if (!(isGreen || isRed)) { temp.push(row+1)}
    }
    return temp
  }
}

function isEmailBlacklisted(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var blacklistSheet = ss.getSheetByName(blacklisted_tab_name);

  var emailToCheck = email.toLowerCase().trim();
  var blacklistEmails = blacklistSheet.getRange('A:A').getValues();

  for (var i = 0; i < blacklistEmails.length; i++) {
    var blacklistedEmail = blacklistEmails[i][0].toString().toLowerCase().trim();
    if (emailToCheck === blacklistedEmail) {
      return true;
    }
  }
  return false;
}

function newWorkflow(type,input){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName(ongoing_workflow_tab_name);
  var newRow = [type,0,String(input)];
  sheet.appendRow(newRow)

  var sheetData = sheet.getDataRange();
  var data = sheetData.getValues();
  return data.length
}

function revertWorkflow(ongoing_workflow_row){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ongoing_workflow_tab_name);
  
  var stage = getCellValue(ongoing_workflow_tab_name,ongoing_workflow_row,2)

  if(stage == 0){ // just initialise, not yet send out email cause no quota
    sheet.deleteRow(ongoing_workflow_row)
  }
  else{ // receipt, not yet send out email to next step
    var last_receipt_column = getColumnAndIDOfLastRequest(ongoing_workflow_row).Column
    sheet.getRange(ongoing_workflow_row,last_receipt_column).setValue("")
  }
}

function checkReceiptEmail(input_time_id,email){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ongoing_workflow_tab_name);
  var data = sheet.getDataRange().getValues();
  
  // Search for the input in existing data
  for (var i = 0; i < data.length; i++) {
    if (data[i][2] === input_time_id) {
      var type = data[i][0] // check the type 
      var stage = data[i][1] // check the stage
      console.log(data[i][2],type,stage)
      if(getTypeStage_PICEmail(type,stage) == email){
        return true
      }
      else{
        return false
      }
    }
  }
  return false // didn't exist
}

function sendEmail(email, cc_email, email_subject, message) {

    //MailApp.sendEmail({
    //  to: email,
    //  cc: cc_email,
    //  subject: email_subject,
    //  htmlBody: message,
    //});

    GmailApp.sendEmail(email,email_subject,email_subject,
      {
        cc: cc_email,
        htmlBody: message
      }
    )
}

function generateHTMLBody(status,question_answer_array,requester_email,email_list,input_id,input_type){
  var html_body = " <p> Current status of the "+ input_type +" workflow: " + status + "</p>"
  if(status == "FULLY APPROVED" || status == "REJECTED"){
    html_body += 
    " <p> Next action item by REQUESTER: " + " ( "+ requester_email +" ) </p>"
  }
  else{
    html_body += 
    " <p> Next action item by: " + email_list.Designation[0] +" - " + email_list.Name[0] + " ( "+ email_list.Gmail[0] +" ) </p>" + 
    " <p> Please use <a href="+ approval_form_pre_filled(input_id,input_type) + ">THIS </a> link to respond to the request </p>"
  }

  html_body += "<p> DO NOT FAKE SIGNATURE OR RESPOND AS YOUR EMAIL WILL BE BLACKLISTED AND WILL BE FACING DISCIPLINE ACTIONS </p>"
  
  for(var i = 0;i < question_answer_array.length ; i++){
    if(question_answer_array[i]!=null){
      html_body += generateHTMLTable(question_answer_array[i])
    }
  }

  html_body += "<p> MLDA@EEE </p>"
  html_body += "<p> Learn📚, Apply📈, Connect🌟 </p>"

  return html_body
}

function generateHTMLTable(question_answer){ 
  //This would be the header of the table
  var table = "<table border=1><tr><th>Question</th><th>Answer</th></br>";

  for (var i = 0; i < question_answer.question.length; i++){
    table = table + "<tr>"+"<td>"+ question_answer.question[i] +"<td>"+ question_answer.answer[i] +"</td>"+"</tr>";
  }

  table=table+"</table>";
  return table
}

function processEmailList(positions) {
  var positions_array = positions.split(',').map(function (email) {
    return email.trim().replace(/^0+/, ''); // Remove leading zeros from each email
  });
  var Designation = []
  var Name = []
  var Gmail = []
  var School_Email = []
  for (var i = 0; i < positions_array.length; i++) {
    var NameAndEmails = getNameAndEmailByRole(positions_array[i])
    Designation.push(NameAndEmails[0])
    Name.push(NameAndEmails[1])
    Gmail.push(NameAndEmails[2])
    School_Email.push(NameAndEmails[3])
  }
  return { Designation: Designation, Name: Name , Gmail: Gmail , School_Email: School_Email };
}

function collateInputAndReceipt(time_id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ongoing_workflow_tab_name);
  var data = sheet.getDataRange().getValues();
  var id_list = []
  var return_responses = []
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][2] === time_id) {
      id_list = data[i];
    }
  }
  
  return_responses.push(collateQuestionAndAnswer(id_list[0],id_list[2]))

  if(id_list.length > 3){
    for (var i = 3 ; i < id_list.length ; i++){ // ignore first value as it is the type, second value as it is the stage
      return_responses.push(collateQuestionAndAnswer(receipt_tab_name,id_list[i]))
    }
  }
  
  return return_responses
}

function collateQuestionAndAnswer(tabName, time_id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tabName);
  
  if (!sheet) {
    Logger.log("Sheet '" + tabName + "' not found.");
    return null;
  }
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { // Start from the 2nd row
    if (data[i][0] === time_id) {
      var question = data[0]; // Question is in the 1st row
      var answer = data[i]; // Answer is in the respective row
      return { question: question, answer: answer };
    }
  }

  return null;
}

function approval_form_pre_filled(input_id,input_type) {
  var form = approval_form;
  var items = form.getItems();
  
  var formResponse = form.createResponse();

  var formItem = items[0].asTextItem();
  var response = formItem.createResponse(input_id);
  formResponse.withItemResponse(response);

  var formItem = items[1].asTextItem();
  var response = formItem.createResponse(input_type);
  formResponse.withItemResponse(response);

  // Get prefilled form URL
  var url = formResponse.toPrefilledUrl();
  return url
};

function getNameAndEmailByRole(role) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rolesSheet = ss.getSheetByName(mailing_tab_name); 

  var rolesData = rolesSheet.getDataRange().getValues();
  var NameAndEmails = []

  for (var i = 0; i < rolesData.length; i++) {
    var roleInSheet = rolesData[i][0].toString().trim();
    if (roleInSheet.toLowerCase() === role.toLowerCase()) {
      NameAndEmails=rolesData[i];
      break;
    }
  }
  return NameAndEmails
}

function getCellValue(sheetName, row, column) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" not found.');
  }
  var cellValue = sheet.getRange(row, column).getValue();
  return cellValue;
}

function getTypeStage_PICEmail(type,stage){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(workflow_tab_name);
  var data = sheet.getDataRange().getValues();

  for (var i = 0; i < data.length; i++) {
    if(data[i][0] == type){
      var hold = processEmailList(data[i][stage]).Gmail
      return hold[0]
    }
  }
  return ""
}

function getRequesterEmailOfInputID(input_time_id,input_type){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(input_type);
  var data = sheet.getDataRange().getValues();

  // Search for the input in existing data
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === input_time_id) {
      return data[i][1] 
    }
  }
  return false // didn't exist
}

function getRowOfReceiptID(time_id){
  var row = 2;
  while(getCellValue(receipt_tab_name,row,1)!=time_id){row+=1}
  return row
}

function getColumnAndIDOfLastRequest(ongoing_workflow_row){
  var row = ongoing_workflow_row
  var last = 3;
  while(getCellValue(ongoing_workflow_tab_name,row,last+1)!=""){last+=1}
  return {Column:last , ID:getCellValue(ongoing_workflow_tab_name,row,last)}
}

function updateRowColor(sheetName, row, color) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log("Sheet '" + sheetName + "' not found.");
    return;
  }
  
  var dataRange = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  dataRange.setBackground(color);
}

function updateCellValue(sheetName,row, column, newValue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" not found.');
  }

  sheet.getRange(row, column).setValue(newValue);
}

function updateWorkflow(input_time_id, receipt) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ongoing_workflow_tab_name);
  
  var data = sheet.getDataRange().getValues();
  
  // Search for the input in existing data
  for (var i = 0; i < data.length; i++) {
    if (data[i][2] === input_time_id) {
      var newColumn = 3
      while(getCellValue(ongoing_workflow_tab_name,i+1,newColumn)!=""){ newColumn+=1}
      sheet.getRange(i + 1, newColumn).setValue(String(receipt));
      return i + 1
    }
  }
  return false // not an existing workflow
}

function checkQuota(email_count){
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
  if(email_count <= emailQuotaRemaining){
    return true
  }
  else {
    return false
  }
}

function testQuota(){
  checkQuota(0);
}
