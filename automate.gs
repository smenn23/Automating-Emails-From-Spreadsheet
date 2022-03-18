//adds a scripting menu
function setup() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Scripts").addItem('Send Emails','sendMail').addToUi();
}

function sendMail() {
  var ui = SpreadsheetApp.getUi();
  let monthNames = [,,];
  monthNames[0] = ui.prompt("Enter first month").getResponseText(); //ask inputs for months
  monthNames[1] = ui.prompt("Enter second month").getResponseText();
  monthNames[2] = ui.prompt("Enter third month").getResponseText();
  let monthcol = [,,];
  let successmonths = 0; //variable to check input months are there
  for(let i = 1; i <= SpreadsheetApp.getActiveSheet().getLastColumn(); i++) {
    currCell =  SpreadsheetApp.getActiveSheet().getRange(1,i).getValue();
    if(currCell == monthNames[0]) {
      monthcol[0] = i;
      successmonths++;
    }
    if(currCell == monthNames[1]) {
      monthcol[1] = i;
      successmonths++;
    }
    if(currCell == monthNames[2]) {
      monthcol[2] = i;
      successmonths++;
    }
  }
  if (successmonths != 3) {
    ui.alert("At least one of the months were not found, please try again with 3 months found on spreadsheet.")
    return;
  }
  //goes down list of names and sends email 
  for(let i = 3; SpreadsheetApp.getActiveSheet().getRange(i,1).getValue() != ''; i++) {
    currRow = SpreadsheetApp.getActiveSheet().getRange(i,1,1,SpreadsheetApp.getActiveSheet().getLastColumn());
    var htmlBody = getEmailHtml(currRow,monthcol,monthNames);
    MailApp.sendEmail("sethm923@gmail.com","Hours Confirmation", 'failure', {htmlBody: htmlBody});
  }
  
}
//this function fills all the necesary variables to the html template
function getEmailHtml(row,col,monthNames) {
  
  var htmlTemplate = HtmlService.createTemplateFromFile("template.html");
  htmlTemplate.name = row.getCell(1,1).getValue();
  htmlTemplate.totalHours = row.getCell(1,36).getValue();
  htmlTemplate.monthNames = monthNames;
  //these arrays will be the table entries for the respective months
  htmlTemplate.month1 = [row.getCell(1,col[0]).getValue(),row.getCell(1,col[0]+2).getValue(),row.getCell(1,col[0]+1
      ).getValue(),row.getCell(1,col[0]+3).getValue(),row.getCell(1,col[0]+5).getValue(),row.getCell(1,col[0]+6).getValue()];
  htmlTemplate.month2 = [row.getCell(1,col[1]).getValue(),row.getCell(1,col[1]+2).getValue(),row.getCell(1,col[1]+1
      ).getValue(),row.getCell(1,col[1]+3).getValue(),row.getCell(1,col[1]+5).getValue(),row.getCell(1,col[1]+6).getValue()];
  htmlTemplate.month3 = [row.getCell(1,col[2]).getValue(),row.getCell(1,col[2]+2).getValue(),row.getCell(1,col[2]+1
      ).getValue(),row.getCell(1,col[2]+3).getValue(),row.getCell(1,col[2]+5).getValue(),row.getCell(1,col[2]+6).getValue()];
  
  //changing any empty string values to '-' or '0'
  for(let i = 0; i < 6; i++) {
    if(htmlTemplate.month1[i] == '') {
      if(i < 2) {
        htmlTemplate.month1[i] = '-';
      }
      else{
         htmlTemplate.month1[i] = '0';
      }
    }
    if(htmlTemplate.month2[i] == '') {
      if(i < 2) {
        htmlTemplate.month2[i] = '-';
      }
      else{
         htmlTemplate.month2[i] = '0';
      }
    }
    if(htmlTemplate.month3[i] == '') {
      if(i < 2) {
        htmlTemplate.month3[i] = '-';
      }
      else{
         htmlTemplate.month3[i] = '0';
      }
    }
  }
  var htmlBody = htmlTemplate.evaluate().getContent();
  return htmlBody;
}
