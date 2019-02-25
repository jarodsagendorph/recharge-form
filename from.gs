//onFormSubmit(e: Event): void
function onFormSubmit(e) {
  //get responses and define last response
  var responses = SpreadsheetApp.getActive().getSheetByName('Form Responses 1').getDataRange().getValues();
  var lastResponse = responses[responses.length-1];
  
  //parse addresses
  var rechargeAddr = parseRechargeAddress(lastResponse[2]);
  
  //copy template file, name after invoice #
  var driveFile = DriveApp.getFileById('1k9qX80hxzHyIiVXlhNN_TPjk63-FDXNcJOm3sqpHVJk').makeCopy(lastResponse[4]);
  var body = DocumentApp.openByUrl(driveFile.getUrl()).getBody();
  
  //set preliminary header
  body.replaceText('{adr1}', rechargeAddr[0]);
  body.replaceText('{adr2}', rechargeAddr[1]);
  body.replaceText('{adr3}', rechargeAddr[2]);
  body.replaceText('{adr4}', rechargeAddr[3]);
  body.replaceText('{adr5}', rechargeAddr[4]);
  body.replaceText('{acct}', lastResponse[3]);
  body.replaceText('{inv}', lastResponse[4]);
  body.replaceText('{date}', parseDate(lastResponse[5]));
  body.replaceText('{njobs}', lastResponse[6]);
  body.replaceText('{desc}', lastResponse[39]);
  
  var table = body.getTables()[0];
  var jobCellStyle = {}
  jobCellStyle[DocumentApp.Attribute.FONT_SIZE] = 6;
  jobCellStyle[DocumentApp.Attribute.BOLD] = false;
  jobCellStyle[DocumentApp.Attribute.BORDER_WIDTH] = 0;
  //Defines Job object and appends table                               
  for(var i = 0; i < lastResponse[6]; ++i){
    var startCol = 31 - 8*i
    var job = {
      jobNum:lastResponse[startCol],
      refNum:lastResponse[startCol+1],
      date:parseDate(lastResponse[startCol+2]),
      pAdd:lastResponse[startCol+3],
      pTime:parseTime(lastResponse[startCol+4]),
      dTime:parseTime(lastResponse[startCol+5]),
      dAdd:lastResponse[startCol+6],
      miles:lastResponse[startCol+7],
      elapsedTime:parseTime(lastResponse[startCol+5]-lastResponse[startCol+4])
    }
    var tableRow = table.appendTableRow();
    tableRow.setAttributes(jobCellStyle);
    Logger.log(tableRow.getNumCells());
    var arr = [job.jobNum, job.refNum, job.date, getPickup(job), getDelivery(job), getCharges(job)];
    
    for(var j = 0; j < arr.length; ++j){
      tableRow.appendTableCell(arr[j]);
    }
  }
  
}

//parseRechargeAddress(addr: Multiline String): String[]
function parseRechargeAddress(addr){
  var addrSplit = addr.split("\n");
  Logger.log(addrSplit);
  //normalizes array  to 5 lines, replaces excess {addr} lines with " "
  while(addrSplit.length < 5){
    addrSplit.push(" ");
  }
  return addrSplit;
}

//parseDate(date: Object): String
function parseDate(date){
  return date.toString().substring(4, 15);
}


//Input timestamp format: Sat Dec 30 1899 12:15:00 GMT-0500 (EST)
//parseTime(time: Object): String
function parseTime(time){
  return time.toString().substring(16, 21);
}

//getPickup(job: Job): String
function getPickup(job){
  var firstLine = "Ready: " + job.date + " " + job.pTime;
  var lastLine = "Clbk: " + job.date + " " + job.dTime;
  return firstLine + "\n" + job.pAdd + "\n" + lastLine;
}

//getDelivery(job: Job): String
function getDelivery(job){
  var firstLine = "Deadln: " + job.date + " " + job.dTime;
  return firstLine + "\n" + job.dAdd;
}

//getCharges(job: Job): String
function getCharges(job){
  var base = "Base: \t" + job.elapsedTime;
  var miles = "Miles: \t" + job.miles + "\t + $" + job.miles*0.58
  return base + "\n" + miles;
}
