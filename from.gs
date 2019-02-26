//onFormSubmit(e: Event): void
function onFormSubmit(e) {
  //get responses and define last response
  var responses = SpreadsheetApp.getActive().getSheetByName('Form Responses 1').getDataRange().getValues();
  var lastResponse = responses[responses.length-1];
  
  //parse addresses
  var rechargeAddr = parseRechargeAddress(lastResponse[2]);
  
  //copy template file, name after invoice #
  var driveFile = DriveApp.getFileById('17HrtgOxjCFSEDdxNxBYX819ygKuZOq7VTkg_839S-Zk').makeCopy(lastResponse[4],
                                                                                               DriveApp.getFolderById(
                                                                                                 '1EM0UJWFgmGxswT5hj4PX9RFedFr6BB_I'));
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
  
  var totalCost = 0;
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
      miles:lastResponse[startCol+7]
    }
    Logger.log(job.elapsedTime);
    var tableRow = table.appendTableRow();
    var charges = getCharges(job);
    var arr = [job.jobNum, job.refNum, job.date, getPickup(job), getDelivery(job), charges[0]];
    totalCost = charges[1];
    for(var j = 0; j < arr.length; ++j){
      tableRow.appendTableCell(arr[j]);
    }
    tableRow.setAttributes(jobCellStyle);
  }
  body.replaceText('{tot}', totalCost);
  
  //email user completed form
  GmailApp.sendEmail(lastResponse[1], 'Recharge form for Invoice #'+lastResponse[4], 'Please see the attached file. It can also be viewed at: \n'
                     + driveFile.getUrl(), { attachements: [driveFile.getAs(MimeType.PDF)], name: 'Automatic Recharge Script'});
}

//parseRechargeAddress(addr: Multiline String): String[]
function parseRechargeAddress(addr){
  var addrSplit = addr.split("\n");
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

//getCharges(job: Job): [String, number]
function getCharges(job){
  var hours = parseInt(job.dTime.substring(0, 2), 10)-parseInt(job.pTime.substring(0,2), 10);
  var minutes = parseInt(job.dTime.substring(3), 10)-parseInt(job.pTime.substring(3), 10);
  var totalTime = hours+(minutes/60);
  var base = "Time: \t" + hours + ":" + minutes + "\t $" + (totalTime*25).toFixed(2);
  var miles = "Miles: \t" + job.miles + "\t $" + job.miles*0.58
  var totalCost = (totalTime*25 + job.miles*0.58).toFixed(2);
  var tot = "Total: \t \t $" + totalCost.toString(10);
  return [base + "\n" + miles + "\n" + tot, totalCost];
}
