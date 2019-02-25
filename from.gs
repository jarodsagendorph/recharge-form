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
  Logger.log(typeof lastResponse[5]);
  
  //create Job object
  jobs = [];
  for(var i = 0; i < lastResponse[6]; ++i){
    var startCol = 31 - 8*i
    jobs.push({
      jobNum:lastResponse[startCol],
      refNum:lastResponse[startCol+1],
      date:parseDate(lastResponse[startCol+2]),
      pAdd:parseShippingAddress(lastResponse[startCol+3]),
      pTime:lastResponse[startCol+4],
      dTime:lastResponse[startCol+5],
      dAdd:parseShippingAddress(lastResponse[startCol+6]),
      miles:lastResponse[startCol+7]
    });
  }
}

//parseRechargeAddress(addr: Multiline String): String[]
function parseRechargeAddress(addr){
  var addrSplit = addr.split("\n");
  Logger.log(addrSplit);
  while(addrSplit.length < 5){
    addrSplit.push(" ");
  }
  return addrSplit;
}

//parseShippingAddress(addr: Multiline String): String[]
function parseShippingAddress(addr){
  var addrSplit = addr.split("\n");
  while(addrSplit.length < 3){
    addrSplit.push(" ");
  }
  return addrSplit;
}
//parseDate(date: Object): String
function parseDate(date){
  return date.toString().substring(4, 15);
}
