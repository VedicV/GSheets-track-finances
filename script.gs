function getBanks() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Connection");

  // get token 
  var userid = mainSheet.getRange("B25").getValue();
  var userkey = mainSheet.getRange("B28").getValue();

  var raw = JSON.stringify({"secret_id":userid,"secret_key":userkey});
  var myHeaders = {"accept": "application/json",
                   "Content-Type": "application/json"}

  var requestOptions = {
    'method': 'POST',
    'headers': myHeaders,
    'payload': raw
  };

  var response = UrlFetchApp.fetch("https://ob.nordigen.com/api/v2/token/new/", requestOptions);
  var json = response.getContentText();
  var token = JSON.parse(json).access;

  // get banks
  mainSheet.getRange("J1:J1000").clear();
  var country = mainSheet.getRange("B34").getValue();

  var url = "https://ob.nordigen.com/api/v2/institutions/?country="+country;
  var headers = {
             "headers":{"accept": "application/json",
                        "Authorization": "Bearer " + token}
             };

  var response = UrlFetchApp.fetch(url, headers);
  var json = response.getContentText();
  var data = JSON.parse(json);

  for (var i in data) {
  mainSheet.getRange(Number(i)+1,10).setValue([data[i].name]);
  }
  
}

function createLink() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Connection");

  // get token 
  var userid = mainSheet.getRange("B25").getValue();
  var userkey = mainSheet.getRange("B28").getValue();

  var raw = JSON.stringify({"secret_id":userid,"secret_key":userkey});
  var myHeaders = {"accept": "application/json",
                   "Content-Type": "application/json"}

  var requestOptions = {
    'method': 'POST',
    'headers': myHeaders,
    'payload': raw
  };

  var response = UrlFetchApp.fetch("https://ob.nordigen.com/api/v2/token/new/", requestOptions);
  var json = response.getContentText();
  var token = JSON.parse(json).access;

  // create link

  var bank = mainSheet.getRange("B43").getValue();
  var country = mainSheet.getRange("B34").getValue();

  var url = "https://ob.nordigen.com/api/v2/institutions/?country="+country;
  var headers = {
             "headers":{"accept": "application/json",
                        "Authorization": "Bearer " + token}
             };

  var response = UrlFetchApp.fetch(url, headers);
  var json = response.getContentText();
  var data = JSON.parse(json);

  for (var j in data) {
    if (data[j].name == bank) {
      var institution_id = data[j].id;
    }
  }

  var myHeaders = {"accept": "application/json",
                   "Content-Type": "application/json",
                    "Authorization": "Bearer " + token}

  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SS.getActiveSheet();
  var redirect_link = '';
  redirect_link += SS.getUrl();
  redirect_link += '#gid=';
  redirect_link += ss.getSheetId(); 

  var raw = JSON.stringify({"redirect":redirect_link, "institution_id":institution_id});
  var type = "application/json";

  var requestOptions = {
    'method': 'POST',
    'headers': myHeaders,
    'payload': raw
  };

  var response = UrlFetchApp.fetch("https://ob.nordigen.com/api/v2/requisitions/", requestOptions);
  var json = response.getContentText();
  var requisition_id = JSON.parse(json).id;

  var myHeaders = {"accept": "application/json",
                   "Content-Type": "application/json",
                    "Authorization": "Bearer " + token}

  var json = response.getContentText();

  var link = JSON.parse(json).link;

  mainSheet.getRange(53,2).setValue([link]);
  mainSheet.getRange(1,12).setValue([requisition_id]);
  
}

function getTransactions() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Connection");
  var transactionsSheet = ss.getSheetByName("Transactions");

  transactionsSheet.getRange("A2:A1000").clearContent();
  transactionsSheet.getRange("B2:B1000").clearContent();
  transactionsSheet.getRange("C2:C1000").clearContent();

  // get token 
  var userid = mainSheet.getRange("B25").getValue();
  var userkey = mainSheet.getRange("B28").getValue();

  var raw = JSON.stringify({"secret_id":userid,"secret_key":userkey});
  var myHeaders = {"accept": "application/json",
                   "Content-Type": "application/json"}

  var requestOptions = {
    'method': 'POST',
    'headers': myHeaders,
    'payload': raw
  };

  var response = UrlFetchApp.fetch("https://ob.nordigen.com/api/v2/token/new/", requestOptions);
  var json = response.getContentText();
  var token = JSON.parse(json).access;

  // get transactions

  var requisition_id = mainSheet.getRange("L1").getValue();

  var url = "https://ob.nordigen.com/api/v2/requisitions/" + requisition_id + "/";
  var headers = {
             "headers":{"accept": "application/json",
                        "Authorization": "Bearer " + token}
             };

  var response = UrlFetchApp.fetch(url, headers);
  var json = response.getContentText();
  var accounts = JSON.parse(json).accounts;
  
  row_counter = 2

  for (var i in accounts) {

      var account_id = accounts[i]

      var url = "https://ob.nordigen.com/api/v2/accounts/" + account_id + "/transactions/";
      var headers = {
                "headers":{"accept": "application/json",
                            "Authorization": "Bearer " + token}
                };

      var response = UrlFetchApp.fetch(url, headers);
      var json = response.getContentText();
      var transactions = JSON.parse(json).transactions.booked;

      for (var i in transactions) {

        transactionsSheet.getRange(row_counter,1).setValue([transactions[i].bookingDate]);

        if (transactions[i].creditorName) {
            var trx_text = transactions[i].creditorName
        } 
        else if (transactions[i].debitorName) {
            var trx_text = transactions[i].debitorName
        } 
        else if (transactions[i].remittanceInformationUnstructured) {
            var trx_text = transactions[i].remittanceInformationUnstructured
        } 
        else if (transactions[i].remittanceInformationUnstructuredArray) {
            var trx_text = transactions[i].remittanceInformationUnstructuredArray
        } else {
          var trx_text = ""
        }
        
        transactionsSheet.getRange(row_counter,2).setValue([trx_text]);
        transactionsSheet.getRange(row_counter,3).setValue([transactions[i].transactionAmount.amount]);

        row_counter += 1
  }

  }
    
}
