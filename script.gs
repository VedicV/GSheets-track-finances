const ss = SpreadsheetApp.getActiveSpreadsheet();
const mainSheet = ss.getSheetByName("Connection");

// todo сделать чтобы брался только если недействительный и сохранять в табл
function getToken() {

    const userid = mainSheet.getRange("B25").getValue();
    const userkey = mainSheet.getRange("B28").getValue();

    let raw = JSON.stringify({ "secret_id": userid, "secret_key": userkey });
    let myHeaders = { "accept": "application/json", "Content-Type": "application/json" }

    let requestOptions = {
        'method': 'POST',
        'headers': myHeaders,
        'payload': raw,
    };

    let response = UrlFetchApp.fetch("https://ob.nordigen.com/api/v2/token/new/", requestOptions);
    let json = response.getContentText();
    let token = JSON.parse(json).access;

    return token
}

function getBanks() {

    mainSheet.getRange("J1:J").clear();

    let data = getBanksList();
    let bankList = data.map(bank => [bank.name]);

    mainSheet.getRange(1, 10).setValues(bankList);
}

function getBanksList() {
    const token = getToken();
    let country = mainSheet.getRange("B34").getValue();

    let url = "https://ob.nordigen.com/api/v2/institutions/?country=" + country;
    let headers = {
        "headers": {
            "accept": "application/json",
            "Authorization": "Bearer " + token
        }
    };

    let response = UrlFetchApp.fetch(url, headers);
    let json = response.getContentText();
    let data = JSON.parse(json);
    return data
}

function createLink() {

    // todo переделать чтобы не надо было 2й раз получать список банков - id в таблицу писать сразу
    // get Bank id 
    let bank = mainSheet.getRange("B43").getValue();
    let data = getBanksList();
    for (let j in data) {
        if (data[j].name == bank)
            var institution_id = data[j].id;
    }

    // get requisition ID
    const token = getToken();
    let myHeaders = {
        "accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token
    }

    let sheetID = mainSheet.getSheetId();
    let redirect_link = ss.getUrl() + '#gid=' + sheetID;

    let raw = JSON.stringify({ "redirect": redirect_link, "institution_id": institution_id });

    let requestOptions = {
        'method': 'POST',
        'headers': myHeaders,
        'payload': raw
    };

    let response = UrlFetchApp.fetch("https://ob.nordigen.com/api/v2/requisitions/", requestOptions);
    let json = JSON.parse(response.getContentText());

    let requisition_id = json.id;
    let link = json.link;

    mainSheet.getRange(53, 2).setValue(link);
    mainSheet.getRange(1, 12).setValue(requisition_id);

}

function getTransactions() {

    let transactionsSheet = ss.getSheetByName("Transactions");

    transactionsSheet.getRange("A2:A1000").clearContent();
    transactionsSheet.getRange("B2:B1000").clearContent();
    transactionsSheet.getRange("C2:C1000").clearContent();

    const token = getToken();

    // get accounts
    let requisition_id = mainSheet.getRange("L1").getValue();

    let urlReq = "https://ob.nordigen.com/api/v2/requisitions/" + requisition_id + "/";
    let requestOptions = {
        "headers": {
            "accept": "application/json",
            "Authorization": "Bearer " + token
        }
    };

    let response = UrlFetchApp.fetch(urlReq, requestOptions);
    let json = response.getContentText();
    let accounts = JSON.parse(json).accounts;

    // get transactions
    let dataTable = [];

    for (let i in accounts) {

        let account_id = accounts[i];
        let urlAcconts = "https://ob.nordigen.com/api/v2/accounts/" + account_id + "/transactions/";

        let response = UrlFetchApp.fetch(urlAcconts, requestOptions);
        let json = response.getContentText();
        let transactions = JSON.parse(json).transactions.booked;

        for (let j in transactions) {
            let row = transactions[j];

            dataTable.push([
                row.bookingDate || '',
                row.creditorName || row.debitorName || row.remittanceInformationUnstructured || row.remittanceInformationUnstructuredArray || '',
                row.transactionAmount.amount || '',
            ])

        }
    }
    transactionsSheet.getRange(2, 1, dataTable.length, dataTable[0].length).setValues(dataTable);
}
