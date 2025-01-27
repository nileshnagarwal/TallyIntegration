function showDatePrompt() {
  var ui = SpreadsheetApp.getUi();
  
  var startResponse = ui.prompt('Enter the start date for filtering invoices', 'Please use the format YYYY-MM-DD', ui.ButtonSet.OK_CANCEL);
  if (startResponse.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  var startDate = new Date(startResponse.getResponseText());

  var endResponse = ui.prompt('Enter the end date for filtering invoices', 'Please use the format YYYY-MM-DD', ui.ButtonSet.OK_CANCEL);
  if (endResponse.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  var endDate = new Date(endResponse.getResponseText());

  convertSheetToGSTR1JSON(startDate, endDate);
}

function convertSheetToGSTR1JSON(startDate, endDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var gstin = "27AJKPA4618F1ZG";
  var fp = "092024";  
  var gt = 11981565.00;  
  var cur_gt = 7244065.00;  
  
  var jsonData = {
    gstin: gstin,
    fp: fp,
    gt: gt,
    cur_gt: cur_gt,
    b2b: []
  };
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    var clientGSTIN = row[getColumnIndexByName("GSTIN", data[0])];
    var invoiceNo = row[getColumnIndexByName("Invoice No", data[0])];
    var invoiceDate = new Date(row[getColumnIndexByName("Invoice Date", data[0])]);
    var totalAmount = parseFloat(row[getColumnIndexByName("Total Bill Amount (Excl GST)", data[0])]).toFixed(2);
    var gstAmount = parseFloat(row[getColumnIndexByName("GST Amount", data[0])]).toFixed(2);
    var serviceTaxPayableBy = row[getColumnIndexByName("Service Tax Payable by", data[0])];
    var gstr1Applicable = row[getColumnIndexByName("GSTR1 Applicable?", data[0])];
    
    if (invoiceDate >= startDate && invoiceDate <= endDate) {
      if (gstr1Applicable == "Y") {
        var rchrg = (serviceTaxPayableBy == "Carrier") ? "N" : "Y";  
        var gstRate = (gstAmount / totalAmount) * 100;
        var isIntraState = clientGSTIN.substring(0, 2) == "27";  
        
        var igst = isIntraState ? 0 : parseFloat(gstAmount).toFixed(2);
        var cgst = isIntraState ? (parseFloat(gstAmount) / 2).toFixed(2) : 0;
        var sgst = isIntraState ? (parseFloat(gstAmount) / 2).toFixed(2) : 0;
        
        jsonData.b2b.push({
          "ctin": clientGSTIN,
          "inv": [{
            "inum": invoiceNo,
            "idt": formatDateForGST(invoiceDate),  
            "val": parseFloat(totalAmount),  
            "pos": clientGSTIN.substring(0, 2),  
            "rchrg": rchrg,
            "inv_typ": "R",  
            "itms": [{
              "num": 1,  
              "itm_det": {
                "rt": parseFloat(gstRate),  
                "txval": parseFloat(totalAmount),  
                "iamt": parseFloat(igst),  
                "camt": parseFloat(cgst),  
                "samt": parseFloat(sgst)  
              }
            }]
          }]
        });
      }
    }
  }

  // Convert the JSON object to a pretty-printed string
  var jsonString = JSON.stringify(jsonData, null, 2);  // 2 spaces for indentation
  
  // Create a file in Google Drive and save the formatted JSON
  var file = DriveApp.createFile("GSTR1_JSON_" + fp + ".json", jsonString, "application/json");
  
  // Provide the download link for the user
  var ui = SpreadsheetApp.getUi();
  ui.alert('Pretty-printed JSON file created. You can download it from: ' + file.getUrl());
}

function formatDateForGST(date) {
  var d = new Date(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy");
}

function getColumnIndexByName(columnName, headerRow) {
  return headerRow.indexOf(columnName);
}
