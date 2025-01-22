/****************************************************
 * MENU & ENTRY POINTS
 ****************************************************/

/**
 * onOpen() adds a custom menu to your spreadsheet.
 * The user can click these menu items to run the scripts.
 */
function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu("Tally Integration")
      .addItem("Generate Ledger & Purchase XML...", "promptUserForChallanRange")
      .addItem("Verify Ledgers...", "verifyLedgers")
      .addToUi();
  }
  
  /**
   * Prompts the user for a range of rows (e.g., "2-10") from the "Challans" sheet,
   * then generates two XML files:
   *  1) Ledger Creation (for any missing ledgers)
   *  2) Purchase Entries (for all challans in the range)
   */
  function promptUserForChallanRange() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      "Generate Tally XML",
      'Enter row range in "Challans" sheet (e.g. "2-10"):',
      ui.ButtonSet.OK_CANCEL
    );
  
    if (response.getSelectedButton() == ui.Button.OK) {
      var rangeStr = response.getResponseText();
      var parts = rangeStr.split("-");
      if (parts.length === 2) {
        var startRow = parseInt(parts[0], 10);
        var endRow = parseInt(parts[1], 10);
        generateTallyXml(startRow, endRow);
      } else {
        ui.alert('Invalid range. Please enter something like "2-10".');
      }
    }
  }
  
  /****************************************************
   * MAIN FUNCTION TO GENERATE BOTH XML FILES
   ****************************************************/
  function generateTallyXml(startRow, endRow) {
    // 1. Get the "Challans" sheet and read data
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var challanSheet = ss.getSheetByName("Challans");
    if (!challanSheet) {
      SpreadsheetApp.getUi().alert('Sheet named "Challans" not found!');
      return;
    }
  
    // Adjust columns as necessary. For this example, we assume:
    // - Transporter Name is in column E (index = 5)
    // - Challan Date is in column C (index = 3)
    // - Challan No is in column A (index = 1)
    // - Total Challan Amount is in column ? (for example, let's assume column Z = 26)
    // etc.
    // Please map them to your actual column positions.
    var numCols = 55; // or however many columns you want to read
    var dataRange = challanSheet.getRange(startRow, 1, endRow - startRow + 1, numCols);
    var values = dataRange.getValues();
  
    // 2. Collect relevant data from each row
    var challanData = [];
    values.forEach(function(row) {
      var obj = {};
      // Example mapping (adjust indices to match your sheet):
      obj.challanNo = row[0];              // Column A
      obj.invoiceNo = row[1];             // Column B
      // row[2] might be date or something else. Adjust accordingly.
      obj.challanDate = row[2];           // Column C (Challan Date)
      obj.transporterName = row[4];       // Column E (Transporter Name)
      obj.comments = row[27];             // Column AB (Comments) (if that’s your 28th column)
      obj.totalChallanAmt = row[25];      // Column Z? (Total Challan Amount)
      // ... and so on, capturing the fields you need.
  
      challanData.push(obj);
    });
  
    // 3. Determine missing ledgers by comparing the unique transporter names
    //    against the "Ledgers" sheet.
    var missingTransporters = getMissingTransporters(challanData);
  
    // 4. Generate the Ledger Creation XML (if needed)
    var ledgerXml = "";
    if (missingTransporters.length > 0) {
      ledgerXml = generateLedgerCreationXml(missingTransporters);
    }
  
    // 5. Generate the Purchase (Challans) XML for the entire range
    //    (even if some transporters are missing, we still produce it,
    //     but user must import the ledger-xml first in Tally).
    var purchaseXml = generatePurchaseXml(challanData);
  
    // 6. Write both XML strings to Google Drive
    var folder = DriveApp.getRootFolder(); // or getFolderById("YOUR_FOLDER_ID")
    var timeStamp = new Date().toISOString().replace(/[^0-9]/g, "").slice(0,14);
  
    // Ledger Creation File
    if (ledgerXml) {
      var ledgerFileName = "LedgerCreation_" + timeStamp + ".xml";
      var ledgerFile = folder.createFile(ledgerFileName, ledgerXml, MimeType.PLAIN_TEXT);
    }
  
    // Purchase File
    var purchaseFileName = "PurchaseChallans_" + timeStamp + ".xml";
    var purchaseFile = folder.createFile(purchaseFileName, purchaseXml, MimeType.PLAIN_TEXT);
  
    // 7. Notify the user
    var ui = SpreadsheetApp.getUi();
    var msg = "XML Generation Complete!\n";
    if (ledgerXml) {
      msg += "\n1) Ledger Creation XML: " + ledgerFile.getUrl();
      msg += "\n   (Import this first in Tally, then update 'Ledgers' sheet.)";
    } else {
      msg += "\nNo new ledgers needed. All transporters already exist!";
    }
    msg += "\n\n2) Purchase Challans XML: " + purchaseFile.getUrl();
    msg += "\n   (Import this after ledgers are updated.)";
  
    ui.alert("Tally XML Creation", msg, ui.ButtonSet.OK);
  }
  
  /****************************************************
   * HELPER: GET MISSING TRANSPORTERS
   ****************************************************/
  function getMissingTransporters(challanData) {
    // Unique set of transporters from challanData
    var transporters = {};
    challanData.forEach(function(item) {
      if (item.transporterName) {
        transporters[item.transporterName.trim()] = true;
      }
    });
    var uniqueChallanTransporters = Object.keys(transporters);
  
    // Read the "Ledgers" sheet for existing ledger names
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ledgerSheet = ss.getSheetByName("Ledgers");
    if (!ledgerSheet) {
      // If there's no "Ledgers" sheet, all are missing for now
      return uniqueChallanTransporters;
    }
    var ledgerData = ledgerSheet.getDataRange().getValues(); 
    // Adjust which column holds ledger names; assume column A for demonstration
    var existingLedgers = {};
    ledgerData.forEach(function(row, index) {
      if (index === 0) return; // skip header if any
      var ledgerName = row[0]; // column A
      if (ledgerName && ledgerName.trim() !== "") {
        existingLedgers[ledgerName.trim().toLowerCase()] = true;
      }
    });
  
    // Compare
    var missing = [];
    uniqueChallanTransporters.forEach(function(transporter) {
      if (!existingLedgers[transporter.toLowerCase()]) {
        missing.push(transporter);
      }
    });
  
    return missing;
  }
  
  /****************************************************
   * HELPER: GENERATE LEDGER CREATION XML
   *  - Minimal example referencing your ledger_creation.py structure
   ****************************************************/
  function generateLedgerCreationXml(missingTransporters) {
    // Basic Tally Envelope for ledger import
    // This structure references your ledger_creation.py approach
    var xmlParts = [];
    xmlParts.push('<?xml version="1.0" encoding="UTF-8"?>');
    xmlParts.push('<ENVELOPE>');
    xmlParts.push('  <HEADER>');
    xmlParts.push('    <TALLYREQUEST>Import Data</TALLYREQUEST>');
    xmlParts.push('  </HEADER>');
    xmlParts.push('  <BODY>');
    xmlParts.push('    <IMPORTDATA>');
    xmlParts.push('      <REQUESTDESC>');
    xmlParts.push('        <REPORTNAME>All Masters</REPORTNAME>');
    xmlParts.push('        <STATICVARIABLES>');
    xmlParts.push('          <SVCURRENTCOMPANY>Your Company Name</SVCURRENTCOMPANY>');
    xmlParts.push('        </STATICVARIABLES>');
    xmlParts.push('      </REQUESTDESC>');
    xmlParts.push('      <REQUESTDATA>');
  
    // For each missing transporter, create a TALLYMESSAGE block for ledger creation
    missingTransporters.forEach(function(name, idx) {
      xmlParts.push('        <TALLYMESSAGE xmlns:UDF="TallyUDF">');
      // Example: two ledgers, "LH Payable" & "LH Acc" or just one ledger?
      // Adjust as per your needs. Below is a single ledger creation example:
      xmlParts.push('          <LEDGER NAME="' + escXml(name + ' LH Payable') + '" RESERVEDNAME="">');
      xmlParts.push('            <PARENT>Lorry Hire Payable</PARENT>');
      xmlParts.push('            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      xmlParts.push('            <ALTERID>' + (1000 + idx) + '</ALTERID>'); // just a dummy example
      xmlParts.push('            <OPENINGBALANCE>0</OPENINGBALANCE>');
      // LanguageName block
      xmlParts.push('            <LANGUAGENAME.LIST>');
      xmlParts.push('              <NAME.LIST TYPE="String">');
      xmlParts.push('                <NAME>' + escXml(name + ' LH Payable') + '</NAME>');
      xmlParts.push('              </NAME.LIST>');
      xmlParts.push('              <LANGUAGEID>1033</LANGUAGEID>');
      xmlParts.push('            </LANGUAGENAME.LIST>');
      xmlParts.push('          </LEDGER>');
      xmlParts.push('        </TALLYMESSAGE>');
    });
  
    xmlParts.push('      </REQUESTDATA>');
    xmlParts.push('    </IMPORTDATA>');
    xmlParts.push('  </BODY>');
    xmlParts.push('</ENVELOPE>');
  
    return xmlParts.join('\n');
  }
  
  /****************************************************
   * HELPER: GENERATE PURCHASE XML
   *  - Mimicking your purchase.py structure
   ****************************************************/
  function generatePurchaseXml(challanData) {
    var xmlParts = [];
    xmlParts.push('<?xml version="1.0" encoding="UTF-8"?>');
    xmlParts.push('<ENVELOPE>');
    xmlParts.push('  <HEADER>');
    xmlParts.push('    <TALLYREQUEST>Import Data</TALLYREQUEST>');
    xmlParts.push('  </HEADER>');
    xmlParts.push('  <BODY>');
    xmlParts.push('    <IMPORTDATA>');
    xmlParts.push('      <REQUESTDESC>');
    xmlParts.push('        <REPORTNAME>Vouchers</REPORTNAME>');
    xmlParts.push('        <STATICVARIABLES>');
    xmlParts.push('          <SVCURRENTCOMPANY>Nimbus Logistics 2020-21</SVCURRENTCOMPANY>');
    xmlParts.push('        </STATICVARIABLES>');
    xmlParts.push('      </REQUESTDESC>');
    xmlParts.push('      <REQUESTDATA>');
  
    challanData.forEach(function(obj, idx) {
      // Format the date in YYYYMMDD as your python code does
      var dateStr = formatTallyDate(obj.challanDate); 
      var transporter = obj.transporterName || 'Unknown Transporter';
      var challanNo = obj.challanNo || '';
      var totalAmt = obj.totalChallanAmt || 0;
      var comments = obj.comments || '';
  
      xmlParts.push('        <TALLYMESSAGE xmlns:UDF="TallyUDF">');
      xmlParts.push('          <VOUCHER ACTION="Create" VCHTYPE="Purchase">');
      xmlParts.push('            <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>');
      xmlParts.push('            <DATE>' + dateStr + '</DATE>');
      xmlParts.push('            <NARRATION>' + escXml(comments) + '</NARRATION>');
      xmlParts.push('            <VOUCHERNUMBER>' + escXml(challanNo) + '</VOUCHERNUMBER>');
      xmlParts.push('            <GUID></GUID>');    // or leave blank / generate if needed
      xmlParts.push('            <ALTERID></ALTERID>');
  
      // Ledger Entries:
      //   1) Debit the Transporter LH Payable
      //   2) Credit the Transporter LH Acc (assuming same logic as python code)
      // Adjust the logic & ledger naming to match your workflow
      // For example, your python code did "LH Payable" as positive, "LH Acc" as negative
      // (the code below is just an example).
      
      // Leg 1 (LH Payable, positive)
      xmlParts.push('            <ALLLEDGERENTRIES.LIST>');
      xmlParts.push('              <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>');
      xmlParts.push('              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
      xmlParts.push('              <LEDGERNAME>' + escXml(transporter + ' LH Payable') + '</LEDGERNAME>');
      xmlParts.push('              <AMOUNT>' + totalAmt + '</AMOUNT>');
      xmlParts.push('            </ALLLEDGERENTRIES.LIST>');
  
      // Leg 2 (LH Acc, negative)
      xmlParts.push('            <ALLLEDGERENTRIES.LIST>');
      xmlParts.push('              <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>');
      xmlParts.push('              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
      xmlParts.push('              <LEDGERNAME>' + escXml(transporter + ' LH Acc') + '</LEDGERNAME>');
      xmlParts.push('              <AMOUNT>-' + totalAmt + '</AMOUNT>');
      xmlParts.push('            </ALLLEDGERENTRIES.LIST>');
  
      xmlParts.push('          </VOUCHER>');
      xmlParts.push('        </TALLYMESSAGE>');
    });
  
    xmlParts.push('      </REQUESTDATA>');
    xmlParts.push('    </IMPORTDATA>');
    xmlParts.push('  </BODY>');
    xmlParts.push('</ENVELOPE>');
  
    return xmlParts.join('\n');
  }
  
  /****************************************************
   * VERIFY LEDGERS FUNCTION
   *  - Optionally run AFTER user imports ledger XML & updates the "Ledgers" sheet
   *  - It checks a user-specified challan range and ensures all transporters exist
   ****************************************************/
  function verifyLedgers() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      "Verify Ledgers",
      'Enter row range in "Challans" sheet (e.g. "2-10") to verify ledger existence:',
      ui.ButtonSet.OK_CANCEL
    );
  
    if (response.getSelectedButton() == ui.Button.OK) {
      var rangeStr = response.getResponseText();
      var parts = rangeStr.split("-");
      if (parts.length === 2) {
        var startRow = parseInt(parts[0], 10);
        var endRow = parseInt(parts[1], 10);
        var missing = checkMissingLedgersInRange(startRow, endRow);
        if (missing.length === 0) {
          ui.alert("Success", "All transporters in this range are present in Ledgers!", ui.ButtonSet.OK);
        } else {
          ui.alert("Missing Ledgers", 
                   "The following ledgers are still missing:\n" + missing.join(", "), 
                   ui.ButtonSet.OK);
        }
      } else {
        ui.alert('Invalid range. Please enter something like "2-10".');
      }
    }
  }
  
  /**
   * Helper to check missing ledgers in a given row range in "Challans".
   */
  function checkMissingLedgersInRange(startRow, endRow) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var challanSheet = ss.getSheetByName("Challans");
    var ledgerSheet = ss.getSheetByName("Ledgers");
  
    // Get the relevant challan rows
    var numCols = 55; 
    var dataRange = challanSheet.getRange(startRow, 1, endRow - startRow + 1, numCols);
    var values = dataRange.getValues();
  
    // Unique set of transporters in that range
    var transporters = {};
    values.forEach(function(row) {
      var tName = row[4]; // column E example
      if (tName) {
        transporters[tName.trim().toLowerCase()] = true;
      }
    });
    var needed = Object.keys(transporters);
  
    // Read the "Ledgers" sheet
    var ledgerData = ledgerSheet.getDataRange().getValues();
    var existing = {};
    ledgerData.forEach(function(row, index) {
      if (index === 0) return; // skip header
      var ledgerName = row[0]; // column A
      if (ledgerName) {
        existing[ledgerName.trim().toLowerCase()] = true;
      }
    });
  
    // Compare
    var missing = [];
    needed.forEach(function(name) {
      if (!existing[name]) {
        missing.push(name);
      }
    });
    return missing;
  }
  
  /****************************************************
   * UTILITY FUNCTIONS
   ****************************************************/
  
  /**
   * Format date as Tally expects (YYYYMMDD) – similar to your Python code.
   * If the cell is not a valid date, we default to an empty string.
   */
  function formatTallyDate(dateValue) {
    // If dateValue is not a valid Date object, try converting
    var d = new Date(dateValue);
    if (isNaN(d.getTime())) {
      return ""; // or handle invalid date as needed
    }
  
    var yyyy = d.getFullYear().toString();
    var mm = (d.getMonth() + 1).toString().padStart(2, "0");
    var dd = d.getDate().toString().padStart(2, "0");
  
    return yyyy + mm + dd; 
  }
  
  /**
   * Escapes XML special characters (&, <, >, ", ') to avoid broken XML.
   */
  function escXml(str) {
    if (!str || typeof str !== "string") return "";
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }
  