  /**
   * Prompts the user for a range of rows (e.g., "2-10") from the "Challans" sheet,
   * then generates two XML files:
   *  1) Ledger Creation (for any missing ledgers)
   *  2) Purchase Entries (for all challans in the range)
   */
  function promptForChallanRange() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      "Generate Tally XML",
      'Enter range of Challan numbers (e.g. "75-76"):',
      ui.ButtonSet.OK_CANCEL
    );
  
    if (response.getSelectedButton() == ui.Button.OK) {
      try {
        const range = response.getResponseText();
        const [startChallan, endChallan] = range.split("-").map(Number);
        
        if (!startChallan || !endChallan || startChallan > endChallan) {
          ui.alert('Error', 'Invalid range. Please enter something like "75-76".', ui.ButtonSet.OK);
          return;
        }
  
        // Get challan data for the specified range
        const challanData = getChallanDataForRange(range);
        
        // Get unique transporters for the specified range
        const transporters = getTransportersInRange(range);
        
        // Verify ledgers and get missing ones
        const missingLedgers = verifyLedgers(transporters);
        
        // Generate purchase XML
        const purchaseXml = generatePurchaseXml(challanData);
        const purchaseFile = DriveApp.createFile('purchase.xml', purchaseXml, 'application/xml');
        
        // If there are missing ledgers, create ledger XML and show comprehensive message
        if (Object.keys(missingLedgers).length > 0) {
          // Create ledger creation XML
          const ledgerXml = createLedgerXML(missingLedgers);
          const ledgerFile = DriveApp.createFile('ledger_creation.xml', ledgerXml, 'application/xml');
          
          // Create message listing missing ledgers
          let missingLedgersList = '';
          Object.entries(missingLedgers).forEach(([transporter, ledgers]) => {
            missingLedgersList += `\n${transporter}: ${ledgers.join(', ')}`;
          });
          
          // Show comprehensive message to user
          const message = "Two XML files have been created:\n\n" +
            "1. Ledger Creation XML: " + ledgerFile.getUrl() + "\n" +
            "2. Purchase Voucher XML: " + purchaseFile.getUrl() + "\n\n" +
            "Missing ledgers:" + missingLedgersList + "\n\n" +
            "Please follow these steps:\n" +
            "1. First, import the ledger creation XML in Tally\n" +
            "2. Update the 'Ledgers' sheet with the newly created ledgers\n" +
            "3. Run 'Verify Ledgers' function to confirm all ledgers exist\n" +
            "4. Once verified, import the purchase voucher XML";
          
          ui.alert("XML Generation Complete", message, ui.ButtonSet.OK);
        } else {
          // If no missing ledgers, show simple success message
          ui.alert("XML Generation Complete", 
                  "Purchase XML file created successfully!\n\n" + purchaseFile.getUrl(), 
                  ui.ButtonSet.OK);
        }
      } catch (error) {
        // Show error message to user
        ui.alert('Error', 
                'Could not generate XML files:\n\n' + error.message, 
                ui.ButtonSet.OK);
        return;
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
      obj.comments = row[27];             // Column AB (Comments) (if that's your 28th column)
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
  function verifyLedgersOnly() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      "Verify Ledgers",
      'Enter range of Challan numbers (e.g. "75-76"):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const range = response.getResponseText();
      const transporters = getTransportersInRange(range);
      const missingLedgers = verifyLedgers(transporters);
      
      if (Object.keys(missingLedgers).length > 0) {
        let message = "The following ledgers are still missing:\n\n";
        Object.entries(missingLedgers).forEach(([transporter, ledgers]) => {
          message += `${transporter}: ${ledgers.join(', ')}\n`;
        });
        message += "\nPlease create these ledgers in Tally and update the 'Ledgers' sheet before importing the purchase voucher XML.";
        ui.alert("Missing Ledgers", message, ui.ButtonSet.OK);
      } else {
        ui.alert("Success", "All required ledgers exist! You can now proceed with importing the purchase voucher XML.", ui.ButtonSet.OK);
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
   * Format date as Tally expects (YYYYMMDD)
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
  
  function getTransportersInRange(range) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Challans");
    if (!sheet) {
      throw new Error("'Challans' sheet not found in the spreadsheet");
    }
  
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const transporterIndex = headers.indexOf('Transporter Name');
    const challanNoIndex = headers.indexOf('Challan No');
    
    if (transporterIndex === -1 || challanNoIndex === -1) {
      throw new Error("Required columns 'Transporter Name' or 'Challan No' not found");
    }
  
    const [startChallan, endChallan] = range.split('-').map(Number);
    
    // Get unique transporters in the range
    const transporters = new Set();
    data.slice(1).forEach(row => {
      const challanNo = Number(row[challanNoIndex]);
      if (challanNo >= startChallan && challanNo <= endChallan) {
        const transporter = row[transporterIndex];
        if (transporter && transporter !== "Nimbus Logistics") {
          transporters.add(transporter.trim());
        }
      }
    });
    
    return Array.from(transporters);
  }
  
  function verifyLedgers(transporters) {
    // Get the Ledgers sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ledgerSheet = ss.getSheetByName('Ledgers');
    if (!ledgerSheet) {
      throw new Error("'Ledgers' sheet not found in the spreadsheet");
    }
    
    // Get all ledger names from the first column
    const ledgers = ledgerSheet.getRange(1, 1, ledgerSheet.getLastRow(), 1)
      .getValues()
      .map(row => row[0].toString().trim());
    
    // Object to store missing ledgers for each transporter
    const missingLedgers = {};
    
    // Check each transporter's ledgers
    transporters.forEach(transporter => {
      if (transporter === "Nimbus Logistics") return; // Skip Nimbus Logistics
      
      const requiredLedgers = [
        `${transporter} LH Payable`,
        `${transporter} LH Acc`
      ];
      
      const missing = requiredLedgers.filter(ledger => !ledgers.includes(ledger));
      if (missing.length > 0) {
        missingLedgers[transporter] = missing;
      }
    });
    
    return missingLedgers;
  }
  
  function createLedgerXML(missingLedgers) {
    let xml = '<?xml version="1.0" encoding="UTF-8"?>\n';
    xml += '<ENVELOPE>\n';
    xml += '\t<HEADER>\n';
    xml += '\t\t<TALLYREQUEST>Import Data</TALLYREQUEST>\n';
    xml += '\t</HEADER>\n';
    xml += '\t<BODY>\n';
    xml += '\t\t<IMPORTDATA>\n';
    xml += '\t\t\t<REQUESTDESC>\n';
    xml += '\t\t\t\t<REPORTNAME>All Masters</REPORTNAME>\n';
    xml += '\t\t\t\t<STATICVARIABLES>\n';
    xml += '\t\t\t\t\t<SVCURRENTCOMPANY>Nimbus Logistics 2020-21</SVCURRENTCOMPANY>\n';
    xml += '\t\t\t\t</STATICVARIABLES>\n';
    xml += '\t\t\t</REQUESTDESC>\n';
    xml += '\t\t\t<REQUESTDATA>\n';
    
    // Create ledger entries for each missing ledger
    Object.entries(missingLedgers).forEach(([transporter, ledgers]) => {
      ledgers.forEach(ledgerName => {
        xml += '\t\t\t\t<TALLYMESSAGE xmlns:UDF="TallyUDF">\n';
        xml += '\t\t\t\t\t<LEDGER NAME="' + escXml(ledgerName) + '">\n';
        xml += '\t\t\t\t\t\t<PARENT>Lorry Hire Payable</PARENT>\n';
        xml += '\t\t\t\t\t\t<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>\n';
        xml += '\t\t\t\t\t\t<OPENINGBALANCE>0</OPENINGBALANCE>\n';
        xml += '\t\t\t\t\t\t<LANGUAGENAME.LIST>\n';
        xml += '\t\t\t\t\t\t\t<NAME.LIST TYPE="String">\n';
        xml += '\t\t\t\t\t\t\t\t<NAME>' + escXml(ledgerName) + '</NAME>\n';
        xml += '\t\t\t\t\t\t\t</NAME.LIST>\n';
        xml += '\t\t\t\t\t\t\t<LANGUAGEID>1033</LANGUAGEID>\n';
        xml += '\t\t\t\t\t\t</LANGUAGENAME.LIST>\n';
        xml += '\t\t\t\t\t</LEDGER>\n';
        xml += '\t\t\t\t</TALLYMESSAGE>\n';
      });
    });
    
    xml += '\t\t\t</REQUESTDATA>\n';
    xml += '\t\t</IMPORTDATA>\n';
    xml += '\t</BODY>\n';
    xml += '</ENVELOPE>';
    
    return xml;
  }
  
  function getChallanDataForRange(range) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Challans");
    if (!sheet) {
        throw new Error("'Challans' sheet not found!");
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Define required fields (excluding Comments)
    const requiredFields = [
        'Challan No',
        'Invoice Number',
        'Challan Date',
        'Transporter Name',
        'Total Challan Amount (Not Incl TDS Non Deductible Charges)'
    ];
    
    // Verify all required columns exist
    const missingColumns = requiredFields.filter(field => !headers.includes(field));
    if (missingColumns.length > 0) {
        throw new Error(`Required columns not found in sheet: ${missingColumns.join(', ')}`);
    }
    
    // Get column indices for all fields
    const fieldIndices = {
        challanNo: headers.indexOf('Challan No'),
        invoiceNo: headers.indexOf('Invoice Number'),
        challanDate: headers.indexOf('Challan Date'),
        transporterName: headers.indexOf('Transporter Name'),
        totalChallanAmount: headers.indexOf('Total Challan Amount (Not Incl TDS Non Deductible Charges)'),
        comments: headers.indexOf('Comments') // Optional field
    };

    const [startChallan, endChallan] = range.split('-').map(Number);
    
    // Object to store missing field information
    const missingFieldsInfo = [];
    
    // Collect relevant data from each row
    const challanData = [];
    data.slice(1).forEach((row, rowIndex) => {
        const challanNo = Number(row[fieldIndices.challanNo]);
        if (challanNo >= startChallan && challanNo <= endChallan) {
            const transporterName = row[fieldIndices.transporterName];
            
            // Skip Nimbus Logistics
            if (transporterName === "Nimbus Logistics") {
                return;
            }

            // Check for missing or invalid data
            const missingFields = [];
            
            // Check Challan No
            if (!challanNo || isNaN(challanNo)) {
                missingFields.push('Challan No');
            }
            
            // Check Invoice Number
            if (!row[fieldIndices.invoiceNo]) {
                missingFields.push('Invoice Number');
            }
            
            // Check Challan Date
            const challanDate = row[fieldIndices.challanDate];
            if (!challanDate || !(challanDate instanceof Date) || isNaN(challanDate.getTime())) {
                missingFields.push('Challan Date');
            }
            
            // Check Transporter Name
            if (!transporterName || transporterName.trim() === '') {
                missingFields.push('Transporter Name');
            }
            
            // Check Total Challan Amount
            const amount = row[fieldIndices.totalChallanAmount];
            if (!amount || isNaN(Number(amount)) || Number(amount) <= 0) {
                missingFields.push('Total Challan Amount (Not Incl TDS Non Deductible Charges)');
            }

            // If any required fields are missing, add to missing fields info
            if (missingFields.length > 0) {
                missingFieldsInfo.push({
                    challanNo: challanNo || `Row ${rowIndex + 2}`,
                    missingFields: missingFields
                });
                return; // Skip this row
            }

            // If all required fields are present, add to challanData
            challanData.push({
                challanNo: challanNo,
                invoiceNo: row[fieldIndices.invoiceNo],
                challanDate: challanDate,
                transporterName: transporterName,
                totalChallanAmt: Number(row[fieldIndices.totalChallanAmount]),
                comments: fieldIndices.comments !== -1 ? (row[fieldIndices.comments] || '') : ''
            });
        }
    });
    
    // If there are missing fields, throw an error with detailed information
    if (missingFieldsInfo.length > 0) {
        let errorMessage = 'Missing or invalid data found:\n\n';
        missingFieldsInfo.forEach(info => {
            errorMessage += `Challan ${info.challanNo}:\n`;
            errorMessage += `  Missing/Invalid fields: ${info.missingFields.join(', ')}\n`;
        });
        throw new Error(errorMessage);
    }
    
    return challanData;
  }
  