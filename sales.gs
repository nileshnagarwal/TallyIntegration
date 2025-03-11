/*
XML generation for sales entries in Google Sheets to Tally.
*/

function promptForBillRange() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
        "Generate Sales XML",
        'Enter range of Bill numbers (e.g. "75-76"):',
        ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() == ui.Button.OK) {
        try {
            const range = response.getResponseText();
            const [startBill, endBill] = range.split("-").map(Number);
            
            if (!startBill || !endBill || startBill > endBill) {
                ui.alert('Error', 'Invalid range. Please enter something like "75-76".', ui.ButtonSet.OK);
                return;
            }

            // Get bill data for the specified range
            const billData = getBillDataForRange(range);
            
            // Get unique clients for the specified range
            const clients = getClientsInRange(range);
            
            // Verify ledgers and get missing ones
            const missingLedgers = verifyClientLedgers(clients);
            
            // Generate sales XML
            const salesXml = generateSalesXml(billData);
            const salesFile = DriveApp.createFile('sales.xml', salesXml, 'application/xml');
            
            // If there are missing ledgers, create ledger XML and show comprehensive message
            if (Object.keys(missingLedgers).length > 0) {
                const ledgerXml = createClientLedgerXML(missingLedgers);
                const ledgerFile = DriveApp.createFile('client_ledger_creation.xml', ledgerXml, 'application/xml');
                
                // Create message listing missing ledgers
                let missingLedgersList = '';
                Object.entries(missingLedgers).forEach(([client, ledger]) => {
                    missingLedgersList += `\n${client}`;
                });
                
                const message = "Two XML files have been created:\n\n" +
                    "1. Client Ledger Creation XML: " + ledgerFile.getUrl() + "\n" +
                    "2. Sales Voucher XML: " + salesFile.getUrl() + "\n\n" +
                    "Missing client ledgers:" + missingLedgersList + "\n\n" +
                    "Please follow these steps:\n" +
                    "1. First, import the ledger creation XML in Tally\n" +
                    "2. Update the 'Ledgers' sheet with the newly created client ledgers\n" +
                    "3. Run 'Verify Client Ledgers' to confirm all ledgers exist\n" +
                    "4. Once verified, import the sales voucher XML";
                
                ui.alert("XML Generation Complete", message, ui.ButtonSet.OK);
            } else {
                ui.alert("XML Generation Complete", 
                        "Sales XML file created successfully!\n\n" + salesFile.getUrl(), 
                        ui.ButtonSet.OK);
            }
        } catch (error) {
            ui.alert('Error', 'Could not generate XML files:\n\n' + error.message, ui.ButtonSet.OK);
            return;
        }
    }
}

function getBillDataForRange(range) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Bills");
    if (!sheet) {
        throw new Error("'Bills' sheet not found!");
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Define required fields (excluding Comments and GST)
    const requiredFields = [
        'Invoice No',
        'Invoice Date',
        'Client Name',
        'Total Bill Amount (Incl GST)',
        'From/To',
        'Vehicle Number',
        'Challan No'
    ];
    
    // Optional fields that we'll use if present
    const optionalFields = [
        'GST (Payable by Us)',
        'Comments',
        'Note for Detention Charges'
    ];
    
    // Verify all required columns exist
    const missingColumns = requiredFields.filter(field => !headers.includes(field));
    if (missingColumns.length > 0) {
        throw new Error(`Required columns not found in sheet: ${missingColumns.join(', ')}`);
    }
    
    // Get column indices for all fields
    const fieldIndices = {
        invoiceNo: headers.indexOf('Invoice No'),
        invoiceDate: headers.indexOf('Invoice Date'),
        clientName: headers.indexOf('Client Name'),
        totalAmount: headers.indexOf('Total Bill Amount (Incl GST)'),
        totalExclGST: headers.indexOf('Total Bill Amount (Excl GST)'),
        gstAmount: headers.indexOf('GST (Payable by Us)'),
        fromTo: headers.indexOf('From/To'),
        vehicleNumber: headers.indexOf('Vehicle Number'),
        challanNo: headers.indexOf('Challan No'),
        comments: headers.indexOf('Comments'),
        detentionNote: headers.indexOf('Note for Detention Charges')
    };

    const [startBill, endBill] = range.split('-').map(Number);
    const billData = [];
    const missingFieldsInfo = [];

    // Process each row
    data.slice(1).forEach((row, rowIndex) => {
        const billNo = Number(row[fieldIndices.invoiceNo]);
        if (billNo >= startBill && billNo <= endBill) {
            const missingFields = [];
            
            // Check Invoice No
            if (!billNo || isNaN(billNo)) {
                missingFields.push('Invoice No');
            }
            
            // Check Invoice Date
            const invoiceDate = row[fieldIndices.invoiceDate];
            if (!invoiceDate || !(invoiceDate instanceof Date)) {
                missingFields.push('Invoice Date');
            }
            
            // Check Client Name
            const clientName = row[fieldIndices.clientName];
            if (!clientName || clientName.trim() === '') {
                missingFields.push('Client Name');
            }
            
            // Check Total Amount
            const totalAmount = cleanAmount(row[fieldIndices.totalAmount]);
            if (!totalAmount || isNaN(totalAmount) || totalAmount <= 0) {
                missingFields.push('Total Bill Amount (Incl GST)');
            }

            // Check Total Excl GST
            const totalExclGST = cleanAmount(row[fieldIndices.totalExclGST]);
            if (!totalExclGST || isNaN(totalExclGST) || totalExclGST <= 0) {
                missingFields.push('Total Bill Amount (Excl GST)');
            }

            // If any required fields are missing, add to missing fields info
            if (missingFields.length > 0) {
                missingFieldsInfo.push({
                    billNo: billNo || `Row ${rowIndex + 2}`,
                    missingFields: missingFields
                });
                return; // Skip this row
            }

            // If all required fields are present, add to billData
            billData.push({
                invoiceNo: billNo,
                invoiceDate: invoiceDate,
                clientName: clientName.trim(),
                totalAmount: totalAmount,
                totalExclGST: totalExclGST,
                gstAmount: fieldIndices.gstAmount !== -1 ? cleanAmount(row[fieldIndices.gstAmount]) : null,
                fromTo: row[fieldIndices.fromTo] || '',
                vehicleNumber: row[fieldIndices.vehicleNumber] || '',
                challanNo: row[fieldIndices.challanNo] || '',
                comments: fieldIndices.comments !== -1 ? (row[fieldIndices.comments] || '') : '',
                detentionNote: fieldIndices.detentionNote !== -1 ? (row[fieldIndices.detentionNote] || '') : ''
            });
        }
    });
    
    // If there are missing fields, throw an error with detailed information
    if (missingFieldsInfo.length > 0) {
        let errorMessage = 'Missing or invalid data found:\n\n';
        missingFieldsInfo.forEach(info => {
            errorMessage += `Bill ${info.billNo}:\n`;
            errorMessage += `  Missing/Invalid fields: ${info.missingFields.join(', ')}\n`;
        });
        throw new Error(errorMessage);
    }
    
    return billData;
}

function verifyClientLedgers(clients) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ledgerSheet = ss.getSheetByName('Ledgers');
    if (!ledgerSheet) {
        throw new Error("'Ledgers' sheet not found in the spreadsheet");
    }
    
    // Get all ledger names from the first column
    const ledgers = ledgerSheet.getRange(1, 1, ledgerSheet.getLastRow(), 1)
        .getValues()
        .map(row => row[0].toString().trim());
    
    // Object to store missing ledgers
    const missingLedgers = {};
    
    // Check each client's ledger
    clients.forEach(client => {
        if (!ledgers.includes(client)) {
            missingLedgers[client] = client; // Client name is the ledger name
        }
    });
    
    return missingLedgers;
}

function createClientLedgerXML(missingLedgers) {
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
    
    // Create ledger entries for each missing client
    Object.keys(missingLedgers).forEach((clientName, idx) => {
        xml += '\t\t\t\t<TALLYMESSAGE xmlns:UDF="TallyUDF">\n';
        xml += '\t\t\t\t\t<LEDGER NAME="' + escXml(clientName) + '">\n';
        xml += '\t\t\t\t\t\t<PARENT>Clients</PARENT>\n';
        xml += '\t\t\t\t\t\t<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>\n';
        xml += '\t\t\t\t\t\t<OPENINGBALANCE>0</OPENINGBALANCE>\n';
        xml += '\t\t\t\t\t\t<LANGUAGENAME.LIST>\n';
        xml += '\t\t\t\t\t\t\t<NAME.LIST TYPE="String">\n';
        xml += '\t\t\t\t\t\t\t\t<NAME>' + escXml(clientName) + '</NAME>\n';
        xml += '\t\t\t\t\t\t\t</NAME.LIST>\n';
        xml += '\t\t\t\t\t\t\t<LANGUAGEID>1033</LANGUAGEID>\n';
        xml += '\t\t\t\t\t\t</LANGUAGENAME.LIST>\n';
        xml += '\t\t\t\t\t</LEDGER>\n';
        xml += '\t\t\t\t</TALLYMESSAGE>\n';
    });
    
    xml += '\t\t\t</REQUESTDATA>\n';
    xml += '\t\t</IMPORTDATA>\n';
    xml += '\t</BODY>\n';
    xml += '</ENVELOPE>';
    
    return xml;
}

function formatInvoiceNumber(invoiceNo) {
    // Convert to string and pad with leading zeros to ensure 3 digits
    const paddedNumber = invoiceNo.toString().padStart(3, '0');
    return `24-25/${paddedNumber}`;
}

function generateSalesXml(billData) {
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

    billData.forEach(function(bill) {
        const dateStr = formatTallyDate(bill.invoiceDate);
        const clientName = bill.clientName;
        const invoiceNo = bill.invoiceNo;
        const totalAmount = bill.totalAmount;
        
        // Get all amounts first
        const totalInclGST = bill.totalAmount;
        const totalExclGST = bill.totalExclGST;
        const gstPayableByUs = bill.gstAmount;

        // FCM Case (We pay GST)
        if (gstPayableByUs > 0) {
            freightForwardingAmount = totalExclGST;
            clientAmount = -totalInclGST;  // Client pays (Freight + GST)
            gstAmount = gstPayableByUs;
        } 
        // RCM Case (Client pays GST directly)
        else {
            freightForwardingAmount = totalExclGST;
            clientAmount = -freightForwardingAmount;  // Client pays only freight
            gstAmount = 0;
        }

        const narration = `${bill.fromTo}. ${bill.vehicleNumber}. FM${bill.challanNo}` + 
                         (bill.comments ? ` Comments: ${bill.comments}.` : '') +
                         (bill.detentionNote ? ` Detention: ${bill.detentionNote}` : '');

        const formattedInvoiceNo = formatInvoiceNumber(invoiceNo);

        // Generate XML
        xmlParts.push('        <TALLYMESSAGE xmlns:UDF="TallyUDF">');
        xmlParts.push('          <VOUCHER ACTION="Create" VCHTYPE="Sales">');
        xmlParts.push('            <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>');
        xmlParts.push('            <DATE>' + dateStr + '</DATE>');
        xmlParts.push('            <REFERENCE></REFERENCE>');
        xmlParts.push('            <NARRATION>' + escXml(narration) + '</NARRATION>');
        xmlParts.push('            <PARTYNAME>' + escXml(clientName) + '</PARTYNAME>');
        xmlParts.push('            <VOUCHERNUMBER>' + escXml(formattedInvoiceNo) + '</VOUCHERNUMBER>');
        xmlParts.push('            <GUID></GUID>');
        xmlParts.push('            <ALTERID></ALTERID>');

        // Client Ledger Entry (Debit)
        xmlParts.push('            <ALLLEDGERENTRIES.LIST>');
        xmlParts.push('              <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>');
        xmlParts.push('              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>');
        xmlParts.push('              <LEDGERNAME>' + escXml(clientName) + '</LEDGERNAME>');
        xmlParts.push('              <ISPARTLEDGER>Yes</ISPARTLEDGER>');
        xmlParts.push('              <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>');
        xmlParts.push('              <AMOUNT>' + clientAmount.toFixed(2) + '</AMOUNT>');
        xmlParts.push('              <BILLALLOCATIONS.LIST>');
        xmlParts.push('                <NAME>' + escXml(formattedInvoiceNo) + '</NAME>');
        xmlParts.push('                <BILLTYPE>New Ref</BILLTYPE>');
        xmlParts.push('                <AMOUNT>' + clientAmount.toFixed(2) + '</AMOUNT>');
        xmlParts.push('              </BILLALLOCATIONS.LIST>');
        xmlParts.push('            </ALLLEDGERENTRIES.LIST>');

        // Freight Forwarding Entry (Credit)
        xmlParts.push('            <ALLLEDGERENTRIES.LIST>');
        xmlParts.push('              <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>');
        xmlParts.push('              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
        xmlParts.push('              <LEDGERNAME>Freight Forwarding</LEDGERNAME>');
        xmlParts.push('              <ISPARTLEDGER>No</ISPARTLEDGER>');
        xmlParts.push('              <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>');
        xmlParts.push('              <AMOUNT>' + freightForwardingAmount.toFixed(2) + '</AMOUNT>');
        xmlParts.push('            </ALLLEDGERENTRIES.LIST>');

        // GST Entry (Credit) - Only if GST is payable by us
        if (gstAmount) {
            xmlParts.push('            <ALLLEDGERENTRIES.LIST>');
            xmlParts.push('              <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>');
            xmlParts.push('              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>');
            xmlParts.push('              <LEDGERNAME>GST Payable</LEDGERNAME>');
            xmlParts.push('              <ISPARTLEDGER>No</ISPARTLEDGER>');
            xmlParts.push('              <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>');
            xmlParts.push('              <AMOUNT>' + gstAmount.toFixed(2) + '</AMOUNT>');
            xmlParts.push('            </ALLLEDGERENTRIES.LIST>');
        }

        xmlParts.push('          </VOUCHER>');
        xmlParts.push('        </TALLYMESSAGE>');
    });

    xmlParts.push('      </REQUESTDATA>');
    xmlParts.push('    </IMPORTDATA>');
    xmlParts.push('  </BODY>');
    xmlParts.push('</ENVELOPE>');

    return xmlParts.join('\n');
}

function getClientsInRange(range) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Bills");
    if (!sheet) {
        throw new Error("'Bills' sheet not found in the spreadsheet");
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const clientNameIndex = headers.indexOf('Client Name');
    const invoiceNoIndex = headers.indexOf('Invoice No');
    
    if (clientNameIndex === -1 || invoiceNoIndex === -1) {
        throw new Error("Required columns 'Client Name' or 'Invoice No' not found");
    }

    const [startBill, endBill] = range.split('-').map(Number);
    
    // Get unique clients in the range
    const clients = new Set();
    data.slice(1).forEach(row => {
        const billNo = Number(row[invoiceNoIndex]);
        if (billNo >= startBill && billNo <= endBill) {
            const clientName = row[clientNameIndex];
            if (clientName && clientName.trim() !== "") {
                clients.add(clientName.trim());
            }
        }
    });
    
    return Array.from(clients);
}

function cleanAmount(amount) {
    if (typeof amount === 'number') return amount;
    if (!amount) return 0;
    
    // Remove currency symbols, commas and '/-'
    return Number(amount.toString().replace(/[₹,/-]/g, '').trim());
}

function verifyClientLedgersOnly() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
        'Verify Client Ledgers',
        'Enter range of Bill numbers (e.g. "75-76"):',
        ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
        const range = response.getResponseText();
        const clients = getClientsInRange(range);
        const missingLedgers = verifyClientLedgers(clients);
        
        if (Object.keys(missingLedgers).length > 0) {
            let message = "The following client ledgers are missing:\n\n";
            Object.keys(missingLedgers).forEach(client => {
                message += `${client}\n`;
            });
            message += "\nPlease create these ledgers in Tally and update the 'Ledgers' sheet before importing the sales voucher XML.";
            ui.alert("Missing Ledgers", message, ui.ButtonSet.OK);
        } else {
            ui.alert("Success", "All required client ledgers exist! You can now proceed with importing the sales voucher XML.", ui.ButtonSet.OK);
        }
    }
}

// Reuse the helper functions from Code.gs
function formatTallyDate(dateValue) {
    var d = new Date(dateValue);
    if (isNaN(d.getTime())) {
        return "";
    }
    return d.getFullYear().toString() +
           (d.getMonth() + 1).toString().padStart(2, "0") +
           d.getDate().toString().padStart(2, "0");
}

function escXml(str) {
    if (!str || typeof str !== "string") return "";
    return str
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

