/**
 * Matches bank entries to ledgers based on narration patterns
 * and FM numbers.
 */
 
function processBankEntries() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bankSheet = ss.getSheetByName("Bank");
    const challansSheet = ss.getSheetByName("Challans");
    const ledgersSheet = ss.getSheetByName("Ledgers");
    
    if (!bankSheet || !challansSheet || !ledgersSheet) {
        throw new Error("Required sheets not found!");
    }

    // Get data from sheets
    const bankData = bankSheet.getDataRange().getValues();
    const bankHeaders = bankData[0];
    
    // Get column indices
    const dateIndex = bankHeaders.indexOf("Date");
    const narrationIndex = bankHeaders.indexOf("Narration");
    const debitIndex = bankHeaders.indexOf("Debit");
    const creditIndex = bankHeaders.indexOf("Credit");
    const correctLedgerIndex = bankHeaders.indexOf("Correct Ledger Names");
    
    if (correctLedgerIndex === -1) {
        throw new Error("Required column 'Correct Ledger Names' not found!");
    }

    // Process entries and prepare for export
    const processedEntries = bankData.slice(1).map(row => {
        const date = row[dateIndex];
        const narration = row[narrationIndex];
        const debit = parseFloat(String(row[debitIndex]).replace(/,/g, '')) || 0;
        const credit = parseFloat(String(row[creditIndex]).replace(/,/g, '')) || 0;
        const correctLedger = row[correctLedgerIndex];
        
        // Skip if no transaction amount or no matched ledger
        if ((!debit && !credit) || !correctLedger) return null;
        
        // Determine if it's a payment or receipt
        const isPayment = debit > 0;
        const amount = isPayment ? debit : credit;
        const amountWithSign = isPayment ? amount : -amount;
        
        // Extract FM numbers for bill allocation
        const fmNumbers = extractFMNumbers(narration);
        
        return {
            date: formatDate(date),
            voucherType: isPayment ? "Payment" : "Receipt",
            narration: narration,
            ledger1: "IDFC Bank",
            ledger2: correctLedger,
            amountWithSign: amountWithSign,
            billAllocations: fmNumbers.map(fm => ({
                reference: fm,
                amount: amount / fmNumbers.length
            })),
            allocationInLedger: 2
        };
    }).filter(entry => entry !== null);
    
    // Generate XML
    const xmlContent = generateTallyXML(processedEntries);
    
    // Save XML content to file in Google Drive
    saveXMLToFile(xmlContent);
}

function generateTallyXML(entries) {
    let xml = '<?xml version="1.0" encoding="UTF-8"?>\n';
    xml += '<ENVELOPE>\n';
    xml += '  <HEADER>\n';
    xml += '    <TALLYREQUEST>Import Data</TALLYREQUEST>\n';
    xml += '  </HEADER>\n';
    xml += '  <BODY>\n';
    xml += '    <IMPORTDATA>\n';
    xml += '      <REQUESTDESC>\n';
    xml += '        <REPORTNAME>Vouchers</REPORTNAME>\n';
    xml += '        <STATICVARIABLES>\n';
    xml += '          <SVCURRENTCOMPANY>Nimbus Logistics 2020-21</SVCURRENTCOMPANY>\n';
    xml += '        </STATICVARIABLES>\n';
    xml += '      </REQUESTDESC>\n';
    xml += '      <REQUESTDATA>\n';

    // Generate vouchers for each entry
    entries.forEach(entry => {
        xml += generateVoucherXML(entry);
    });

    xml += '      </REQUESTDATA>\n';
    xml += '    </IMPORTDATA>\n';
    xml += '  </BODY>\n';
    xml += '</ENVELOPE>';

    return xml;
}

function generateVoucherXML(entry) {
    let xml = '        <TALLYMESSAGE xmlns:UDF="TallyUDF">\n';
    xml += `          <VOUCHER ACTION="Create" VCHTYPE="${entry.voucherType}">\n`;
    xml += `            <VOUCHERTYPENAME>${entry.voucherType}</VOUCHERTYPENAME>\n`;
    xml += `            <DATE>${entry.date}</DATE>\n`;
    xml += '            <REFERENCE></REFERENCE>\n';
    xml += `            <NARRATION>${escapeXml(entry.narration)}</NARRATION>\n`;
    xml += '            <GUID></GUID>\n';
    xml += '            <ALTERID></ALTERID>\n';

    // First ledger entry (n=1)
    xml += '            <ALLLEDGERENTRIES.LIST>\n';
    xml += '              <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>\n';
    xml += '              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>\n';
    
    if (entry.amountWithSign < 0) {
        // Receipt: Bank is debited
        xml += `              <LEDGERNAME>${escapeXml(entry.ledger1)}</LEDGERNAME>\n`;
        xml += `              <AMOUNT>${entry.amountWithSign}</AMOUNT>\n`;
    } else {
        // Payment: Contra account is credited
        xml += `              <LEDGERNAME>${escapeXml(entry.ledger2)}</LEDGERNAME>\n`;
        xml += `              <AMOUNT>-${entry.amountWithSign}</AMOUNT>\n`;
        
        // Add bill allocations for payments
        if (entry.billAllocations.length > 0) {
            entry.billAllocations.forEach(bill => {
                xml += '              <BILLALLOCATIONS.LIST>\n';
                xml += `                <NAME>${bill.reference}</NAME>\n`;
                xml += '                <BILLTYPE>Advance</BILLTYPE>\n';
                xml += `                <AMOUNT>-${bill.amount}</AMOUNT>\n`;
                xml += '              </BILLALLOCATIONS.LIST>\n';
            });
        }
    }
    xml += '            </ALLLEDGERENTRIES.LIST>\n';

    // Second ledger entry (n=2)
    xml += '            <ALLLEDGERENTRIES.LIST>\n';
    xml += '              <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>\n';
    xml += '              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>\n';
    
    if (entry.amountWithSign < 0) {
        // Receipt: Contra account is credited
        xml += `              <LEDGERNAME>${escapeXml(entry.ledger2)}</LEDGERNAME>\n`;
        xml += `              <AMOUNT>${Math.abs(entry.amountWithSign)}</AMOUNT>\n`;
    } else {
        // Payment: Bank is debited
        xml += `              <LEDGERNAME>${escapeXml(entry.ledger1)}</LEDGERNAME>\n`;
        xml += `              <AMOUNT>${entry.amountWithSign}</AMOUNT>\n`;
    }
    xml += '            </ALLLEDGERENTRIES.LIST>\n';

    xml += '          </VOUCHER>\n';
    xml += '        </TALLYMESSAGE>\n';
    return xml;
}

function generateLedgerEntryXML({ ledgerName, amount, isDeemedPositive, isFirst, entry }) {
    let xml = '            <ALLLEDGERENTRIES.LIST>\n';
    xml += '              <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>\n';
    xml += `              <ISDEEMEDPOSITIVE>${isDeemedPositive}</ISDEEMEDPOSITIVE>\n`;
    xml += `              <LEDGERNAME>${escapeXml(ledgerName)}</LEDGERNAME>\n`;
    
    // Handle amount based on entry type and position
    const displayAmount = isFirst ? 
        (amount < 0 ? amount : `-${amount}`) : 
        (amount < 0 ? Math.abs(amount) : amount);
    
    xml += `              <AMOUNT>${displayAmount}</AMOUNT>\n`;

    // Add bill allocations if this is the ledger that needs them
    if (entry.allocationInLedger === (isFirst ? 1 : 2) && entry.billAllocations.length > 0) {
        entry.billAllocations.forEach(bill => {
            xml += generateBillAllocationXML(bill, amount < 0);
        });
    }

    xml += '            </ALLLEDGERENTRIES.LIST>\n';
    return xml;
}

function generateBillAllocationXML(bill, isReceipt) {
    let xml = '              <BILLALLOCATIONS.LIST>\n';
    xml += `                <NAME>${bill.reference}</NAME>\n`;
    xml += '                <BILLTYPE>Advance</BILLTYPE>\n';
    const billAmount = isReceipt ? bill.amount : `-${bill.amount}`;
    xml += `                <AMOUNT>${billAmount}</AMOUNT>\n`;
    xml += '              </BILLALLOCATIONS.LIST>\n';
    return xml;
}

function escapeXml(unsafe) {
    return unsafe
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function saveXMLToFile(xmlContent) {
    try {
        const fileName = `bank_export_${new Date().toISOString().split('T')[0]}.xml`;
        
        // Create blob with XML content
        const blob = Utilities.newBlob(xmlContent, 'text/xml', fileName);
        
        // Get or create folder
        let folder;
        try {
            folder = DriveApp.getFoldersByName('Bank Exports').next();
        } catch (e) {
            folder = DriveApp.createFolder('Bank Exports');
        }
        
        // Create file
        const file = folder.createFile(blob);
        Logger.log(`XML file created: ${file.getUrl()}`);
        
        // Show success message to user
        SpreadsheetApp.getActiveSpreadsheet().toast(
            `XML file created successfully: ${fileName}`,
            'Export Complete'
        );
        
        return file.getUrl();
    } catch (error) {
        Logger.log(`Error saving XML file: ${error.toString()}`);
        SpreadsheetApp.getActiveSpreadsheet().toast(
            `Error creating XML file: ${error.toString()}`,
            'Export Error',
            10
        );
        throw error;
    }
}

// Helper functions from previous implementation remain the same
function extractFMNumbers(narration) {
    const fmPattern = /FM[- ]?(\d+)/gi;
    const matches = [];
    let match;
    
    while ((match = fmPattern.exec(narration)) !== null) {
        matches.push(match[1]);
    }
    
    return matches;
}

function formatDate(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyyMMdd");
}

function matchFMNumber(narration, challansSheet) {
    // Look for FM followed by numbers, allowing for multiple FM numbers
    const fmPattern = /FM[- ]?(\d+)/gi;
    let matches = [];
    let match;
    
    while ((match = fmPattern.exec(narration)) !== null) {
        const number = match[1];
        // Handle cases where multiple FM numbers are concatenated
        if (number.length > 4) {
            const chunks = number.match(/.{1,4}/g) || [];
            matches.push(...chunks);
        } else {
            matches.push(number);
        }
    }
    
    if (matches.length > 0) {
        const challanData = challansSheet.getDataRange().getValues();
        const headers = challanData[0];
        
        const challanNoIndex = headers.indexOf("Challan No");
        const transporterIndex = headers.indexOf("Transporter Name");
        
        if (challanNoIndex === -1 || transporterIndex === -1) {
            return { success: false };
        }
        
        // Try each FM number until we find a match
        for (const challanNo of matches) {
            const matchingRow = challanData.find(row => 
                row[challanNoIndex].toString() === challanNo.toString()
            );
            
            if (matchingRow) {
                return {
                    success: true,
                    transporterName: matchingRow[transporterIndex],
                    challanNo: challanNo
                };
            }
        }
    }
    
    return { success: false };
}

function findBestLedgerMatch(narration, ledgers) {
    // Define common business suffixes to be given lower weight
    const commonSuffixes = [
        'ROADWAYS', 'ROADLINES', 'TRANSPORT', 'CARGO', 'MOVERS', 
        'LOGISTICS', 'LIMITED', 'PVT', 'PRIVATE', 'LTD'
    ];
    
    // Define preferred suffixes in order of priority
    const preferredSuffixes = ['LH PAYABLE', 'SALARY PAYABLE'];
    
    let bestMatch = {
        ledgerName: null,
        confidence: 0
    };

    // Clean narration
    const cleanNarration = narration.toUpperCase()
        .replace(/[^A-Z0-9\s]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    
    // Group ledgers by base name (without suffixes)
    const ledgerGroups = {};
    ledgers.forEach(ledger => {
        const baseName = ledger.toString().toUpperCase();
        let strippedName = baseName;
        
        // Remove all known suffixes to get the base name
        preferredSuffixes.forEach(suffix => {
            strippedName = strippedName.replace(new RegExp(`\\s*${suffix}$`), '');
        });
        
        if (!ledgerGroups[strippedName]) {
            ledgerGroups[strippedName] = [];
        }
        ledgerGroups[strippedName].push(ledger);
    });

    for (const [baseName, groupLedgers] of Object.entries(ledgerGroups)) {
        const cleanBaseName = baseName
            .replace(/[^A-Z0-9\s]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
            
        // Skip empty base names
        if (!cleanBaseName) continue;

        const baseNameWords = cleanBaseName.split(' ').filter(word => word.length > 2);
        const uniqueWords = baseNameWords.filter(word => !commonSuffixes.includes(word));
        
        // Calculate match score using only the base name
        const matchScore = calculateMatchScore(cleanNarration, uniqueWords);
        
        if (matchScore > bestMatch.confidence) {
            // If we have a match, choose the version with preferred suffix
            const preferredLedger = chooseBestLedgerVersion(groupLedgers, preferredSuffixes);
            bestMatch = {
                ledgerName: preferredLedger,
                confidence: matchScore
            };
        }
    }
    
    return bestMatch;
}

function chooseBestLedgerVersion(ledgers, preferredSuffixes) {
    // Try each suffix in order of preference
    for (const suffix of preferredSuffixes) {
        const found = ledgers.find(ledger => 
            ledger.toString().toUpperCase().endsWith(suffix)
        );
        if (found) return found;
    }
    
    // If no preferred suffix found, return the first ledger
    return ledgers[0];
}

function calculateMatchScore(narration, uniqueWords) {
    let score = 0;
    let matches = 0;
    let totalWeight = 0;
    
    for (const word of uniqueWords) {
        if (word.length <= 3) continue;
        
        const weight = Math.min(1, word.length / 5);
        totalWeight += weight;
        
        if (narration.includes(word)) {
            matches += weight;
        }
    }
    
    return totalWeight > 0 ? matches / totalWeight : 0;
}

function createCommonPatternsSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let patternsSheet = ss.getSheetByName("Common Patterns");
    
    if (!patternsSheet) {
        patternsSheet = ss.insertSheet("Common Patterns");
        
        // Set headers
        const headers = [
            ["Pattern", "Ledger Name", "Type", "Priority", "Notes", "Example Narration"],
            ["TDS|ETDS|TAX DEDUCTED", "TDS Payable", "Payment", "1", "TDS related entries", "TDS PAID FOR Q1"],
            ["SALARY|WAGES", "Salary", "Payment", "1", "Salary payments", "SALARY PAYMENT JUNE"],
            ["GST PMT|GSTIN", "GST Payable", "Payment", "1", "GST payments", "GST PAYMENT FOR MAY"]
        ];
        
        patternsSheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
        patternsSheet.getRange("1:1").setFontWeight("bold");
    }
}

function processBankLedgerMatches() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bankSheet = ss.getSheetByName("Bank");
    const challansSheet = ss.getSheetByName("Challans");
    const ledgersSheet = ss.getSheetByName("Ledgers");
    
    if (!bankSheet || !challansSheet || !ledgersSheet) {
        throw new Error("Required sheets not found!");
    }

    // Get bank data
    const bankData = bankSheet.getDataRange().getValues();
    const bankHeaders = bankData[0];

    // Get column indices
    const narrationIndex = bankHeaders.indexOf("Narration");
    const matchingLedgerIndex = bankHeaders.indexOf("Matching Ledger");
    const confidenceIndex = bankHeaders.indexOf("Confidence");

    // Add columns if they don't exist
    let lastColumn = bankHeaders.length;
    if (matchingLedgerIndex === -1) {
        bankSheet.getRange(1, lastColumn + 1).setValue("Matching Ledger");
        lastColumn++;
    }
    if (confidenceIndex === -1) {
        bankSheet.getRange(1, lastColumn + 1).setValue("Confidence");
        lastColumn++;
    }

    // Get ledgers data
    const ledgers = ledgersSheet.getDataRange()
        .getValues()
        .slice(1) // Remove header
        .map(row => row[0])
        .filter(ledger => ledger); // Remove empty rows

    // Process each row and prepare results
    const results = bankData.slice(1).map(row => {
        const narration = String(row[narrationIndex] || "");
        
        // First try FM number matching
        const fmMatch = matchFMNumber(narration, challansSheet);
        if (fmMatch.success) {
            const transporterLedger = ledgers.find(ledger => 
                ledger.includes(fmMatch.transporterName) && ledger.endsWith("LH Payable")
            );
            
            if (transporterLedger) {
                return [transporterLedger, "High (FM Match)"];
            }
        }

        // Try fuzzy matching with ledgers
        const ledgerMatch = findBestLedgerMatch(narration, ledgers);
        if (ledgerMatch.confidence > 0.8) {
            return [ledgerMatch.ledgerName, `High (${Math.round(ledgerMatch.confidence * 100)}% Match)`];
        }

        return ["", "Low (No Match)"];
    });

    // Update the bank sheet with results
    if (results.length > 0) {
        const matchingLedgerCol = matchingLedgerIndex === -1 ? bankHeaders.length + 1 : matchingLedgerIndex + 1;
        const confidenceCol = confidenceIndex === -1 ? bankHeaders.length + 2 : confidenceIndex + 1;
        
        bankSheet.getRange(2, matchingLedgerCol, results.length, 2).setValues(results);
    }
} 