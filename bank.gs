/**
 * Matches bank entries to ledgers based on narration patterns
 * and FM numbers.
 */
 
function processBankEntries() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get required sheets
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
    if (matchingLedgerIndex === -1) {
        bankSheet.getRange(1, bankHeaders.length + 1).setValue("Matching Ledger");
        bankHeaders.push("Matching Ledger");
    }
    if (confidenceIndex === -1) {
        bankSheet.getRange(1, bankHeaders.length + 1).setValue("Confidence");
        bankHeaders.push("Confidence");
    }

    // Get ledgers list and ensure they're strings
    const ledgers = ledgersSheet.getRange("A:A").getValues()
        .flat()
        .filter(ledger => ledger !== "" && ledger !== "Ledger Name")
        .map(ledger => String(ledger).trim());

    // Process each bank entry
    const results = bankData.slice(1).map(row => {
        const narration = String(row[narrationIndex] || "");
        
        if (!narration) {
            return ["", ""];
        }

        // 1. Try FM number matching
        const fmMatch = matchFMNumber(narration, challansSheet);
        if (fmMatch.success) {
            // Find matching ledger with "LH Payable" suffix
            const transporterLedger = ledgers.find(ledger => 
                ledger.includes(fmMatch.transporterName) && ledger.endsWith("LH Payable")
            );
            
            if (transporterLedger) {
                return [transporterLedger, "High (FM Match)"];
            }
            return ["", "Low (No Match)"]; // No matching ledger found
        }

        // 2. Try fuzzy matching with ledgers
        const ledgerMatch = findBestLedgerMatch(narration, ledgers);
        if (ledgerMatch.confidence > 0.8) {
            return [ledgerMatch.ledgerName, `High (${Math.round(ledgerMatch.confidence * 100)}% Match)`];
        }

        return ["", "Low (No Match)"];
    });

    // Update the bank sheet with results
    if (results.length > 0) {
        const matchingLedgerCol = matchingLedgerIndex === -1 ? bankHeaders.length - 1 : matchingLedgerIndex;
        const confidenceCol = confidenceIndex === -1 ? bankHeaders.length : confidenceIndex;
        
        bankSheet.getRange(2, matchingLedgerCol + 1, results.length, 2).setValues(results);
    }
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