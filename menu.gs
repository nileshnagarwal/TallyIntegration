/**
 * Main menu configuration for Tally Integration
 * This file centralizes all menu-related functionality
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const tally_menu = ui.createMenu("Tally Integration");

    // Purchase related menu items
    tally_menu.addSubMenu(ui.createMenu("Purchase")
        .addItem("Verify Transporter Ledgers", "verifyLedgersOnly")
        .addItem("Generate Ledger & Purchase XML", "promptForChallanRange"))
        
        
    // Sales related menu items
    tally_menu.addSubMenu(ui.createMenu("Sales")
        .addItem("Verify Client Ledgers", "verifyClientLedgersOnly")
        .addItem("Generate Ledger & Sales XML ", "promptForBillRange"))
        
    
    // Bank related menu items
    tally_menu.addSubMenu(ui.createMenu("Bank")
        .addItem("Match Bank Entries to Ledgers", "processBankLedgerMatches")
        .addItem("Generate Bank XML", "processBankEntries"))
    tally_menu.addToUi();

    const gst_menu = ui.createMenu('GST Tools');
    
    gst_menu.addItem('Convert to GSTR-1 JSON', 'showDatePrompt');
    gst_menu.addToUi();
}
