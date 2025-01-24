/**
 * Main menu configuration for Tally Integration
 * This file centralizes all menu-related functionality
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("Tally Integration");
    
    // Sales related menu items
    menu.addSubMenu(ui.createMenu("Sales")
        .addItem("Generate Sales XML...", "promptForBillRange")
        .addItem("Verify Client Ledgers...", "verifyClientLedgersOnly"))
    
    // Purchase related menu items
    menu.addSubMenu(ui.createMenu("Purchase")
        .addItem("Generate Ledger & Purchase XML...", "promptForChallanRange")
        .addItem("Verify Transporter Ledgers...", "verifyLedgersOnly"))

    // Bank related menu items
    menu.addSubMenu(ui.createMenu("Bank")
        .addItem("Match Bank Entries to Ledgers...", "processBankEntries"))
    
    menu.addToUi();
} 