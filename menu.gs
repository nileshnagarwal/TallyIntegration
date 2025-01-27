/**
 * Main menu configuration for Tally Integration
 * This file centralizes all menu-related functionality
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const tally_menu = ui.createMenu("Tally Integration");
    
    // Sales related menu items
    tally_menu.addSubMenu(ui.createMenu("Sales")
        .addItem("Generate Sales XML...", "promptForBillRange")
        .addItem("Verify Client Ledgers...", "verifyClientLedgersOnly"))
    
    // Purchase related menu items
    tally_menu.addSubMenu(ui.createMenu("Purchase")
        .addItem("Generate Ledger & Purchase XML...", "promptForChallanRange")
        .addItem("Verify Transporter Ledgers...", "verifyLedgersOnly"))
    
    tally_menu.addToUi();

    const gst_menu = ui.createMenu('GST Tools');
    
    gst_menu.addItem('Convert to GSTR-1 JSON', 'showDatePrompt');
    gst_menu.addToUi();
} 