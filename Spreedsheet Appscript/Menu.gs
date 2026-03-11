function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Data Organization")
    .addItem("Task 1: Update Raw Data", "RD_updateRawDataHeadersAndPT")
    .addItem("Task 2: Build Organized Tab", "runFullPipeline")
    .addToUi();

  ui.createMenu("Skip Tracing Steps")
    .addItem("Task 1: Ready For XLeads", "readyForXLeads")
    .addItem("Task 2: XLeads Results", "ST_part1_readyVsXleads_noMatchOrNoPhone_pullForImport")
    .addItem("Task 3: Ready For Mojo", "readyForMojo")
    .addItem("Task 4: Set Mojo Results Column","setMojoResultsColumns_InsertMissingInOrder")
    .addItem("Task 5: Mojo Results", "ST_part2_notInXleadsVsMojo_noMatchOrNoPhone_pullForImport")
    .addItem("Task 6: Ready For Direct Skip", "ST_part3_notInMojoVsDirect_noMatchOrNoPhone_pullForImport")
    .addItem("Task 7: Set Direct Skip Results Columns","setDirectSkipResultsColumns_DeleteExtras")
    .addItem("Task 8: Direct Skip Results", "ST_part3_notInMojoVsDirect_noMatchOrNoPhone_pullForImport")
    .addToUi();

  ui.createMenu("Skip Trace Done (Merge)")
    .addItem("Task 1: Append XLeads", "ST_merge_A_xleads_buildHeaderAndWrite")
    .addItem("Task 2: Append Mojo", "ST_merge_B_appendMojo")
    .addItem("Task 3: Append Direct", "ST_merge_C_appendDirect")
    .addItem("Task 4: Inital Merge", "runFullSkipTracePipeline")
    .addItem("Task 5: Final Merge","ST_appendForImportRepData")
    .addItem("Task 6: Set Columns", "ST_setColumns_DeleteExtras")
    .addItem("Task 7: Update Prospect Type & Fields","addOutOfStateTag_HighEquity_LongTermOwner_EstateZ")
    .addItem("Task 8: Merge & Update Missing Details","ST_resolveDuplicates_M_vs_F_and_RemoveAbove550kMatches")
    .addItem("Task 9: Need Manual DM","ST_buildNeedManualDM")
    .addItem("Task 10: Check Multiple Properties","ST_buildMultipleProperties")
    .addToUi();

  ui.createMenu("Duplication")
    .addItem('Task 1: Mobile Duplication', 'ST_buildMobileReady4GHL')
    .addItem('Task 2: Landline Duplication', 'ST_buildLandlineReady4GHL')
    .addItem('Task 3: Email Duplication', 'ST_buildEmailReady4GHL')
    .addItem('Build All Ready Tabs', 'ST_buildAllReadyTabs_GHL')
    .addToUi();

  ui.createMenu("Prepend & Update GHL")
    .addItem("Task 1: Reorder Columns", "reorderColumnsToRequiredOrder")
    .addItem("Task 2: Prepend New Import", "prependNewImport")
    .addItem("Task 3: For GHL Updates", "createForGHLUpdates")
    .addItem("Task 4: Push Updates to ALL GHL Records", "pushUpdatesToGHL_fast")
    .addToUi();
}

