/**
 * @fileoverview Controller.gs
 * Handles user interface lifecycles, dynamic menu rendering based on user roles,
 * modal dialog triggers, and RPC endpoints for the UCAS sidebar.
 */

/**
 * Standard trigger executed whenever the spreadsheet is opened.
 * Builds role-tailored menus based on the active user's email address.
 * @param {Object} e The open event object.
 */
function onOpen(e) {
  buildDynamicMenu();
}

/**
 * Constructs the custom Google Sheets menu.
 * Enforces role-based security by verifying the user's email against super-user
 * and report-specific permissions before displaying administrative actions.
 */
function buildDynamicMenu() {
  const ui = SpreadsheetApp.getUi();
  const email = Session.getActiveUser().getEmail();
  const menu = ui.createMenu('Typeless Reports');
  
  // If the script runs in an unauthenticated or limited context, prompt for initial authorisation.
  if (!email) {
    menu.addItem('Authorise Script', 'authoriseScript').addToUi();
    return;
  }
  
  const isSuperUser = CONFIG.AUTH.SUPER_USERS.includes(email);
  const isUcasUser = CONFIG.AUTH.REPORT_SPECIFIC.UCAS.includes(email);
  let menuHasItems = false;
  
  if (isSuperUser) {
    // Administrative tools restricted to Super Users
    menu.addItem('Setup Subject Sheets', 'triggerSetup');
    menu.addItem('Freeze Import Data', 'triggerFreeze');
    menu.addItem('Thaw Import Data', 'triggerThaw');
    menu.addSeparator();
    menu.addItem('Run Progress Review', 'triggerProgressReview');
    menu.addItem('Run Next Steps Summary', 'triggerNextStepsSummary');
    menu.addItem('Run EOY Report', 'triggerEoyReport');
    menu.addSeparator();
    menu.addItem('Run UCAS Collection (Sidebar)', 'showUcasSidebar');
    menuHasItems = true;
  } else if (isUcasUser) {
    // Tailored view for UCAS advisers to prevent accidental changes to global data
    menu.addItem('Run UCAS Collection (Sidebar)', 'showUcasSidebar');
    menuHasItems = true;
  }
  
  if (menuHasItems) {
    menu.addToUi();
  }
}

/**
 * Feedback handler called when an unauthenticated user triggers initial permission grants.
 */
function authoriseScript() {
  SpreadsheetApp.getUi().alert('Authorisation complete. Please refresh the page to see your custom menu.');
}

/** Triggers subject sheet initialisation via Setup.gs */
function triggerSetup() { Setup.triggerCreateSubjectSheets(); }

/** Freezes dynamic formulas on the import sheet to preserve calculation state */
function triggerFreeze() { Setup.freezeImportSheet(); }

/** Restores dynamic formulas on the import sheet for MIS data refresh */
function triggerThaw() { Setup.thawImportSheet(); }

// --- KS5 REPORT TRIGGERS ---
function triggerProgressReview() { showBatchModal('PROGRESS_REVIEW', 'Progress Reviews'); }
function triggerNextStepsSummary() { showBatchModal('NEXT_STEPS_SUMMARY', 'Next Steps Summaries'); }
function triggerEoyReport() { showBatchModal('EOY_REPORT', 'End of Year Reports'); }

/**
 * Opens the chunking modal dialog for heavy batch generation tasks.
 * Validates that import data is frozen beforehand to avoid reading unstable formulas.
 * @param {string} configKey Key corresponding to CONFIG.REPORTS.
 * @param {string} friendlyName Human-readable report title for the UI.
 */
function showBatchModal(configKey, friendlyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('import');
  
  // Ensure data stability: live MIS calculations must be flattened prior to generating documents.
  if (importSheet) {
    const status = importSheet.getRange('A1').getValue();
    if (status !== '🥶') {
      SpreadsheetApp.getUi().alert(
        'Validation Error', 
        'The import sheet must be frozen (🥶) before generating reports.', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
  }

  // Evaluate the HTML template with injected contextual properties
  const template = HtmlService.createTemplateFromFile('-10- BatchGeneration');
  template.configKey = configKey;
  template.friendlyName = friendlyName;
  
  const html = template.evaluate()
      .setWidth(450)
      .setHeight(380)
      .setTitle('Batch Generator');
      
  SpreadsheetApp.getUi().showModalDialog(html, 'Report Engine');
}

/**
 * Server-side initialisation for modal chunked processing.
 * Conducts a pre-flight data audit and constructs the target Drive folder.
 * @param {string} configKey Report configuration identifier.
 * @param {string|boolean} auditMode 'ignore', 'drop', or false if pre-flight audit check.
 * @return {Object} Status payload containing issues or folder details.
 */
function server_initBatch(configKey, auditMode) {
  const reportConfig = CONFIG.REPORTS[configKey];
  const payload = DataService.buildStudentDataPayload(reportConfig);

  if (payload.length === 0) return { error: "No student data found." };

  // Conduct the audit check only if the user has not yet acknowledged missing data
  if (!auditMode) {
    const studentsWithIssues = payload.filter(s => s.auditIssues && s.auditIssues.length > 0);
    if (studentsWithIssues.length > 0) {
      const issuesList = studentsWithIssues.map(s => `<b>${s.name}</b>: ${s.auditIssues.join(' | ')}`);
      return {
        status: 'audit_warning',
        issues: issuesList,
        totalStudents: payload.length
      };
    }
  }

  // Create batch folder once audit is resolved or bypassed
  const folderId = DocumentBuilder.createBatchFolder(reportConfig, payload[0]);
  
  return {
    status: 'ready',
    folderId: folderId,
    folderUrl: `https://drive.google.com/drive/folders/${folderId}`,
    totalStudents: payload.length
  };
}

/**
 * Server-side chunk execution called cyclically by the modal dialog.
 * Slices student data into manageable chunks to stay well beneath GAS timeout limits.
 * @param {string} configKey Report configuration key.
 * @param {string} folderId Destination folder ID.
 * @param {number} startIndex Zero-based start index of the chunk.
 * @param {number} chunkSize Number of students to process.
 * @param {string} auditMode Audit flag ('ignore' or 'drop').
 * @return {Object} Simple success confirmation.
 */
function server_processChunk(configKey, folderId, startIndex, chunkSize, auditMode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportConfig = CONFIG.REPORTS[configKey];
  
  // Re-build student payload dynamically in memory to maintain lean execution state
  const payload = DataService.buildStudentDataPayload(reportConfig);
  const chunk = payload.slice(startIndex, startIndex + chunkSize);

  ss.toast(`Merging chunk: ${startIndex + 1} to ${startIndex + chunk.length}...`, 'Background Engine');
  
  // Delegate document creation to the builder module
  DocumentBuilder.generateChunk(reportConfig, chunk, folderId, auditMode);
  
  return { success: true };
}

/**
 * Opens the dedicated HTML sidebar for on-demand UCAS Reference collation.
 */
function showUcasSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('-09- Sidebar')
      .setTitle('UCAS References')
      .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Fetches the master student directory to populate the searchable list in the sidebar.
 * Sorts students alphabetically by surname/forename for quick lookup.
 * @return {Array<Object>} Lightweight student descriptor objects.
 */
function sidebarGetStudentList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentMap = DataService._getMasterStudentList(ss);
  
  const list = Object.values(studentMap).map(s => ({
    name: s.name,
    adNo: s.adNo,
    reg: s.reg,
    earlyApp: s.earlyApp
  }));
  
  // Alphabetical sort ensures consistent navigation in the client-side list
  list.sort((a, b) => a.name.localeCompare(b.name));
  return list;
}

/**
 * Fetches the comprehensive cohort UCAS dataset for client-side caching.
 * Invoked asynchronously in the background upon sidebar initialisation.
 * @return {Object<string, Object>|Object} Map of admission numbers to reference payloads or error object.
 */
function sidebarGetCohortUcasData() {
  try {
    return DataService.getCohortUcasData(CONFIG.REPORTS.UCAS_REFERENCE);
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Fetches structured, multi-section preview data for a single student.
 * Enables the coordinator to review Section 2 (Tutor), Section 3 (Subjects),
 * and Section 4 (Predictions) prior to running a merge.
 * @param {string|number} adno The admission number of the selected student.
 * @return {Object} Structured preview payload or an error descriptor.
 */
function sidebarGetUcasPreview(adno) {
  try {
    const result = DataService.getUcasPreviewText(CONFIG.REPORTS.UCAS_REFERENCE, adno);
    if (!result) return { error: "Student not found or no UCAS data available." };
    return result;
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Executes an on-demand merge using student reference payloads supplied directly
 * from browser memory (client cache). Bypasses sheet re-reading for maximum performance.
 * @param {Array<Object>} cachedStudentPayloads Array of student data objects from the sidebar.
 * @return {Object} Operation summary containing document URL, folder URL, and count.
 */
function sidebarRunUcasMergeFromCache(cachedStudentPayloads) {
  try {
    if (!Array.isArray(cachedStudentPayloads) || cachedStudentPayloads.length === 0) {
      return { error: "No student records provided for merge." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast(`Generating references for ${cachedStudentPayloads.length} student(s)...`, 'UCAS Engine');

    return DocumentBuilder.generateUcasDocsFromCache(cachedStudentPayloads);
  } catch (e) {
    return { error: "System Error: " + e.message };
  }
}

/**
 * Executes an on-demand merge for one or more selected students.
 * Returns direct document links when processing an individual student,
 * or folder links when handling a small group of early applicants.
 * @param {string} adnoString Comma-separated list of target admission numbers.
 * @return {Object} Operation summary containing document/folder links and count.
 */
function sidebarRunUcasMerge(adnoString) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Sanitise and strip empty entries from the comma-separated parameter
    const adnos = adnoString.split(',').map(s => s.trim()).filter(Boolean);
    
    let payload = DataService.buildStudentDataPayload(CONFIG.REPORTS.UCAS_REFERENCE);
    
    // Filter the full cohort down to the requested admission numbers
    payload = payload.filter(s => 
      adnos.includes(String(s.adNo)) || adnos.includes(String(s.adNo).padStart(6, '0'))
    );
    
    if (payload.length === 0) return { error: "No matching students found." };
    
    ss.toast(`Generating UCAS references for ${payload.length} student(s)...`, 'Typeless');
    const result = DocumentBuilder.generateBatch(CONFIG.REPORTS.UCAS_REFERENCE, payload);
    
    return {
      success: true,
      count: payload.length,
      folderUrl: result.folderUrl,
      docUrl: payload.length === 1 ? result.lastDocUrl : null
    };
  } catch (e) {
    return { error: "System Error: " + e.message };
  }
}

/**
 * Shared execution logic for modal-less standard batch runs.
 * Retained for backwards compatibility with non-chunked workflows.
 * @private
 * @param {Object} reportConfig Target report configuration object.
 * @param {string} reportFriendlyName Human-readable name for toasts and prompts.
 */
function _runReportBatch(reportConfig, reportFriendlyName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('import');
  
  if (importSheet) {
    const status = importSheet.getRange('A1').getValue();
    if (status !== '🥶') {
      ui.alert('Validation Error', 'The import sheet must be frozen (🥶) before generating reports.', ui.ButtonSet.OK);
      return;
    }
  }
  
  const batchPrompt = ui.prompt(
    'Batch Run', 
    `Generate ${reportFriendlyName}?\n\nEnter a number to run a test batch, or leave blank to run the whole cohort:`, 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (batchPrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  ss.toast(`Gathering and auditing data for ${reportFriendlyName}...`, 'Typeless');
  let payload = DataService.buildStudentDataPayload(reportConfig);
  
  if (payload.length === 0) {
    ui.alert('Error', 'No student data found. Please check the master list and subject sheets.', ui.ButtonSet.OK);
    return;
  }
  
  const batchInput = batchPrompt.getResponseText().trim();
  if (batchInput !== '' && !isNaN(batchInput)) {
    const limit = parseInt(batchInput, 10);
    if (limit > 0) payload = payload.slice(0, limit);
  }
  
  const studentsWithIssues = payload.filter(s => s.auditIssues && s.auditIssues.length > 0);
  if (studentsWithIssues.length > 0) {
    let issueText = `Missing data detected for ${studentsWithIssues.length} student(s) in this run.\n\nAffected Students:\n`;
    studentsWithIssues.forEach(stu => {
      issueText += `• ${stu.name}: ${stu.auditIssues.join(' | ')}\n`;
    });
    issueText += `\nDo you want to generate the reports with missing data anyway?`;
    
    const proceed = ui.alert('Pre-Flight Data Warning', issueText, ui.ButtonSet.YES_NO);
    if (proceed !== ui.Button.YES) {
      ss.toast('Generation cancelled by user.', 'Typeless');
      return;
    }
  }
  
  ss.toast(`Generating documents for ${payload.length} students...`, 'Typeless');
  const batchResult = DocumentBuilder.generateBatch(reportConfig, payload);
  ui.alert('Merge Complete', `Documents generated successfully.\nFolder URL: ${batchResult.folderUrl}`, ui.ButtonSet.OK);
}