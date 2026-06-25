/** * Controller.gs 
 * Handles the user interface, custom menus, and authorisation routing. 
 */
function onOpen(e) {
  buildDynamicMenu();
}

function buildDynamicMenu() {
  const ui = SpreadsheetApp.getUi();
  const email = Session.getActiveUser().getEmail();
  const menu = ui.createMenu('Typeless Reports');
  
  if (!email) {
    menu.addItem('Authorise Script', 'authoriseScript').addToUi();
    return;
  }
  
  const isSuperUser = CONFIG.AUTH.SUPER_USERS.includes(email);
  let menuHasItems = false;
  
  if (isSuperUser) {
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
  }
  
  if (menuHasItems) {
    menu.addToUi();
  }
}

function authoriseScript() {
  SpreadsheetApp.getUi().alert('Authorisation complete. Please refresh the page to see your custom menu.');
}

function triggerSetup() { Setup.triggerCreateSubjectSheets(); }
function triggerFreeze() { Setup.freezeImportSheet(); }
function triggerThaw() { Setup.thawImportSheet(); }

// --- KS5 REPORT TRIGGERS ---
function triggerProgressReview() { showBatchModal('PROGRESS_REVIEW', 'Progress Reviews'); }
function triggerNextStepsSummary() { showBatchModal('NEXT_STEPS_SUMMARY', 'Next Steps Summaries'); }
function triggerEoyReport() { showBatchModal('EOY_REPORT', 'End of Year Reports'); }

/**
 * Opens the new Chunking Modal for heavy report generation.
 */
function showBatchModal(configKey, friendlyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('import');
  
  if (importSheet) {
    const status = importSheet.getRange('A1').getValue();
    if (status !== '🥶') {
      SpreadsheetApp.getUi().alert('Validation Error', 'The import sheet must be frozen (🥶) before generating reports.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  }

  // Use HtmlTemplate to pass variables to the HTML file
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
 * Called by the Modal (Step 1): Prepares the folder and audits the data.
 */
function server_initBatch(configKey, forceProceed) {
  const reportConfig = CONFIG.REPORTS[configKey];
  const payload = DataService.buildStudentDataPayload(reportConfig);

  if (payload.length === 0) return { error: "No student data found." };

  // 1. Audit Check
  if (!forceProceed) {
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

  // 2. Folder Creation
  const folderId = DocumentBuilder.createBatchFolder(reportConfig, payload[0]);
  
  return {
    status: 'ready',
    folderId: folderId,
    folderUrl: `https://drive.google.com/drive/folders/${folderId}`,
    totalStudents: payload.length
  };
}

/**
 * Called by the Modal (Step 3 Loop): Processes a specific chunk of students.
 */
function server_processChunk(configKey, folderId, startIndex, chunkSize) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportConfig = CONFIG.REPORTS[configKey];
  
  // Re-build payload dynamically (fast and keeps memory lean)
  const payload = DataService.buildStudentDataPayload(reportConfig);
  
  // Slice out just the 10 students requested
  const chunk = payload.slice(startIndex, startIndex + chunkSize);

  ss.toast(`Merging chunk: ${startIndex + 1} to ${startIndex + chunk.length}...`, 'Background Engine');
  
  // Send to builder
  DocumentBuilder.generateChunk(reportConfig, chunk, folderId);
  
  return { success: true };
}

/**
 * Opens the HTML sidebar for UCAS Reference operations.
 */
function showUcasSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('-09- Sidebar')
      .setTitle('UCAS References')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Called by the Sidebar on load to populate the searchable student list.
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
  
  list.sort((a, b) => a.name.localeCompare(b.name));
  return list;
}

/**
 * Called by the Sidebar to fetch a plain-text preview of the references.
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
 * Called by the Sidebar to run a specific merge for requested Admission Numbers.
 */
function sidebarRunUcasMerge(adnoString) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const adnos = adnoString.split(',').map(s => s.trim()).filter(s => s);
    
    let payload = DataService.buildStudentDataPayload(CONFIG.REPORTS.UCAS_REFERENCE);
    
    payload = payload.filter(s => 
      adnos.includes(String(s.adNo)) || adnos.includes(String(s.adNo).padStart(6, '0'))
    );
    
    if (payload.length === 0) return "Error: No matching students found.";
    
    const studentsWithIssues = payload.filter(s => s.auditIssues && s.auditIssues.length > 0);
    if (studentsWithIssues.length > 0) {
      const names = studentsWithIssues.map(s => s.name).join(', ');
      return `Error: Missing UCAS data detected for ${names}. Please resolve on the subject sheets before merging.`;
    }
    
    ss.toast(`Generating UCAS references for ${payload.length} student(s)...`, 'Typeless');
    const folderId = DocumentBuilder.generateBatch(CONFIG.REPORTS.UCAS_REFERENCE, payload);
    
    return `https://drive.google.com/drive/folders/${folderId}`;
  } catch (e) {
    return "System Error: " + e.message;
  }
}

/** * Shared execution logic for all report types. 
 * @private 
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
  
  const batchPrompt = ui.prompt('Batch Run', `Generate ${reportFriendlyName}?\n\nEnter a number to run a test batch, or leave blank to run the whole cohort:`, ui.ButtonSet.OK_CANCEL);
  
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
  const folderId = DocumentBuilder.generateBatch(reportConfig, payload);
  ui.alert('Merge Complete', `Documents generated successfully.\nFolder ID: ${folderId}`, ui.ButtonSet.OK);
}