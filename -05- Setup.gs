/** * Setup.gs 
 * Handles the initialisation of subject sheets and the freeze/thaw state of the import data. 
 */
const Setup = {
  triggerCreateSubjectSheets: function() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const result = ui.alert('Confirm Setup', 'Are you sure you want to create the subject sheets?', ui.ButtonSet.YES_NO);
    if (result !== ui.Button.YES) {
      ui.alert('Request cancelled.');
      return;
    }
    
    const importSheet = ss.getSheetByName(CONFIG.IMPORT.targetSheetName);
    if (importSheet) {
      const currentStatus = importSheet.getRange(CONFIG.IMPORT.statusCell).getValue();
      if (currentStatus !== '🥶') {
        const freezeResult = ui.alert(
          'Freeze Import Data?',
          'The import sheet is currently active (thawed). It is recommended to freeze it before setup.\n\nWould you like to freeze the import data now?',
          ui.ButtonSet.YES_NO_CANCEL
        );
        if (freezeResult === ui.Button.YES) {
          this.freezeImportSheet();
        } else if (freezeResult === ui.Button.CANCEL || freezeResult === ui.Button.CLOSE) {
          ui.alert('Setup cancelled.');
          return;
        }
      }
    }
    
    ss.toast('Starting sheet generation...', 'Typeless');
    this._generateSubjectSheets();
  },
  
  freezeImportSheet: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.IMPORT.targetSheetName);
    
    if (!sheet) {
      ss.toast(`Error: Sheet "${CONFIG.IMPORT.targetSheetName}" not found.`, 'Typeless');
      return;
    }
    
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    
    if (lastCol === 0 || lastRow < CONFIG.IMPORT.anchorRowStart) return;
    
    const anchorRange = sheet.getRange(CONFIG.IMPORT.anchorRowStart, 1, CONFIG.IMPORT.anchorRowCount, lastCol);
    const anchorFormulas = anchorRange.getFormulas();
    let foundFormulas = [];
    
    for (let colIndex = 0; colIndex < lastCol; colIndex++) {
      let formula = "";
      let rowOffset = -1;
      if (anchorFormulas[0][colIndex] !== "") {
        formula = anchorFormulas[0][colIndex];
        rowOffset = 0;
      } else if (anchorFormulas[1][colIndex] !== "") {
        formula = anchorFormulas[1][colIndex];
        rowOffset = 1;
      }
      if (formula !== "") {
        foundFormulas.push({ formula: formula, colIndex: colIndex, rowOffset: rowOffset });
      }
    }
    
    if (foundFormulas.length === 0) {
      ss.toast("No formulas found in rows 6 or 7 to freeze.", "Typeless");
      return;
    }
    
    let backupData = [["Cell", "Formula", "Start Row", "Start Col", "End Col"]];
    for (let i = 0; i < foundFormulas.length; i++) {
      const current = foundFormulas[i];
      const startRow = CONFIG.IMPORT.anchorRowStart + current.rowOffset;
      const startCol = current.colIndex + 1;
      let endCol = (i < foundFormulas.length - 1) ? foundFormulas[i + 1].colIndex : lastCol;
      const cellAddress = sheet.getRange(startRow, startCol).getA1Notation();
      backupData.push([cellAddress, "'" + current.formula, startRow, startCol, endCol]);
    }
    
    let backupSheet = ss.getSheetByName(CONFIG.IMPORT.backupSheetName);
    if (!backupSheet) {
      backupSheet = ss.insertSheet(CONFIG.IMPORT.backupSheetName);
      backupSheet.hideSheet();
    } else {
      backupSheet.clear();
    }
    
    backupSheet.getRange(1, 1, backupData.length, 5).setValues(backupData);
    
    const rowsToFlatten = lastRow - CONFIG.IMPORT.anchorRowStart + 1;
    const freezeRange = sheet.getRange(CONFIG.IMPORT.anchorRowStart, 1, rowsToFlatten, lastCol);
    freezeRange.setValues(freezeRange.getValues());
    
    sheet.getRange(CONFIG.IMPORT.statusCell).setValue('🥶');
    ss.toast(`Successfully froze ${foundFormulas.length} formula(s).`, 'Typeless');
  },
  
  thawImportSheet: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.IMPORT.targetSheetName);
    const backupSheet = ss.getSheetByName(CONFIG.IMPORT.backupSheetName);
    
    if (!sheet || !backupSheet) {
      ss.toast("Error: Missing 'import' or backup sheet.", "Typeless");
      return;
    }
    
    const backupData = backupSheet.getDataRange().getValues();
    if (backupData.length <= 1) return;
    
    const maxRows = sheet.getMaxRows();
    let restoreCount = 0;
    
    for (let i = 1; i < backupData.length; i++) {
      const cellAddress = backupData[i][0];
      let formulaString = backupData[i][1].toString();
      if (formulaString.startsWith("'")) formulaString = formulaString.substring(1);
      
      const startRow = parseInt(backupData[i][2]);
      const startCol = parseInt(backupData[i][3]);
      const endCol = parseInt(backupData[i][4]);
      const numCols = endCol - startCol + 1;
      
      const rowsToClear = maxRows - startRow + 1;
      sheet.getRange(startRow, startCol, rowsToClear, numCols).clearContent();
      sheet.getRange(cellAddress).setFormula(formulaString);
      restoreCount++;
    }
    
    backupSheet.clear();
    sheet.getRange(CONFIG.IMPORT.statusCell).setValue('🫠');
    ss.toast(`Successfully thawed ${restoreCount} formula(s).`, 'Typeless');
  },
  
  _generateSubjectSheets: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const ksRange = ss.getRangeByName(CONFIG.SCOPE.keyStage);
    if (!ksRange) return;
    
    let ksVal = String(ksRange.getValue()).trim();
    if (ksVal.toUpperCase().startsWith('KS')) ksVal = ksVal.substring(2);
    const expectedNameHeader = ('nameks' + ksVal).toLowerCase();
    
    const subjectRange = ss.getRangeByName(CONFIG.SCOPE.subjectDetailsRange);
    if (!subjectRange) return;
    
    const data = subjectRange.getValues();
    if (data.length < 2) return;
    
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const codeColIdx = headers.indexOf('code');
    const nameColIdx = headers.indexOf(expectedNameHeader);
    
    if (codeColIdx === -1 || nameColIdx === -1) return;
    
    const subjects = [];
    for (let i = 1; i < data.length; i++) {
      const code = data[i][codeColIdx];
      const name = data[i][nameColIdx];
      if (code) subjects.push({ code: String(code).trim(), name: String(name).trim() });
    }
    
    const templateSheet = ss.getSheetByName('_Xx');
    if (!templateSheet) return;
    
    const templateProtections = templateSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    const templateProtection = templateProtections.length > 0 ? templateProtections[0] : null;
    
    const sheetsToCreate = subjects.filter(subject => !ss.getSheetByName(subject.code));
    const totalSheets = sheetsToCreate.length;
    
    if (totalSheets === 0) {
      ss.toast('All subject sheets already exist. Setup complete.', 'Typeless Setup');
      return;
    }
    
    let successCount = 0;
    sheetsToCreate.forEach((subject, index) => {
      ss.toast(`Generating sheet ${index + 1} of ${totalSheets} (${subject.code})...`, 'Setup Progress', 10);
      this._cloneTemplateForSubject(ss, templateSheet, templateProtection, subject);
      successCount++;
    });
    
    ss.toast(`Setup complete. ${successCount} sheets successfully generated.`, 'Typeless Setup', 5);
  },
  
  _cloneTemplateForSubject: function(ss, templateSheet, templateProtection, subject) {
    const newSheet = templateSheet.copyTo(ss).setName(subject.code);
    newSheet.getRange(2, 4).setValue(subject.name);
    
    const nameRangeStr = `${subject.code}!${CONFIG.SCOPE.targetSubjectNameRange}`;
    const targetNameRange = ss.getRangeByName(nameRangeStr);
    if (targetNameRange) targetNameRange.setValue(subject.name);
    
    if (templateProtection) {
      const newSheetProtection = newSheet.protect();
      newSheetProtection.setDescription(templateProtection.getDescription());
      newSheetProtection.setWarningOnly(templateProtection.isWarningOnly());
      const unprotectedRanges = templateProtection.getUnprotectedRanges();
      const newUnprotectedRanges = unprotectedRanges.map(range => newSheet.getRange(range.getA1Notation()));
      newSheetProtection.setUnprotectedRanges(newUnprotectedRanges);
    }
    
    newSheet.showSheet();
    SpreadsheetApp.flush();
  }
};