/**
 * Setup.gs
 * Handles the initialisation of subject sheets and the freeze/thaw state of the import data.
 */

const Setup = {
  
  /**
   * Prompts the user to confirm the creation of subject sheets.
   */
  triggerCreateSubjectSheets: function() {
    const ui = SpreadsheetApp.getUi();
    
    const result = ui.alert(
      'Confirm Setup',
      'Are you sure you want to create the subject sheets?',
      ui.ButtonSet.YES_NO
    );

    if (result === ui.Button.YES) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Starting sheet generation...', 'Typeless');
      this._generateSubjectSheets();
    } else {
      ui.alert('Request cancelled.');
    }
  },

  /**
   * PHASE A: Freeze the sheet and calculate column ownership based on next formula
   */
  freezeImportSheet: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.IMPORT.targetSheetName);
    
    if (!sheet) {
      console.log(`Sheet "${CONFIG.IMPORT.targetSheetName}" not found.`);
      ss.toast(`Error: Sheet "${CONFIG.IMPORT.targetSheetName}" not found.`, 'Typeless');
      return;
    }

    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    
    if (lastCol === 0 || lastRow < CONFIG.IMPORT.anchorRowStart) {
      console.log("No data found in or below row 6 to freeze.");
      return;
    }

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
        foundFormulas.push({
          formula: formula,
          colIndex: colIndex,
          rowOffset: rowOffset
        });
      }
    }

    if (foundFormulas.length === 0) {
      console.log("No formulas found in rows 6 or 7 to freeze.");
      ss.toast("No formulas found in rows 6 or 7 to freeze.", "Typeless");
      return;
    }

    let backupData = [["Cell", "Formula", "Start Row", "Start Col", "End Col"]];

    for (let i = 0; i < foundFormulas.length; i++) {
      const current = foundFormulas[i];
      const startRow = CONFIG.IMPORT.anchorRowStart + current.rowOffset; // Calculate exact row (6 or 7)
      const startCol = current.colIndex + 1;
      let endCol;

      if (i < foundFormulas.length - 1) {
        endCol = foundFormulas[i + 1].colIndex; 
      } else {
        endCol = lastCol;
      }

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
    console.log(`Successfully froze ${foundFormulas.length} formula(s).`);
    ss.toast(`Successfully froze ${foundFormulas.length} formula(s).`, 'Typeless');
  },

  /**
   * PHASE B: Thaw formulas and surgically clear their owned columns
   */
  thawImportSheet: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.IMPORT.targetSheetName);
    const backupSheet = ss.getSheetByName(CONFIG.IMPORT.backupSheetName);
    
    if (!sheet || !backupSheet) {
      console.log("Missing 'import' or backup sheet. Cannot thaw.");
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
      
      if (formulaString.startsWith("'")) {
        formulaString = formulaString.substring(1);
      }
      
      // Read the specific row and column bounds for this formula
      const startRow = parseInt(backupData[i][2]);
      const startCol = parseInt(backupData[i][3]);
      const endCol = parseInt(backupData[i][4]);
      const numCols = endCol - startCol + 1;
      
      // SURGICAL CLEAR: Now starts dynamically at Row 6 OR Row 7
      const rowsToClear = maxRows - startRow + 1;
      sheet.getRange(startRow, startCol, rowsToClear, numCols).clearContent();
      
      sheet.getRange(cellAddress).setFormula(formulaString);
      restoreCount++;
    }
    
    backupSheet.clear();
    sheet.getRange(CONFIG.IMPORT.statusCell).setValue('🫠');
    console.log(`Successfully thawed ${restoreCount} formula(s).`);
    ss.toast(`Successfully thawed ${restoreCount} formula(s).`, 'Typeless');
  },

  /**
   * Main loop to gather subject codes and clone the template for each.
   * @private
   */
  _generateSubjectSheets: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Dynamically retrieve the subject details range using our Config file
    const rangeName = CONFIG.SCOPE.subjectDetailsRange;
    const subjectRange = ss.getRangeByName(rangeName);
    
    if (!subjectRange) {
      SpreadsheetApp.getUi().alert('Error', `Could not find the named range "${rangeName}".`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const data = subjectRange.getValues();
    
    // Safeguard: Ensure we have at least a header row and one row of data
    if (data.length < 2) return; 

    // Find the dynamic index of the 'code' column
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const codeColIdx = headers.indexOf('code');

    if (codeColIdx === -1) {
      SpreadsheetApp.getUi().alert('Error', 'Could not find a column headed "code" in the subject details range.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Extract the subject codes, skipping the header row and ignoring blank cells
    const subjectCodes = [];
    for (let i = 1; i < data.length; i++) {
      const code = data[i][codeColIdx];
      if (code) {
        subjectCodes.push(String(code).trim());
      }
    }

    const templateSheet = ss.getSheetByName('_Xx');
    if (!templateSheet) {
      SpreadsheetApp.getUi().alert('Error', 'Template sheet "_Xx" not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const templateProtections = templateSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    const templateProtection = templateProtections.length > 0 ? templateProtections[0] : null;

    let successCount = 0;

    subjectCodes.forEach(subjectCode => {
      if (!ss.getSheetByName(subjectCode)) {
        this._cloneTemplateForSubject(ss, templateSheet, templateProtection, subjectCode);
        successCount++;
      }
    });

    ss.toast(`Setup complete. ${successCount} sheets successfully generated.`, 'Typeless');
  },

  /**
   * Clones the template sheet, renames it, and applies protections.
   * @private
   */
  _cloneTemplateForSubject: function(ss, templateSheet, templateProtection, subjectCode) {
    const newSheet = templateSheet.copyTo(ss).setName(subjectCode);
    newSheet.getRange(2, 4).setValue(subjectCode);

    if (templateProtection) {
      const newSheetProtection = newSheet.protect();
      newSheetProtection.setDescription(templateProtection.getDescription());
      newSheetProtection.setWarningOnly(templateProtection.isWarningOnly());

      const unprotectedRanges = templateProtection.getUnprotectedRanges();
      const newUnprotectedRanges = unprotectedRanges.map(range => {
        return newSheet.getRange(range.getA1Notation());
      });
      
      newSheetProtection.setUnprotectedRanges(newUnprotectedRanges);
    }

    newSheet.showSheet();
    SpreadsheetApp.flush();
  }

};