/**
 * Setup.gs
 * Handles the initialisation and generation of subject sheets from a master template.
 */

const Setup = {
  
  /**
   * Prompts the user to confirm the creation of subject sheets.
   * This can be linked directly to your Controller's custom menu.
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
   * Main loop to gather subject codes and clone the template for each.
   * @private
   */
  _generateSubjectSheets: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Fetch the list of subjects from the named range
    const subjectCodeRange = ss.getRangeByName('lookup_Subject');
    if (!subjectCodeRange) {
      SpreadsheetApp.getUi().alert('Error', 'Could not find the named range "lookup_Subject".', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // The Why: .flat() converts the 2D array to a 1D array, and .filter(String) removes any blank cells.
    const subjectCodes = subjectCodeRange.getValues().flat().filter(String);

    const templateSheet = ss.getSheetByName('_Xx');
    if (!templateSheet) {
      SpreadsheetApp.getUi().alert('Error', 'Template sheet "_Xx" not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // The Why: We fetch the template's protection settings ONCE outside the loop to save API calls.
    const templateProtections = templateSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    const templateProtection = templateProtections.length > 0 ? templateProtections[0] : null;

    let successCount = 0;

    subjectCodes.forEach(subjectCode => {
      // Safeguard: Only clone if a sheet with this subject code doesn't already exist.
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
   * @param {Spreadsheet} ss The active spreadsheet object.
   * @param {Sheet} templateSheet The master template sheet to copy.
   * @param {Protection} templateProtection The protection object from the template.
   * @param {string} subjectCode The code/name for the new sheet.
   */
  _cloneTemplateForSubject: function(ss, templateSheet, templateProtection, subjectCode) {
    
    // Clone the template and set its new name
    const newSheet = templateSheet.copyTo(ss).setName(subjectCode);
    
    // Insert the subject code into the designated cell (Row 2, Column 4)
    newSheet.getRange(2, 4).setValue(subjectCode);

    // Apply sheet protection settings if the template had them
    if (templateProtection) {
      const newSheetProtection = newSheet.protect();
      newSheetProtection.setDescription(templateProtection.getDescription());
      newSheetProtection.setWarningOnly(templateProtection.isWarningOnly());

      // Map unprotected ranges from the template to the new sheet
      const unprotectedRanges = templateProtection.getUnprotectedRanges();
      const newUnprotectedRanges = unprotectedRanges.map(range => {
        return newSheet.getRange(range.getA1Notation());
      });
      
      newSheetProtection.setUnprotectedRanges(newUnprotectedRanges);
    }

    // Show the newly created sheet
    newSheet.showSheet();
    
    // Force calculation to update the UI incrementally
    SpreadsheetApp.flush();
  }

};