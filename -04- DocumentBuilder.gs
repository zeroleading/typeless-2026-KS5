/**
 * DocumentBuilder.gs
 * Handles the generation of Google Docs by cloning templates and injecting data.
 */

const DocumentBuilder = {

  /**
   * Main function to generate a batch of reports.
   * @param {Object} reportConfig The specific report configuration from CONFIG.REPORTS.
   * @param {Array<Object>} studentPayload The processed data array from DataService.
   */
  generateBatch: function(reportConfig, studentPayload) {
    const templateFile = DriveApp.getFileById(reportConfig.templateId);
    const outputFolder = DriveApp.getFolderById(CONFIG.GLOBAL.OUTPUT_FOLDER_ID);
    
    // Create a specific subfolder for this run to keep things tidy
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "yyyy-MM-dd");
    const batchFolder = outputFolder.createFolder(`${reportConfig.name} - ${dateStr}`);

    // Loop through each student in our payload
    studentPayload.forEach(student => {
      // The Why: We only generate a document if the student actually has subjects.
      // This prevents generating blank reports for students who might have left.
      if (student.subjects && student.subjects.length > 0) {
        this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
      }
    });

    return batchFolder.getId(); // Return the ID so the Controller can alert the user
  },

  /**
   * Helper: Builds a single Google Doc for a single student.
   * @private
   */
  _buildSingleDocument: function(student, templateFile, destinationFolder, reportName) {
    const docName = `${student.adNo} - ${student.name} - ${reportName}`;
    const newFile = templateFile.makeCopy(docName, destinationFolder);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();
    const header = doc.getHeader();
    const footer = doc.getFooter();

    // 1. Replace Global Variables (using your old _Name_ format)
    this._replaceGlobalPlaceholders(body, student);
    if (header) this._replaceGlobalPlaceholders(header, student);
    if (footer) this._replaceGlobalPlaceholders(footer, student);

    // 2. Process the dynamic Subject Table
    this._populateSubjectTable(body, student.subjects);

    // Save and close to ensure changes are flushed to Google's servers
    doc.saveAndClose();
  },

  /**
   * Helper: Replaces standard text placeholders.
   * @private
   */
  _replaceGlobalPlaceholders: function(element, student) {
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "MMMM yyyy");
    
    element.replaceText('_Name_', student.name);
    element.replaceText('_Reg_', student.reg);
    element.replaceText('_AdNo_', student.adNo);
    element.replaceText('_Tutor_', student.tutor);
    element.replaceText('_Date_', dateStr);
  },

  /**
   * Helper: Finds the template table and clones the row for each subject.
   * @private
   */
  _populateSubjectTable: function(body, subjects) {
    const tables = body.getTables();
    let targetTable = null;
    let templateRowIndex = -1;

    // 1. Locate the correct table by searching for our anchor tag: {{subjectName}}
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      for (let r = 0; r < table.getNumRows(); r++) {
        const rowText = table.getRow(r).getText();
        if (rowText.includes('{{subjectName}}')) {
          targetTable = table;
          templateRowIndex = r;
          break;
        }
      }
      if (targetTable) break; // Stop looking once we find it
    }

    // The Why: If the template doesn't have the table or the tag, we gracefully exit 
    // to prevent crashing the entire batch process.
    if (!targetTable || templateRowIndex === -1) return;

    const templateRow = targetTable.getRow(templateRowIndex);

    // 2. Loop through the student's subjects and append new rows
    subjects.forEach(subject => {
      const newRow = targetTable.appendTableRow(templateRow.copy());
      
      // Replace the tags in the newly copied row
      newRow.replaceText('{{subjectName}}', subject.subjectName || '-');
      newRow.replaceText('{{tg}}', subject.tg || '-');
      newRow.replaceText('{{eoy}}', subject.eoy || '-');
      newRow.replaceText('{{ucas}}', subject.ucas || '-');
      newRow.replaceText('{{rank}}', subject.rank || '-');
      newRow.replaceText('{{ucasRef}}', subject.ucasRef || '-');
      newRow.replaceText('{{crnt}}', subject.crnt || '-');
      newRow.replaceText('{{nextSteps}}', subject.nextSteps || '-');
      newRow.replaceText('{{att}}', subject.att || '-');
    });

    // 3. Clean up: Remove the original template row containing the curly braces
    targetTable.removeRow(templateRowIndex);
  }

};