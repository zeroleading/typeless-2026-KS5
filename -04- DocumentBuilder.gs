/**
 * @fileoverview DocumentBuilder.gs
 * Handles copying Google Docs templates, injecting cohort/student metadata,
 * rendering conditional UCAS tables, and compiling sectional reference narratives.
 */

const DocumentBuilder = {
  
  // ---------------------------------------------------------------------------
  // 1. Chunking Engine Methods
  // ---------------------------------------------------------------------------

  /**
   * Creates a timestamped batch folder within Google Drive for the current run.
   * Directs output to report-specific folders when configured (e.g. for UCAS references).
   * @param {Object} reportConfig Target report profile.
   * @param {Object} sampleStudent A sample student payload used to derive year/collection details.
   * @return {string} Created folder ID.
   */
  createBatchFolder: function(reportConfig, sampleStudent) {
    // Determine target folder: report-specific override takes precedence over global default
    const targetFolderId = reportConfig.outputFolderId || CONFIG.GLOBAL.OUTPUT_FOLDER_ID;
    const outputFolder = DriveApp.getFolderById(targetFolderId);
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "yyyy-MM-dd");
    
    const academicYear = sampleStudent?.academicYear || '';
    const collection = sampleStudent?.collection || '';
    const yearGroup = sampleStudent?.yearGroup || '';
    
    let folderName = `${academicYear} ${collection} ${yearGroup} ${dateStr}`.trim();
    if (reportConfig.name === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) folderName += " next-steps";
    if (reportConfig.name === CONFIG.REPORTS.EOY_REPORT.name) folderName += " EOY";
    if (reportConfig.name === CONFIG.REPORTS.UCAS_REFERENCE.name) folderName += " ucas-refs";
    
    const batchFolder = outputFolder.createFolder(folderName);
    return batchFolder.getId();
  },

  /**
   * Processes a sliced chunk of student records and merges them into documents.
   * @param {Object} reportConfig Active report profile.
   * @param {Array<Object>} chunkPayload Subset of student objects to process.
   * @param {string} folderId Destination folder ID.
   * @param {string} auditMode Audit flag ('ignore' or 'drop').
   */
  generateChunk: function(reportConfig, chunkPayload, folderId, auditMode = 'ignore') {
    const templateFile = DriveApp.getFileById(reportConfig.templateId);
    const batchFolder = DriveApp.getFolderById(folderId);
    
    chunkPayload.forEach((student) => {
      // Filter out incomplete subjects if the user elected to drop them during the audit
      let validSubjects = student.subjects || [];
      if (auditMode === 'drop') {
        validSubjects = validSubjects.filter(subj => this._isSubjectComplete(subj, reportConfig.name));
      }

      // Only generate the document if at least one valid subject remains
      if (validSubjects.length > 0) {
        student.subjects = validSubjects;
        this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
      }
    });
  },

  // ---------------------------------------------------------------------------
  // 2. Batch & On-Demand Merging Methods
  // ---------------------------------------------------------------------------

  /**
   * Generates documents for an arbitrary array of student records.
   * Captures the direct URL of the created document when processing an individual
   * student, enabling immediate review from the sidebar interface.
   * @param {Object} reportConfig Target report profile.
   * @param {Array<Object>} studentPayload Student records to merge.
   * @param {string} auditMode Audit flag ('ignore' or 'drop').
   * @return {Object} Metadata containing folder ID, folder URL, last document URL, and document count.
   */
  generateBatch: function(reportConfig, studentPayload, auditMode = 'ignore') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateFile = DriveApp.getFileById(reportConfig.templateId);
    
    // Respect report-specific output folders where defined
    const targetFolderId = reportConfig.outputFolderId || CONFIG.GLOBAL.OUTPUT_FOLDER_ID;
    const outputFolder = DriveApp.getFolderById(targetFolderId);
    
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "yyyy-MM-dd");
    const academicYear = studentPayload[0]?.academicYear || '';
    const collection = studentPayload[0]?.collection || '';
    const yearGroup = studentPayload[0]?.yearGroup || '';
    
    let folderName = `${academicYear} ${collection} ${yearGroup} ${dateStr}`.trim();
    if (reportConfig.name === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      folderName += " next-steps";
    } else if (reportConfig.name === CONFIG.REPORTS.UCAS_REFERENCE.name) {
      folderName += " ucas-refs";
    }
    
    const batchFolder = outputFolder.createFolder(folderName);
    const totalStudents = studentPayload.length;
    let lastCreatedDocUrl = '';
    let generatedCount = 0;
    
    studentPayload.forEach((student, index) => {
      // Filter out incomplete subjects if requested
      let validSubjects = student.subjects || [];
      if (auditMode === 'drop') {
        validSubjects = validSubjects.filter(subj => this._isSubjectComplete(subj, reportConfig.name));
      }

      if (validSubjects.length > 0) {
        student.subjects = validSubjects;
        ss.toast(`Merging document ${index + 1} of ${totalStudents}...\n(${student.name})`, 'Progress Tracker', 10);
        const docFile = this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
        lastCreatedDocUrl = docFile.getUrl();
        generatedCount++;
      }
    });
    
    return {
      folderId: batchFolder.getId(),
      folderUrl: batchFolder.getUrl(),
      lastDocUrl: lastCreatedDocUrl,
      generatedCount: generatedCount
    };
  },

  /**
   * Evaluates whether a subject record satisfies completeness criteria for a given report.
   * @private
   * @param {Object} subj The subject assessment object.
   * @param {string} reportName The canonical name of the active report.
   * @return {boolean} True if complete, false otherwise.
   */
  _isSubjectComplete: function(subj, reportName) {
    if (reportName === CONFIG.REPORTS.EOY_REPORT.name) {
      return subj.eoy !== '';
    } else if (reportName === CONFIG.REPORTS.UCAS_REFERENCE.name) {
      // EOY and rank are supplementary; completeness strictly requires a predicted grade and reference narrative
      return subj.ucas !== '' && subj.ucasRef !== '';
    } else {
      // Progress Reviews and Next Steps Summaries require CRNT, CI indicators, and at least one Next Step
      return subj.crnt !== '' &&
             subj.ci1 !== '' &&
             subj.ci2 !== '' &&
             subj.ci3 !== '' &&
             subj.ci4 !== '' &&
             (subj.nextSteps1 !== '' || subj.nextSteps2 !== '');
    }
  },
  
  /**
   * Constructs an individual Google Doc from a template for a single student.
   * Replaces global metadata across body, headers, and footers, and invokes
   * conditional table population.
   * @private
   * @param {Object} student Aggregated student data object.
   * @param {GoogleAppsScript.Drive.File} templateFile Template doc file.
   * @param {GoogleAppsScript.Drive.Folder} destinationFolder Target Drive folder.
   * @param {string} reportName Canonical report title.
   * @return {GoogleAppsScript.Drive.File} The newly created Google Doc file.
   */
  _buildSingleDocument: function(student, templateFile, destinationFolder, reportName) {
    // Format admission numbers to 6 digits to guarantee consistent sorting and identification
    const paddedAdNo = String(student.adNo).padStart(6, '0');
    
    let fileName = `${student.reg} ${student.name} ${paddedAdNo} ${student.shortName || ''}`.trim();
    if (reportName === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      fileName += " next-steps";
    } else if (reportName === CONFIG.REPORTS.UCAS_REFERENCE.name) {
      fileName += " ucas-ref";
    }
    
    const newDocFile = templateFile.makeCopy(fileName, destinationFolder);
    const newDoc = DocumentApp.openById(newDocFile.getId());
    
    const body = newDoc.getBody();
    const header = newDoc.getHeader();
    const footer = newDoc.getFooter();
    
    // Modular helper replacing common placeholder variables across all doc sections
    const replaceGlobalsInSection = (section) => {
      if (!section) return;
      section.replaceText('_Name_', student.name || '');
      section.replaceText('_Reg_', student.reg || '');
      section.replaceText('_AdNo_', paddedAdNo);
      section.replaceText('_Tutor_', student.tutor || '');
      section.replaceText('_Date_', Utilities.formatDate(new Date(), "Europe/London", "MMMM yyyy"));
      section.replaceText('_YearGroup_', student.yearGroup || '');
      section.replaceText('_Collection_', student.collection || '');
      section.replaceText('_Until_', student.until || '');
      
      if (student.tutorInfo) {
        section.replaceText('_AttTpAs_', student.tutorInfo.attTpAs || '-');
        section.replaceText('_LatesTpAs_', student.tutorInfo.latesTpAs || '0');
      }
    };
    
    replaceGlobalsInSection(body);
    replaceGlobalsInSection(header);
    replaceGlobalsInSection(footer);
    
    // Inject UCAS sectional reference narratives when running UCAS references
    if (reportName === CONFIG.REPORTS.UCAS_REFERENCE.name) {
      this._injectUcasSections(body, student);
    }
    
    // Handle conditional predictions table for standard Progress Reviews
    this._processConditionalUcasTable(body, student, reportName);
    
    // Populate standard tabular subjects
    this._populateSubjectTable(body, student.subjects);
    
    newDoc.saveAndClose();
    return newDocFile;
  },

  /**
   * Injects tutor extenuating circumstances, formatted subject references,
   * and plain-text predictions into the UCAS template document.
   * @private
   * @param {GoogleAppsScript.Document.Body} body Document body.
   * @param {Object} student Student record.
   */
  _injectUcasSections: function(body, student) {
    // Section 2: Tutor extenuating circumstances / contextual narrative
    const tutorRef = (student.tutorInfo && student.tutorInfo.ucasRef)
      ? String(student.tutorInfo.ucasRef).trim()
      : 'None declared.';
    body.replaceText('_Collected Tutor Reference_', tutorRef);

    // Section 3: Subject suitability references
    const subjectRefs = DataService.formatUcasSubjectReferences(student.subjects);
    body.replaceText('_Collected References_', subjectRefs || 'No subject references recorded.');

    // Section 4: Plain-text predicted grades
    const predictions = DataService.formatUcasPredictions(student.subjects);
    body.replaceText('_Collected Predictions_', predictions || 'No predicted grades recorded.');
  },

  /**
   * Conditionally injects or strips the UCAS grades table.
   * Displays only for Year 12 cohorts during 'Progress Review B' where grades exist.
   * @private
   * @param {GoogleAppsScript.Document.Body} body Document body.
   * @param {Object} student Student record.
   * @param {string} reportName Active report identifier.
   */
  _processConditionalUcasTable: function(body, student, reportName) {
    if (reportName !== CONFIG.REPORTS.PROGRESS_REVIEW.name) return;
    
    const isYear12 = String(student.yearGroup).includes('12');
    const isPRB = String(student.collection).trim() === 'Progress Review B';
    
    // Filter down to subjects containing an explicit predicted grade
    const ucasSubjects = student.subjects.filter(subj => subj.ucas && String(subj.ucas).trim() !== '');
    
    // Condition is satisfied only for Year 12, Progress Review B, with at least one grade present
    const shouldDisplay = isYear12 && isPRB && ucasSubjects.length > 0;
    
    // 1. Locate the UCAS table template row
    const tables = body.getTables();
    let ucasTable = null;
    let templateRow = null;
    let templateRowIndex = -1;
    
    for (let t = 0; t < tables.length; t++) {
      const table = tables[t];
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        if (row.getText().includes('{{ucasSubjectName}}')) {
          ucasTable = table;
          templateRow = row.copy();
          templateRowIndex = r;
          break;
        }
      }
      if (ucasTable) break;
    }
    
    // 2. Locate the heading and preceding horizontal separator rule
    let headingParagraph = null;
    let hrToRemove = null;
    
    const headingSearch = body.findText('_UcasHeading_');
    if (headingSearch) {
      headingParagraph = headingSearch.getElement().getParent();
      if (headingParagraph && headingParagraph.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const prevSibling = headingParagraph.getPreviousSibling();
        if (prevSibling) {
          if (prevSibling.getType() === DocumentApp.ElementType.HORIZONTAL_RULE) {
            hrToRemove = prevSibling;
          } else if (prevSibling.getType() === DocumentApp.ElementType.PARAGRAPH && prevSibling.getNumChildren() > 0) {
            const firstChild = prevSibling.getChild(0);
            if (firstChild.getType() === DocumentApp.ElementType.HORIZONTAL_RULE) {
               hrToRemove = prevSibling;
            }
          }
        }
      }
    }
    
    // 3. Execute conditional display: populate rows or cleanly tear down elements
    if (shouldDisplay) {
      if (headingParagraph) body.replaceText('_UcasHeading_', 'UCAS predicted grades');
      
      if (ucasTable && templateRow) {
        ucasTable.removeRow(templateRowIndex);
        ucasSubjects.forEach((subj, index) => {
          const newRow = templateRow.copy();
          newRow.replaceText('{{ucasSubjectName}}', subj.subjectName || '');
          newRow.replaceText('{{ucasGrade}}', subj.ucas || '');
          ucasTable.insertTableRow(templateRowIndex + index, newRow);
        });
      }
    } else {
      // Remove cleanly from document flow to eliminate unnecessary whitespace
      if (hrToRemove) hrToRemove.removeFromParent();
      if (headingParagraph) headingParagraph.removeFromParent();
      if (ucasTable) ucasTable.removeFromParent();
      body.replaceText('_UcasHeading_', ''); 
    }
  },
  
  /**
   * Populates dynamic rows inside the main subject assessment table using
   * token replacement on a cloned template row.
   * @private
   * @param {GoogleAppsScript.Document.Body} body Document body.
   * @param {Array<Object>} subjects Student subjects.
   */
  _populateSubjectTable: function(body, subjects) {
    const tables = body.getTables();
    if (tables.length === 0) return;
    
    let targetTable = null;
    let templateRow = null;
    let templateRowIndex = -1;
    
    // Search across all tables to identify the one containing our template placeholder
    for (let t = 0; t < tables.length; t++) {
      const table = tables[t];
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        if (row.getText().includes('{{subjectName}}')) {
          targetTable = table;
          templateRow = row.copy();
          templateRowIndex = r;
          table.removeRow(r);
          break;
        }
      }
      if (targetTable) break;
    }
    
    if (!targetTable || !templateRow) return;
    
    // Inject a populated row for every subject
    subjects.forEach((subj, index) => {
      const newRow = templateRow.copy();
      
      newRow.replaceText('{{subjectName}}', subj.subjectName || '');
      newRow.replaceText('{{teacher}}', subj.teacher || '');
      newRow.replaceText('{{stg}}', subj.stg || '');
      newRow.replaceText('{{crnt}}', subj.crnt || '');
      newRow.replaceText('{{ci1}}', subj.ci1 || '');
      newRow.replaceText('{{ci2}}', subj.ci2 || '');
      newRow.replaceText('{{ci3}}', subj.ci3 || '');
      newRow.replaceText('{{ci4}}', subj.ci4 || '');
      newRow.replaceText('{{nextSteps1}}', subj.nextSteps1 || '');
      newRow.replaceText('{{nextSteps2}}', subj.nextSteps2 || '');
      
      newRow.replaceText('{{subjAtt}}', subj.subjAtt || '');
      newRow.replaceText('{{subjLates}}', subj.subjLates || '');
      newRow.replaceText('{{ucas}}', subj.ucas || '');
      newRow.replaceText('{{prd}}', subj.prd || '');
      newRow.replaceText('{{eoy}}', subj.eoy || '');
      newRow.replaceText('{{classRank}}', subj.classRank || '');
      
      targetTable.insertTableRow(templateRowIndex + index, newRow);
    });
  }
};