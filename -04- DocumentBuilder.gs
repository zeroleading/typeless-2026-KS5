/** * DocumentBuilder.gs 
 * Handles the generation of Google Docs from templates. 
 */
const DocumentBuilder = {
  
  // --- NEW CHUNKING ENGINE METHODS ---
  
  createBatchFolder: function(reportConfig, sampleStudent) {
    const outputFolder = DriveApp.getFolderById(CONFIG.GLOBAL.OUTPUT_FOLDER_ID);
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "yyyy-MM-dd");
    
    const academicYear = sampleStudent?.academicYear || '';
    const collection = sampleStudent?.collection || '';
    const yearGroup = sampleStudent?.yearGroup || '';
    
    let folderName = `${academicYear} ${collection} ${yearGroup} ${dateStr}`.trim();
    if (reportConfig.name === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) folderName += " next-steps";
    if (reportConfig.name === CONFIG.REPORTS.EOY_REPORT.name) folderName += " EOY";
    
    const batchFolder = outputFolder.createFolder(folderName);
    return batchFolder.getId();
  },

  generateChunk: function(reportConfig, chunkPayload, folderId) {
    const templateFile = DriveApp.getFileById(reportConfig.templateId);
    const batchFolder = DriveApp.getFolderById(folderId);
    
    chunkPayload.forEach((student) => {
      if (student.subjects && student.subjects.length > 0) {
        this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
      }
    });
  },

  // --- EXISTING METHODS (Used by UCAS Sidebar) ---
  
  generateBatch: function(reportConfig, studentPayload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateFile = DriveApp.getFileById(reportConfig.templateId);
    const outputFolder = DriveApp.getFolderById(CONFIG.GLOBAL.OUTPUT_FOLDER_ID);
    
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
    
    studentPayload.forEach((student, index) => {
      if (student.subjects && student.subjects.length > 0) {
        ss.toast(`Merging document ${index + 1} of ${totalStudents}...\n(${student.name})`, 'Progress Tracker', 10);
        this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
      }
    });
    
    return batchFolder.getId();
  },
  
  _buildSingleDocument: function(student, templateFile, destinationFolder, reportName) {
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
    
    const replaceGlobalsInSection = (section) => {
      if (!section) return;
      section.replaceText('_Name_', student.name || '');
      section.replaceText('_Reg_', student.reg || '');
      section.replaceText('_AdNo_', paddedAdNo);
      section.replaceText('_Tutor_', student.tutor || '');
      section.replaceText('_Date_', Utilities.formatDate(new Date(), "Europe/London", "dd/MM/yyyy"));
      section.replaceText('_YearGroup_', student.yearGroup || '');
      section.replaceText('_Collection_', student.collection || '');
      
      if (student.tutorInfo) {
        section.replaceText('_AttTpAs_', student.tutorInfo.attTpAs || '-');
        section.replaceText('_LatesTpAs_', student.tutorInfo.latesTpAs || '0');
      }
    };
    
    replaceGlobalsInSection(body);
    replaceGlobalsInSection(header);
    replaceGlobalsInSection(footer);
    
    if (reportName === CONFIG.REPORTS.UCAS_REFERENCE.name) {
      this._injectUcasReferences(body, student.subjects);
    }
    
    this._populateSubjectTable(body, student.subjects);
    newDoc.saveAndClose();
  },
  
  _injectUcasReferences: function(body, subjects) {
    let combinedRefs = '';
    subjects.forEach(subj => {
      if (subj.ucasRef) {
        combinedRefs += `${subj.subjectName} (${subj.teacher}):\n${subj.ucasRef}\n\n`;
      }
    });
    body.replaceText('_Collected References_', combinedRefs.trim());
  },
  
  _populateSubjectTable: function(body, subjects) {
    const tables = body.getTables();
    if (tables.length === 0) return;
    
    let targetTable = null;
    let templateRow = null;
    let templateRowIndex = -1;
    
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