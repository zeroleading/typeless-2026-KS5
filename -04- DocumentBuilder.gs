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

  generateChunk: function(reportConfig, chunkPayload, folderId, auditMode = 'ignore') {
    const templateFile = DriveApp.getFileById(reportConfig.templateId);
    const batchFolder = DriveApp.getFolderById(folderId);
    
    chunkPayload.forEach((student) => {
      // 1. Filter the subjects if the user selected to drop incomplete ones
      let validSubjects = student.subjects || [];
      if (auditMode === 'drop') {
        validSubjects = validSubjects.filter(subj => this._isSubjectComplete(subj, reportConfig.name));
      }

      // 2. Only generate a document if there is at least one valid subject left
      if (validSubjects.length > 0) {
        student.subjects = validSubjects; // Overwrite payload to exclude dropped subjects
        this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
      }
    });
  },

  // --- EXISTING METHODS (Used by UCAS Sidebar) ---
  
  generateBatch: function(reportConfig, studentPayload, auditMode = 'ignore') {
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
      // 1. Filter the subjects if the user selected to drop incomplete ones
      let validSubjects = student.subjects || [];
      if (auditMode === 'drop') {
        validSubjects = validSubjects.filter(subj => this._isSubjectComplete(subj, reportConfig.name));
      }

      // 2. Only generate a document if there is at least one valid subject left
      if (validSubjects.length > 0) {
        student.subjects = validSubjects; // Overwrite payload to exclude dropped subjects
        ss.toast(`Merging document ${index + 1} of ${totalStudents}...\n(${student.name})`, 'Progress Tracker', 10);
        this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
      }
    });
    
    return batchFolder.getId();
  },

  /**
   * Helper logic to determine if a subject meets the minimum data requirements
   */
  _isSubjectComplete: function(subj, reportName) {
    if (reportName === CONFIG.REPORTS.EOY_REPORT.name) {
      return subj.eoy !== '';
    } else if (reportName === CONFIG.REPORTS.UCAS_REFERENCE.name) {
      return subj.ucas !== '' && subj.classRank !== '' && subj.ucasRef !== '';
    } else {
      // Progress Review and Next Steps Summaries
      return subj.crnt !== '' &&
             subj.ci1 !== '' &&
             subj.ci2 !== '' &&
             subj.ci3 !== '' &&
             subj.ci4 !== '' &&
             (subj.nextSteps1 !== '' || subj.nextSteps2 !== ''); // Fails if BOTH are blank
    }
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
    
    if (reportName === CONFIG.REPORTS.UCAS_REFERENCE.name) {
      this._injectUcasReferences(body, student.subjects);
    }
    
    this._processConditionalUcasTable(body, student, reportName);
    this._populateSubjectTable(body, student.subjects);
    newDoc.saveAndClose();
  },

  /**
   * Conditionally injects or removes the UCAS grades table.
   * Only triggers for Year 12s during 'Progress Review B'.
   */
  _processConditionalUcasTable: function(body, student, reportName) {
    if (reportName !== CONFIG.REPORTS.PROGRESS_REVIEW.name) return;
    
    const isYear12 = String(student.yearGroup).includes('12');
    const isPRB = String(student.collection).trim() === 'Progress Review B';
    
    // Filter to only subjects that have a non-empty UCAS grade
    const ucasSubjects = student.subjects.filter(subj => subj.ucas && String(subj.ucas).trim() !== '');
    
    // Condition is only met if the student is Y12, it's PR B, AND they have at least one valid grade
    const shouldDisplay = isYear12 && isPRB && ucasSubjects.length > 0;
    
    // 1. Locate the UCAS Table
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
    
    // 2. Locate the Heading and preceding Horizontal Rule
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
            // Google Docs sometimes embeds the HR inside an empty wrapper paragraph
            const firstChild = prevSibling.getChild(0);
            if (firstChild.getType() === DocumentApp.ElementType.HORIZONTAL_RULE) {
               hrToRemove = prevSibling;
            }
          }
        }
      }
    }
    
    // 3. Execute Condition
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
      // Remove everything cleanly from the document
      if (hrToRemove) hrToRemove.removeFromParent();
      if (headingParagraph) headingParagraph.removeFromParent();
      if (ucasTable) ucasTable.removeFromParent();
      
      // Fallback cleanup just in case the paragraph removal failed
      body.replaceText('_UcasHeading_', ''); 
    }
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