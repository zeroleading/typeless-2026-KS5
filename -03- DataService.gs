/**
 * @fileoverview DataService.gs
 * Handles data extraction, in-memory aggregation, dictionary lookups,
 * and translation of student assessment and reference records across sheets.
 */

const DataService = {
  /**
   * Compiles the comprehensive student data payload required across reports.
   * Pulls cohort parameters, builds base student maps, and parses each subject sheet.
   * @param {Object} reportConfig Configuration definition for the active report.
   * @return {Array<Object>} List of aggregated student record objects.
   */
  buildStudentDataPayload: function(reportConfig) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Fetch Global Batch Values from the Control Panel named ranges
    const yearGroup = ss.getRangeByName(CONFIG.SCOPE.yearGroup)?.getValue() || '';
    const collection = ss.getRangeByName(CONFIG.SCOPE.collection)?.getValue() || '';
    const academicYear = ss.getRangeByName(CONFIG.SCOPE.academicYear)?.getValue() || '';
    const shortName = ss.getRangeByName(CONFIG.SCOPE.shortName)?.getValue() || '';
    const until = ss.getRangeByName(CONFIG.SCOPE.until)?.getValue() || '';
    
    // 2. Fetch Control Panel Maps & Dictionaries dynamically
    const fieldMap = this._getDynamicFieldMap(ss);
    const translations = this._getTranslationsDictionary(ss);
    
    // 3. Build base maps from the master directory and attach tutor attendance/references
    const studentMap = this._getMasterStudentList(ss);
    this._attachTutorData(ss, studentMap, fieldMap);
    
    // -------------------------------------------------------------------------
    // KS5 Further Maths Logic Setup:
    // In Year 12, students studying Further Maths take single Maths lessons with
    // a specialised programme. We detect these students by inspecting the 'Fm' sheet.
    // -------------------------------------------------------------------------
    const isYear12 = String(yearGroup).includes('12');
    const fmStudents = new Set();
    
    if (isYear12) {
      const fmSheet = ss.getSheetByName('Fm');
      if (fmSheet) {
        const fmRange = ss.getRangeByName('Fm!thisSubjectAssessment');
        if (fmRange) {
          const data = fmRange.getValues();
          if (data.length >= 3) {
            const headers = data[0].map(h => String(h).toLowerCase().trim());
            const adNoColIdx = headers.indexOf((fieldMap['subj_adno'] || '').toLowerCase());
            if (adNoColIdx > -1) {
              for (let i = 2; i < data.length; i++) {
                const rawAdNo = data[i][adNoColIdx];
                if (rawAdNo) fmStudents.add(String(rawAdNo).trim());
              }
            }
          }
        }
      }
    }
    
    // 4. Process all individual Subject Sheets matching standard naming patterns
    // Subject sheets follow 2-letter codes (e.g., 'Bi', 'Ch', 'Ma') or 'EnL'.
    const allSheets = ss.getSheets();
    const subjectRegex = /^([A-Z][a-z]|EnL)$/;
    allSheets.forEach(sheet => {
      if (subjectRegex.test(sheet.getName())) {
        this._processSubjectSheet(ss, sheet, studentMap, fieldMap, translations, reportConfig, isYear12, fmStudents);
      }
    });
    
    // 5. Convert map to array and inject global batch metadata into each student record
    return Object.values(studentMap).map(student => ({
      ...student,
      yearGroup: yearGroup,
      collection: collection,
      academicYear: academicYear,
      shortName: shortName,
      until: until
    }));
  },

  /**
   * Generates structured preview data for a specific student's UCAS submission.
   * Isolates Section 2 (Tutor Extenuating Circumstances), Section 3 (Subject Suitability),
   * and Section 4 (Predicted Grades) to feed the interactive sidebar cards.
   * @param {Object} reportConfig The UCAS report configuration object.
   * @param {string|number} targetAdNo Student admission number.
   * @return {Object|null} Structured preview payload or null if the student is not found.
   */
  getUcasPreviewText: function(reportConfig, targetAdNo) {
    const payload = this.buildStudentDataPayload(reportConfig);
    // Support matching both unpadded numbers and standard 6-digit padded admission numbers
    const student = payload.find(s => 
      String(s.adNo) === String(targetAdNo) || 
      String(s.adNo).padStart(6, '0') === String(targetAdNo).padStart(6, '0')
    );
    
    if (!student) return null;
    
    // Section 2: Tutor extenuating circumstances / contextual commentary
    const tutorRef = (student.tutorInfo && student.tutorInfo.ucasRef) 
      ? String(student.tutorInfo.ucasRef).trim() 
      : '';

    // Section 3: Subject suitability narrative blocks
    const subjectRefs = this.formatUcasSubjectReferences(student.subjects);

    // Section 4: Plain-text predicted grades summary
    const predictions = this.formatUcasPredictions(student.subjects);
    
    return {
      name: student.name,
      adNo: student.adNo,
      reg: student.reg,
      tutorRef: tutorRef,
      subjectRefs: subjectRefs,
      predictions: predictions
    };
  },

  /**
   * Compiles the full cohort's UCAS references in a single operation.
   * Transmits a structured map to the client sidebar for zero-latency
   * in-memory previewing and real-time completeness filtering.
   * @param {Object} reportConfig The UCAS report configuration object.
   * @return {Object<string, Object>} Student dictionary keyed by admission number.
   */
  getCohortUcasData: function(reportConfig) {
    const payload = this.buildStudentDataPayload(reportConfig);
    const cohortCache = {};

    payload.forEach(student => {
      // Extract tutor contextual comments
      const tutorRef = (student.tutorInfo && student.tutorInfo.ucasRef) 
        ? String(student.tutorInfo.ucasRef).trim() 
        : '';

      // Format subject suitability references and predictions
      const subjectRefs = this.formatUcasSubjectReferences(student.subjects);
      const predictions = this.formatUcasPredictions(student.subjects);

      // Evaluate application completeness to drive the wildcard audit filter
      const missingElements = [];
      if (!tutorRef) missingElements.push('Tutor Ref');
      if (!subjectRefs) missingElements.push('Subject Refs');
      if (!predictions) missingElements.push('Predictions');
      if (student.auditIssues && student.auditIssues.length > 0) {
        missingElements.push(...student.auditIssues);
      }

      const isComplete = missingElements.length === 0;

      cohortCache[String(student.adNo)] = {
        name: student.name,
        adNo: student.adNo,
        reg: student.reg,
        earlyApp: student.earlyApp,
        tutorRef: tutorRef,
        subjectRefs: subjectRefs,
        predictions: predictions,
        isComplete: isComplete,
        missingSummary: missingElements.join(', ')
      };
    });

    return cohortCache;
  },

  /**
   * Formats subject references into the required UCAS narrative block.
   * Template specification:
   *   ${subj.subjectName}
   *    12 Exam: ${subj.eoy}, Rank: ${subj.classRank},
   *   ${subj.ucasRef}
   * @param {Array<Object>} subjects List of subject objects for a student.
   * @return {string} Formatted subject narrative block.
   */
  formatUcasSubjectReferences: function(subjects) {
    let output = '';
    (subjects || []).forEach(subj => {
      // Exclude subjects where no reference text has been contributed by the teacher
      if (subj.ucasRef && String(subj.ucasRef).trim() !== '') {
        const eoy = subj.eoy ? String(subj.eoy).trim() : '-';
        const rank = subj.classRank ? String(subj.classRank).trim() : '-';
        output += `${subj.subjectName}\n 12 Exam: ${eoy}, Rank: ${rank},\n${String(subj.ucasRef).trim()}\n\n`;
      }
    });
    return output.trim();
  },

  /**
   * Formats predicted grades into a concise plain-text list for Section 4.
   * Template specification: ${subj.subjectName}: ${subj.ucas}
   * @param {Array<Object>} subjects List of subject objects for a student.
   * @return {string} Plain-text summary of predicted grades.
   */
  formatUcasPredictions: function(subjects) {
    let output = '';
    (subjects || []).forEach(subj => {
      // Include only subjects where a valid prediction has been assigned
      if (subj.ucas && String(subj.ucas).trim() !== '') {
        output += `${subj.subjectName}: ${String(subj.ucas).trim()}\n`;
      }
    });
    return output.trim();
  },
  
  /**
   * Resolves dynamic header mappings from the 'scopeFieldMap' named range.
   * If the range is unconfigured, seamlessly falls back to FALLBACK_FIELD_MAP.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Active spreadsheet.
   * @return {Object} Map of internal field codes to sheet column headers.
   */
  _getDynamicFieldMap: function(ss) {
    const map = { ...CONFIG.FALLBACK_FIELD_MAP };
    const range = ss.getRangeByName(CONFIG.SCOPE.fieldMap);
    if (!range) return map;
    
    const data = range.getValues();
    data.forEach(row => {
      const internalRef = String(row[0]).trim();
      const targetHeader = String(row[1]).trim();
      // Skip commented-out entries prefixed with asterisks
      if (internalRef && targetHeader && !internalRef.includes('**')) {
        map[internalRef] = targetHeader;
      }
    });
    return map;
  },
  
  /**
   * Resolves category translation dictionaries (e.g. Attitude to Learning codes).
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Active spreadsheet.
   * @return {Object} Nested dictionary: { CATEGORY: { CODE: translation } }
   */
  _getTranslationsDictionary: function(ss) {
    const dict = {};
    const range = ss.getRangeByName(CONFIG.SCOPE.translations);
    if (!range) return dict;
    
    const data = range.getValues();
    data.forEach(row => {
      const category = String(row[0]).trim().toUpperCase();
      const code = String(row[1]).trim().toUpperCase();
      const translation = String(row[2]).trim();
      if (category && code && !category.includes('**')) {
        if (!dict[category]) dict[category] = {};
        dict[category][code] = translation;
      }
    });
    return dict;
  },
  
  /**
   * Helper translating raw codes into friendly report text using the dictionary.
   * @private
   */
  _translate: function(rawValue, category, translationsDict) {
    if (rawValue === '' || rawValue === undefined) return '';
    const safeValue = String(rawValue).trim().toUpperCase();
    if (translationsDict[category] && translationsDict[category][safeValue]) {
      return translationsDict[category][safeValue];
    }
    return String(rawValue);
  },
  
  /**
   * Extracts the master student cohort from the 'simpleStudentData' named range.
   * Detects early applicant status to facilitate filtering in the sidebar.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Active spreadsheet.
   * @return {Object} Map of admission number to base student descriptor.
   */
  _getMasterStudentList: function(ss) {
    const studentMap = {};
    const range = ss.getRangeByName('simpleStudentData');
    if (!range) return studentMap;
    
    const data = range.getValues();
    const headers = data[0] ? data[0].map(h => String(h).toLowerCase().trim()) : [];
    const earlyAppIdx = headers.indexOf('earlyapp');
    
    data.forEach(row => {
      const fullName = row[0];
      const rawAdNo = row[2];
      const reg = row[3];
      const tutor = row[5];
      
      if (rawAdNo && String(rawAdNo).toLowerCase() !== 'adno') {
        const adNo = String(rawAdNo).trim();
        
        // Dynamically assign early applicant status based on the earlyApp column
        let isEarly = false;
        if (earlyAppIdx > -1) {
          const val = String(row[earlyAppIdx]).toLowerCase().trim();
          isEarly = (val === 'true' || val === 'yes' || val === 'y');
        }
        
        studentMap[adNo] = {
          adNo: adNo,
          name: fullName,
          reg: reg,
          tutor: tutor,
          earlyApp: isEarly,
          tutorInfo: {},
          subjects: [],
          auditIssues: []
        };
      }
    });
    return studentMap;
  },
  
  /**
   * Attaches form tutor assessment data, attendance percentages, and contextual
   * UCAS reference material from the 'tutorAssessment' named range.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Active spreadsheet.
   * @param {Object} studentMap Student master map indexed by admission number.
   * @param {Object} fieldMap Header mapping dictionary.
   */
  _attachTutorData: function(ss, studentMap, fieldMap) {
    const range = ss.getRangeByName('tutorAssessment');
    if (!range) return;
    
    // Using getDisplayValues() preserves exact percentage/number formatting as rendered in Sheets
    const data = range.getDisplayValues();
    if (data.length < 3) return;
    
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const adNoIdx = headers.indexOf((fieldMap['tut_adno'] || '').toLowerCase());
    const attTpAsIdx = headers.indexOf((fieldMap['tut_attTpAs'] || '').toLowerCase());
    const latesTpAsIdx = headers.indexOf((fieldMap['tut_latesTpAs'] || '').toLowerCase());
    const tutUcasRefIdx = headers.indexOf((fieldMap['tut_ucas_ref'] || '').toLowerCase());
    
    if (adNoIdx === -1) return;
    
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const rawAdNo = row[adNoIdx];
      if (!rawAdNo) continue;
      
      const adNo = String(rawAdNo).trim();
      if (studentMap[adNo]) {
        studentMap[adNo].tutorInfo = {
          attTpAs: attTpAsIdx > -1 ? row[attTpAsIdx] : '',
          latesTpAs: latesTpAsIdx > -1 ? row[latesTpAsIdx] : '',
          ucasRef: tutUcasRefIdx > -1 ? row[tutUcasRefIdx] : ''
        };
      }
    }
  },
  
  /**
   * Reads an individual subject sheet, performs pre-flight audit validation,
   * handles Further Maths naming adjustments, and stores clean subject records.
   * @private
   */
  _processSubjectSheet: function(ss, sheet, studentMap, fieldMap, translations, reportConfig, isYear12, fmStudents) {
    const sheetName = sheet.getName();
    
    // Year 12 Further Maths Exclusion: FM students are assessed within the main Maths sheet
    if (isYear12 && sheetName === 'Fm') return;
    
    const nameRangeStr = `${sheetName}!${CONFIG.SCOPE.targetSubjectNameRange}`;
    const nameRange = ss.getRangeByName(nameRangeStr);
    const fullSubjectName = nameRange ? String(nameRange.getValue()).trim() : sheetName;
    
    const rangeName = `${sheetName}!thisSubjectAssessment`;
    const range = ss.getRangeByName(rangeName);
    if (!range) return;
    
    const data = range.getValues();
    if (data.length < 3) return;
    
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const adNoColIdx = headers.indexOf((fieldMap['subj_adno'] || '').toLowerCase());
    const teacherIdx = headers.indexOf((fieldMap['subj_teacher'] || '').toLowerCase());
    const stgIdx = headers.indexOf((fieldMap['subj_stg'] || '').toLowerCase());
    const crntIdx = headers.indexOf((fieldMap['subj_crnt'] || '').toLowerCase());
    const ci1Idx = headers.indexOf((fieldMap['subj_ci1'] || '').toLowerCase());
    const ci2Idx = headers.indexOf((fieldMap['subj_ci2'] || '').toLowerCase());
    const ci3Idx = headers.indexOf((fieldMap['subj_ci3'] || '').toLowerCase());
    const ci4Idx = headers.indexOf((fieldMap['subj_ci4'] || '').toLowerCase());
    const ns1Idx = headers.indexOf((fieldMap['subj_ns1'] || '').toLowerCase());
    const ns2Idx = headers.indexOf((fieldMap['subj_ns2'] || '').toLowerCase());
    
    // KS5 Specific Assessment Headers
    const subjAttIdx = headers.indexOf((fieldMap['subj_att'] || '').toLowerCase());
    const subjLatesIdx = headers.indexOf((fieldMap['subj_lates'] || '').toLowerCase());
    const ucasIdx = headers.indexOf((fieldMap['subj_ucas'] || '').toLowerCase());
    const prdIdx = headers.indexOf((fieldMap['subj_prd'] || '').toLowerCase());
    const eoyIdx = headers.indexOf((fieldMap['subj_eoy'] || '').toLowerCase());
    const ucasRefIdx = headers.indexOf((fieldMap['subj_ucas_ref'] || '').toLowerCase());
    const classRankIdx = headers.indexOf((fieldMap['subj_class_rank'] || '').toLowerCase());
    
    if (adNoColIdx === -1) return;
    
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const rawAdNo = row[adNoColIdx];
      if (!rawAdNo) continue;
      
      const adNo = String(rawAdNo).trim();
      if (studentMap[adNo]) {
        // Further Maths Rename Logic: Distinguish single Maths from standard cohorts
        let finalSubjectName = fullSubjectName;
        if (isYear12 && sheetName === 'Ma' && fmStudents.has(adNo)) {
          // Keep canonical name for End of Year reporting, rename for Progress Reviews
          if (reportConfig.name !== CONFIG.REPORTS.EOY_REPORT.name) {
            finalSubjectName = 'Mathematics (for Further Maths)';
          }
        }
        
        const rawStg = stgIdx > -1 ? row[stgIdx] : '';
        const rawCrnt = crntIdx > -1 ? row[crntIdx] : '';
        const rawCi1 = ci1Idx > -1 ? row[ci1Idx] : '';
        const rawCi2 = ci2Idx > -1 ? row[ci2Idx] : '';
        const rawCi3 = ci3Idx > -1 ? row[ci3Idx] : '';
        const rawCi4 = ci4Idx > -1 ? row[ci4Idx] : '';
        const rawUcas = ucasIdx > -1 ? row[ucasIdx] : '';
        const rawEoy = eoyIdx > -1 ? row[eoyIdx] : '';
        const rawUcasRef = ucasRefIdx > -1 ? row[ucasRefIdx] : '';
        const rawClassRank = classRankIdx > -1 ? row[classRankIdx] : '';
        
        // --- AUDIT CHECK ---
        // Verify that essential indicators required by each specific report are populated
        let missingElements = [];
        
        if (reportConfig.name === CONFIG.REPORTS.EOY_REPORT.name) {
          if (rawEoy === '') missingElements.push('EOY');
        } else if (reportConfig.name === CONFIG.REPORTS.UCAS_REFERENCE.name) {
          if (rawUcas === '') missingElements.push('UCAS Grade');
          if (rawClassRank === '') missingElements.push('Class Rank');
          if (rawUcasRef === '') missingElements.push('UCAS Ref');
        } else {
          // Standard Progress Review audits
          if (rawCrnt === '') missingElements.push('CRNT');
          if (rawCi1 === '') missingElements.push('CI1');
          if (rawCi2 === '') missingElements.push('CI2');
          if (rawCi3 === '') missingElements.push('CI3');
          if (rawCi4 === '') missingElements.push('CI4');
          const rawNs1 = ns1Idx > -1 ? String(row[ns1Idx]).trim() : '';
          const rawNs2 = ns2Idx > -1 ? String(row[ns2Idx]).trim() : '';
          // Audit fails only if BOTH next step columns are unselected
          if (rawNs1 === '' && rawNs2 === '') missingElements.push('Next Steps');
        }

        if (missingElements.length > 0) {
          studentMap[adNo].auditIssues.push(`${finalSubjectName} (${missingElements.join(', ')})`);
        }
        // -------------------
        
        // Normalise grades, converting internal 'X' placeholders into 'Pending'
        const formatGrade = (grade) => {
          if (!grade) return '';
          const g = String(grade).trim().toUpperCase();
          return g === 'X' ? 'Pending' : g;
        };
        
        const safeStg = formatGrade(rawStg);
        const safeCrnt = formatGrade(rawCrnt);
        const safeUcas = formatGrade(rawUcas);
        const safeEoy = formatGrade(rawEoy);
        const rawPrd = prdIdx > -1 ? row[prdIdx] : '';
        const safePrd = formatGrade(rawPrd);
        
        const subjectData = {
          subjectName: finalSubjectName,
          teacher: teacherIdx > -1 ? row[teacherIdx] : '',
          stg: safeStg,
          crnt: safeCrnt,
          ci1: this._translate(rawCi1, 'CI', translations),
          ci2: this._translate(rawCi2, 'CI', translations),
          ci3: this._translate(rawCi3, 'CI', translations),
          ci4: this._translate(rawCi4, 'CI', translations),
          nextSteps1: ns1Idx > -1 ? row[ns1Idx] : '',
          nextSteps2: ns2Idx > -1 ? row[ns2Idx] : '',
          // KS5 Additions
          subjAtt: subjAttIdx > -1 ? row[subjAttIdx] : '',
          subjLates: subjLatesIdx > -1 ? row[subjLatesIdx] : '',
          ucas: safeUcas,
          prd: safePrd,
          eoy: safeEoy,
          ucasRef: rawUcasRef,
          classRank: rawClassRank
        };
        studentMap[adNo].subjects.push(subjectData);
      }
    }
  },

  /**
   * Diagnostic test to identify exactly why tutor data is failing to map.
   * Isolates the 'tutorAssessment' range and analyses its headers.
   * @return {string} Human-readable diagnostic output.
   */
  testTutorDataMapping: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fieldMap = this._getDynamicFieldMap(ss);
    
    const targetHeader = (fieldMap['tut_attTpAs'] || '').toLowerCase();
    const range = ss.getRangeByName('tutorAssessment');
    
    if (!range) return "FAIL: The named range 'tutorAssessment' could not be found in this spreadsheet.";
    
    const data = range.getDisplayValues();
    if (data.length < 3) return `FAIL: The 'tutorAssessment' range only has ${data.length} rows. It requires at least 3 (Header, Sub-Header, Data).`;
    
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const attTpAsIdx = headers.indexOf(targetHeader);
    const adNoIdx = headers.indexOf((fieldMap['tut_adno'] || '').toLowerCase());
    
    let debugMsg = `🔍 TUTOR DATA DIAGNOSTIC\n\n`;
    debugMsg += `Target Header Expected: '${targetHeader}'\n`;
    debugMsg += `Header Found at Index: ${attTpAsIdx}\n`;
    debugMsg += `AdNo Found at Index: ${adNoIdx}\n\n`;
    
    if (attTpAsIdx === -1) {
      debugMsg += `❌ ERROR: Could not find any column matching '${targetHeader}'.\n`;
      debugMsg += `Here are the headers the script is actually seeing in row 1:\n[ ${headers.join(', ')} ]\n\n`;
      debugMsg += `Check for typos, trailing spaces, or verify the named range includes the header row.`;
      return debugMsg;
    }
    
    // Grab the first valid student as a sample
    let sampleData = "No valid student data rows found.";
    for (let i = 2; i < data.length; i++) {
      if (data[i][adNoIdx]) {
        sampleData = `Sample AdNo: ${data[i][adNoIdx]}\nSample Attendance Value: '${data[i][attTpAsIdx]}'`;
        break;
      }
    }
    
    debugMsg += `✅ SUCCESS!\n${sampleData}`;
    return debugMsg;
  }
};

/**
 * Diagnostic utility function to test tutor data mapping directly from the Apps Script editor.
 */
function RUN_TEST_TUTOR_DATA() {
  const result = DataService.testTutorDataMapping();
  console.log(result);
  try {
    SpreadsheetApp.getUi().alert("Tutor Data Debugger", result, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // Fails gracefully when triggered outside the active spreadsheet context
  }
}