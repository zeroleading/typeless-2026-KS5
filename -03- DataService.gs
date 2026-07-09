/** * DataService.gs 
 * Handles data extraction, in-memory aggregation, and translation of student records. 
 */
const DataService = {
  buildStudentDataPayload: function(reportConfig) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Fetch Global Batch Values
    const yearGroup = ss.getRangeByName(CONFIG.SCOPE.yearGroup)?.getValue() || '';
    const collection = ss.getRangeByName(CONFIG.SCOPE.collection)?.getValue() || '';
    const academicYear = ss.getRangeByName(CONFIG.SCOPE.academicYear)?.getValue() || '';
    const shortName = ss.getRangeByName(CONFIG.SCOPE.shortName)?.getValue() || '';
    const until = ss.getRangeByName(CONFIG.SCOPE.until)?.getValue() || '';
    
    // 2. Fetch Control Panel Maps & Dictionaries
    const fieldMap = this._getDynamicFieldMap(ss);
    const translations = this._getTranslationsDictionary(ss);
    
    // 3. Build base maps and attach data
    const studentMap = this._getMasterStudentList(ss);
    this._attachTutorData(ss, studentMap, fieldMap);
    
    // --- KS5 Further Maths Logic Setup ---
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
    // -------------------------------------
    
    // 4. Process Subject Sheets
    const allSheets = ss.getSheets();
    const subjectRegex = /^([A-Z][a-z]|EnL)$/;
    allSheets.forEach(sheet => {
      if (subjectRegex.test(sheet.getName())) {
        this._processSubjectSheet(ss, sheet, studentMap, fieldMap, translations, reportConfig, isYear12, fmStudents);
      }
    });
    
    // 5. Convert map to array and inject globals
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
   * Generates a plain-text preview of collated UCAS references for a specific student.
   * Isolates the logic from the UI Controller.
   */
  getUcasPreviewText: function(reportConfig, targetAdNo) {
    const payload = this.buildStudentDataPayload(reportConfig);
    // Support matching both raw and 6-digit padded admission numbers
    const student = payload.find(s => 
      String(s.adNo) === String(targetAdNo) || 
      String(s.adNo).padStart(6, '0') === String(targetAdNo).padStart(6, '0')
    );
    
    if (!student) return null;
    
    let combinedRefs = '';
    student.subjects.forEach(subj => {
      if (subj.ucasRef) {
        combinedRefs += `${subj.subjectName} (${subj.teacher}):\n${subj.ucasRef}\n\n`;
      }
    });
    
    return {
      name: student.name,
      previewText: combinedRefs.trim() || 'No references found for this student.'
    };
  },
  
  _getDynamicFieldMap: function(ss) {
    const map = { ...CONFIG.FALLBACK_FIELD_MAP };
    const range = ss.getRangeByName(CONFIG.SCOPE.fieldMap);
    if (!range) return map;
    
    const data = range.getValues();
    data.forEach(row => {
      const internalRef = String(row[0]).trim();
      const targetHeader = String(row[1]).trim();
      if (internalRef && targetHeader && !internalRef.includes('**')) {
        map[internalRef] = targetHeader;
      }
    });
    return map;
  },
  
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
  
  _translate: function(rawValue, category, translationsDict) {
    if (rawValue === '' || rawValue === undefined) return '';
    const safeValue = String(rawValue).trim().toUpperCase();
    if (translationsDict[category] && translationsDict[category][safeValue]) {
      return translationsDict[category][safeValue];
    }
    return String(rawValue);
  },
  
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
        
        // Dynamically assign boolean based on earlyApp column
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
  
  _attachTutorData: function(ss, studentMap, fieldMap) {
    const range = ss.getRangeByName('tutorAssessment');
    if (!range) return;
    
    // Using getDisplayValues() ensures percentages/numbers are pulled exactly as formatted strings
    const data = range.getDisplayValues();
    if (data.length < 3) return;
    
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const adNoIdx = headers.indexOf((fieldMap['tut_adno'] || '').toLowerCase());
    const attTpAsIdx = headers.indexOf((fieldMap['tut_attTpAs'] || '').toLowerCase());
    const latesTpAsIdx = headers.indexOf((fieldMap['tut_latesTpAs'] || '').toLowerCase());
    
    if (adNoIdx === -1) return;
    
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const rawAdNo = row[adNoIdx];
      if (!rawAdNo) continue;
      
      const adNo = String(rawAdNo).trim();
      if (studentMap[adNo]) {
        studentMap[adNo].tutorInfo = {
          attTpAs: attTpAsIdx > -1 ? row[attTpAsIdx] : '',
          latesTpAs: latesTpAsIdx > -1 ? row[latesTpAsIdx] : ''
        };
      }
    }
  },
  
  _processSubjectSheet: function(ss, sheet, studentMap, fieldMap, translations, reportConfig, isYear12, fmStudents) {
    const sheetName = sheet.getName();
    
    // Year 12 Further Maths Exclusion
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
    
    // KS5 Specific Headers
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
        // Further Maths Rename Logic
        let finalSubjectName = fullSubjectName;
        if (isYear12 && sheetName === 'Ma' && fmStudents.has(adNo)) {
          // Skip the rename for the EOY report
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
        let missingElements = [];
        
        if (reportConfig.name === CONFIG.REPORTS.EOY_REPORT.name) {
          if (rawEoy === '') missingElements.push('EOY');
        } else if (reportConfig.name === CONFIG.REPORTS.UCAS_REFERENCE.name) {
          if (rawUcas === '') missingElements.push('UCAS Grade');
          if (rawClassRank === '') missingElements.push('Class Rank');
          if (rawUcasRef === '') missingElements.push('UCAS Ref');
        } else {
          if (rawCrnt === '') missingElements.push('CRNT');
          if (rawCi1 === '') missingElements.push('CI1');
          if (rawCi2 === '') missingElements.push('CI2');
          if (rawCi3 === '') missingElements.push('CI3');
          if (rawCi4 === '') missingElements.push('CI4');
          const rawNs1 = ns1Idx > -1 ? String(row[ns1Idx]).trim() : '';
          const rawNs2 = ns2Idx > -1 ? String(row[ns2Idx]).trim() : '';
          if (rawNs1 === '' && rawNs2 === '') missingElements.push('Next Steps');
        }

        if (missingElements.length > 0) {
          studentMap[adNo].auditIssues.push(`${finalSubjectName} (${missingElements.join(', ')})`);
        }
        // -------------------
        
        // Direct conversion for KS5 requirements with 'X' substitution
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
  }
};