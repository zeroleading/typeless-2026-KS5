/** * Config.gs 
 * Global Configuration File 
 * Acts as the single source of truth for the entire reporting system. 
 */
const CONFIG = Object.freeze({
  // 1. Global Settings
  GLOBAL: {
    OUTPUT_FOLDER_ID: '18JTL77flcaOV7Us93W_hJnf3_veATeFp',
  },
  // 2. Authorisation Controls
  AUTH: {
    SUPER_USERS: [
      'jappleton@csg.school',
      'tnayagam@csg.school'
    ],
    REPORT_SPECIFIC: {}
  },
  // 3. Import Sheet Controls
  IMPORT: {
    targetSheetName: 'import',
    backupSheetName: 'import-backup',
    anchorRowStart: 6,
    anchorRowCount: 2,
    statusCell: 'A1'
  },
  // 4. Setup & Map Controls
  SCOPE: {
    subjectDetailsRange: 'scopeSubjectDetails',
    yearGroup: 'scopeYearGroup',
    keyStage: 'scopeKeyStage',
    academicYear: 'scopeAcademicYear',
    collection: 'scopeCollection',
    targetSubjectNameRange: 'thisSubjectName',
    shortName: 'scopeShortname',
    // New dynamic tables on the Control Panel
    fieldMap: 'scopeFieldMap',
    translations: 'scopeTranslations'
  },
  // 5. Fallback Field Mapper 
  // Used only if the scopeFieldMap named range is missing or broken.
  FALLBACK_FIELD_MAP: {
    tut_adno: 'adno',
    tut_attTpAs: 'attendance tpas',
    tut_latesTpAs: 'lates tpas',
    subj_adno: 'adno',
    subj_teacher: 'teacher',
    subj_stg: 'stg',
    subj_crnt: 'crnt',
    subj_ci1: 'ci1',
    subj_ci2: 'ci2',
    subj_ci3: 'ci3',
    subj_ci4: 'ci4',
    subj_ns1: '≣ nextsteps1',
    subj_ns2: '≣ nextsteps2',
    // --- KS5 Specific Additions ---
    subj_att: 'att %',
    subj_lates: 'lates',
    subj_eoy: 'eoy',
    subj_ucas: 'ucas',
    subj_ucas_ref: '✎ ucas ref.',
    subj_class_rank: 'rank'
  },
  // 6. Report Profiles
  REPORTS: {
    PROGRESS_REVIEW: {
      name: 'Progress Review',
      templateId: '1mqVkM7VBjok1Hpe9KSCxpnZkkyRIRDF70dZt2zJrVUo'
    },
    NEXT_STEPS_SUMMARY: {
      name: 'Next Steps Summary',
      templateId: '1Z6O8k6C67vDBp3heHZ-9reT5Glfnc8lLZ-dqpUivf74'
    },
    // --- KS5 Specific Reports ---
    UCAS_REFERENCE: {
      name: 'UCAS Reference Collection',
      templateId: 'PLACEHOLDER_UCAS_ID'
    },
    EOY_REPORT: {
      name: 'End of Year Report',
      templateId: '1t9q9tDKkwt0zvJ1ppocfR5AdKYo7DQU24o13JfG5Wqg'
    }
  }
});