/**
 * @fileoverview Global Configuration File.
 * Acts as the immutable single source of truth for the entire reporting system,
 * centralising folder identifiers, access-control lists, sheet naming conventions,
 * and report-specific configuration profiles.
 */

/**
 * Global configuration object for Typeless Reports.
 * Wrapped in Object.freeze to prevent accidental runtime mutations across script executions.
 * @type {Readonly<Object>}
 */
const CONFIG = Object.freeze({
  // ---------------------------------------------------------------------------
  // 1. Global Drive Settings
  // ---------------------------------------------------------------------------
  GLOBAL: {
    // The default parent Drive folder where standard report batch subfolders are generated.
    OUTPUT_FOLDER_ID: '18JTL77flcaOV7Us93W_hJnf3_veATeFp',
  },

  // ---------------------------------------------------------------------------
  // 2. Authorisation Controls
  // ---------------------------------------------------------------------------
  AUTH: {
    // Super users have full administrative access: setting up sheets, freezing/thawing,
    // and running any report across Key Stages.
    SUPER_USERS: [
      'jappleton@csg.school',
      'tnayagam@csg.school'
    ],
    // Role-specific access delegations to allow restricted staff members to run
    // their respective workflows without exposing global administrative tools.
    REPORT_SPECIFIC: { 
      UCAS: ['nbayley@csg.school'] 
    }
  },

  // ---------------------------------------------------------------------------
  // 3. Import Sheet Controls
  // ---------------------------------------------------------------------------
  IMPORT: {
    targetSheetName: 'import',
    backupSheetName: 'import-backup',
    // Rows 6 & 7 house the array/lookup formulas that dynamically pull MIS data.
    anchorRowStart: 6,
    anchorRowCount: 2,
    // Visual emoji indicator displaying whether the sheet is frozen (🥶) or live (🫠).
    statusCell: 'A1'
  },

  // ---------------------------------------------------------------------------
  // 4. Setup & Map Controls (Named Ranges on Control Panel)
  // ---------------------------------------------------------------------------
  SCOPE: {
    subjectDetailsRange: 'scopeSubjectDetails',
    yearGroup: 'scopeYearGroup',
    keyStage: 'scopeKeyStage',
    academicYear: 'scopeAcademicYear',
    collection: 'scopeCollection',
    targetSubjectNameRange: 'thisSubjectName',
    shortName: 'scopeShortname',
    until: 'scopeAttUntil',
    // Named ranges referencing dynamic header translation and mapping tables
    fieldMap: 'scopeFieldMap',
    translations: 'scopeTranslations'
  },

  // ---------------------------------------------------------------------------
  // 5. Fallback Field Mapper
  // ---------------------------------------------------------------------------
  // Serves as a failsafe dictionary in the event that the 'scopeFieldMap'
  // named range on the spreadsheet Control Panel is altered, renamed, or corrupt.
  FALLBACK_FIELD_MAP: {
    tut_adno: 'adno',
    tut_attTpAs: 'attendance tpas',
    tut_latesTpAs: 'lates tpas',
    tut_ucas_ref: '✎ ucas ref.',
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
    // KS5 Specific Additions
    subj_att: 'att %',
    subj_lates: 'lates',
    subj_eoy: 'eoy',
    subj_ucas: 'ucas',
    subj_ucas_ref: '✎ ucas ref.',
    subj_class_rank: 'rank'
  },

  // ---------------------------------------------------------------------------
  // 6. Report Profiles
  // ---------------------------------------------------------------------------
  // Each profile encapsulates the template Google Doc ID and any report-specific
  // destination overrides required by the document generation engine.
  REPORTS: {
    PROGRESS_REVIEW: {
      name: 'Progress Review',
      templateId: '1GiRm5ry4MCUWEMzHyAdRpqYGoeJcid3smTDaH7gIizE'
    },
    NEXT_STEPS_SUMMARY: {
      name: 'Next Steps Summary',
      templateId: '1no4SLsNv1N74s9a5l_2Lo4qp1v6ZG-Hnfbvh51puTmw'
    },
    UCAS_REFERENCE: {
      name: 'UCAS Reference Collection',
      templateId: '1bPZxAa-K7oR9hin9cNVP0x6Lfzr9fu0ULPytN6zvrO4',
      // The UCAS references must output to a dedicated destination folder
      // rather than the general batch reporting folder.
      outputFolderId: '13PHDtPVB97Yehk8znsccWmQEKpH8Aidr'
    },
    EOY_REPORT: {
      name: 'End of Year Report',
      templateId: '1t9q9tDKkwt0zvJ1ppocfR5AdKYo7DQU24o13JfG5Wqg'
    }
  }
});