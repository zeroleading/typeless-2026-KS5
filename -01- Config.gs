/**
 * Config.gs
 * Global Configuration File
 * Acts as the single source of truth for the entire reporting system.
 */

const CONFIG = Object.freeze({
  
  // 1. Global Settings
  GLOBAL: {
    CONTROL_SHEET_NAME: 'Control_Panel',
    OUTPUT_FOLDER_ID: 'YOUR_MAIN_OUTPUT_FOLDER_ID', // The parent folder for all generated reports
  },

  // 2. Authorisation Controls
  AUTH: {
    // Users who can run ANY report
    SUPER_USERS: [
      'jappleton@csg.school',
      'tnayagam@csg.school'
    ],
    // Users with restricted access to specific reports
    REPORT_SPECIFIC: {
      UCAS: [
        'nbayley@csg.school'
      ]
    }
  },

  // 3. Report Profiles
  // Each report type has its own distinct configuration profile.
  REPORTS: {
    
    PROGRESS_REVIEW: {
      name: 'Progress Review',
      templateId: 'YOUR_PR_TEMPLATE_ID'
    },

    EOY_MOCK: {
      name: 'End of Year Assessment',
      templateId: 'YOUR_EOY_TEMPLATE_ID'
    },

    UCAS: {
      name: 'UCAS Application',
      templateId: 'YOUR_UCAS_TEMPLATE_ID'
    }
  }
});