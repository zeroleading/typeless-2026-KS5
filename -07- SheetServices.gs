/** * SheetServices.gs 
 * Handles spreadsheet interaction events, such as onEdit triggers. 
 */
const SheetServices = {
  handleOnEdit: function(e) {
    if (!e || !e.range) return;
    
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    
    // Check if it's a valid subject sheet
    const subjectRegex = /^([A-Z][a-z]|EnL)$/;
    if (!subjectRegex.test(sheetName)) return;
    
    const ss = e.source;
    const rangeName = `${sheetName}!thisSubjectAssessment`;
    const tableRange = ss.getRangeByName(rangeName);
    if (!tableRange) return;
    
    const tableStartRow = tableRange.getRow();
    const tableStartCol = tableRange.getColumn();
    const headers = tableRange.getValues()[0].map(h => String(h).toLowerCase().replace(/\s+/g, ''));
    
    const crntIdx = headers.indexOf('crnt');
    const ns1Idx = headers.indexOf('≣nextsteps1');
    const ns2Idx = headers.indexOf('≣nextsteps2');
    
    if (crntIdx === -1) return;
    
    const crntAbsCol = tableStartCol + crntIdx;
    const editedColStart = e.range.getColumn();
    const editedColEnd = e.range.getLastColumn();
    
    // Only proceed if the "CRNT" column was edited
    if (crntAbsCol < editedColStart || crntAbsCol > editedColEnd) return;
    
    const dataStartRow = tableStartRow + 2;
    const editedRowStart = e.range.getRow();
    const editedRowEnd = e.range.getLastRow();
    
    if (editedRowEnd < dataStartRow) return;
    
    const startClearRow = Math.max(editedRowStart, dataStartRow);
    const numRowsToClear = editedRowEnd - startClearRow + 1;
    
    // Clear dependent "Next Steps" columns to enforce re-selection
    if (ns1Idx > -1) sheet.getRange(startClearRow, tableStartCol + ns1Idx, numRowsToClear, 1).clearContent();
    if (ns2Idx > -1) sheet.getRange(startClearRow, tableStartCol + ns2Idx, numRowsToClear, 1).clearContent();
  }
};