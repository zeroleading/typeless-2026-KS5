/** * Returns a list of sheet names and link IDs that match a given Regular Expression. 
 * * @param {"^202.*"} include_regex A regular expression string, or a range of regex strings. 
 * @param {FALSE} trigger Optional: Checkbox reference to force a refresh (Boolean). 
 * @return A 2-column array of matching Names and Link IDs. 
 * @customfunction 
 */
function GETSHEETNAMES_REGEX(include_regex, trigger) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let patterns = [];
  
  if (include_regex) {
    if (Array.isArray(include_regex)) {
      patterns = include_regex.flat().filter(String).map(String);
    } else {
      patterns = [String(include_regex)];
    }
  }
  
  if (patterns.length === 0) return [["Error", "Please provide a valid regex pattern."]];
  
  let regexObjects = [];
  try {
    regexObjects = patterns.map(p => new RegExp(p));
  } catch (e) {
    return [["Regex Error", e.message]];
  }
  
  const output = sheets
    .filter(sheet => {
      const name = sheet.getName();
      return regexObjects.some(regex => regex.test(name));
    })
    .map(sheet => [sheet.getName(), `#gid=${sheet.getSheetId()}`]);
    
  return output.length > 0 ? output : [["No sheets matched", ""]];
}