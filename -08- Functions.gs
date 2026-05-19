/**
 * Returns a list of sheet names and link IDs, excluding specified sheets.
 *
 * @param {"Sheet1, Sheet2"} exclude_list A comma-separated string or a range of names to exclude.
 * @param {FALSE} trigger Optional: Checkbox reference to force a refresh (Boolean).
 * @return A 2-column array of Names and Link IDs.
 * @customfunction
 */
function GETSHEETNAMES(exclude_list, trigger) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var output = [];
  var exclusions = [];

  // Handle exclusions input (can be string or range)
  if (exclude_list) {
    if (Array.isArray(exclude_list)) {
      // If input is a range, flatten it to a 1D array
      exclusions = exclude_list.flat().map(String);
    } else if (typeof exclude_list === 'string') {
      // If input is a string, split by commas
      exclusions = exclude_list.split(',').map(function(s) { return s.trim(); });
    }
  }

  // Loop through sheets
  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    
    // Check if the name is in the exclusion list
    if (exclusions.indexOf(name) === -1) {
      // Create the link fragment (e.g., "#gid=12345")
      var link = "#gid=" + sheet.getSheetId();
      output.push([name, link]);
    }
  });

  return output;
}