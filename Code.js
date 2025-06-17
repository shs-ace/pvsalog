function onFormSubmit(e) {
  var sheet = e.source.getActiveSheet();
  var name = e.values[1].trim(); // Trim the name to remove leading and trailing spaces
  var dateResponse = e.values[2]; // Change the index for the date column (assuming it's the second column)
  var activities = e.values[3]; // Change the index for the activities column
  var hours = e.values[4]; // Change the index for the hours column

  // Check if the sheet with the name exists, create if not
  var nameSheet = e.source.getSheetByName(name);
  if (!nameSheet) {
    nameSheet = e.source.insertSheet(name);
    // Add headers to the newly created sheet
    nameSheet.appendRow(["Date", "Activity", "Hours"]); // Add more headers if needed

    // Set "Total:" and the total value in cells D1 and E1
    nameSheet.getRange("D1").setValue("Total:");
    nameSheet.getRange("E1").setFormula("=SUM(C:C)");
    
    // Set "PVSA(11-15)" in cell F1
    nameSheet.getRange("F1").setValue("PVSA(11-15)");

    // Set "PVSA(16-25)" in cell H1
    nameSheet.getRange("H1").setValue("PVSA(16-25)");
  }

  // Convert the date response to a Date object
  var timestamp = new Date(dateResponse);
  var formattedDate = Utilities.formatDate(timestamp, 'GMT', 'MM/dd/yyyy'); // Format the date part

  // Record the data in the sheet with the corresponding name
  nameSheet.appendRow([formattedDate, activities, hours]); // Add more data as needed
  
  // Calculate the total volunteer hours in column E
  var totalHours = nameSheet.getRange("E1").getValue();
  
  // Determine the PVSA status based on total hours and set it in cell G1
  var pvsaStatus = "";
  if (totalHours >= 50 && totalHours <= 74) {
    pvsaStatus = "BRONZE";
  } else if (totalHours >= 75 && totalHours <= 99) {
    pvsaStatus = "SILVER";
  } else if (totalHours > 100) {
    pvsaStatus = "GOLD";
  } else {
    pvsaStatus = "N/A";
  }
  nameSheet.getRange("G1").setValue(pvsaStatus);


  // Determine the PVSA status for the 16-25 age group based on total hours and set it in cell J1
  var pvsaStatus16to25 = "";
  if (totalHours >= 100 && totalHours <= 174) {
    pvsaStatus16to25 = "BRONZE";
  } else if (totalHours >= 175 && totalHours <= 249) {
    pvsaStatus16to25 = "SILVER";
  } else if (totalHours >= 250) {
    pvsaStatus16to25 = "GOLD";
  } else {
    pvsaStatus16to25 = "N/A";
  }
  nameSheet.getRange("I1").setValue(pvsaStatus16to25);
}
