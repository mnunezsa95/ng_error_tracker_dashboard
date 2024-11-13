function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Function").addItem("Update the Spreadsheet", "updateEntireSpreadsheet").addToUi();
}

var updateEntireSpreadsheet = () => {
  /**
   * Executes three functions: importSpecificProgramErrors, updateGradeNames, and condenseSubjectNames,
   * followed by the classifyCore function.
   * Updates the date on the README Tab.
   * Logs the execution time for the entire update process.
   * Sends out an email to a list of users to notify them whether the update was successful or not.
   *
   * @function updateEntireSpreadsheet
   * @returns {void}
   */

  // Define the frequency of update and recipient list for email notification
  const frequencyOfUpdate = 2;
  const recipientList = ["patrick.bracken@newglobe.education"];

  // Get the start time of the execution
  const startTime = new Date().getTime();

  // Access the README tab of the active spreadsheet
  const readmeTab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("README");

  // Get the current date and time in a readable format
  const currentDate = new Date().toLocaleString("en-US", {
    weekday: "short",
    year: "numeric",
    month: "long",
    day: "numeric",
    timeZoneName: "short",
    hour12: true,
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });

  // Clear previous update information in the README tab
  readmeTab.getRange("L2:N2").clearContent();

  let executionTime;

  try {
    // Execute the functions to update the spreadsheet
    importSpecificProgramErrors();
    updateGradeNames();
    condenseSubjectNames();
    classifyCoreSubjects();

    // Set the "Last Update" timestamp in the README tab
    readmeTab.getRange("K2:K2").setValue("Last Update");

    // Calculate execution time after the try block
    executionTime = new Date().getTime() - startTime;

    // Prepare success message for email notification
    const successMessage = `Successful Update of Global Academic Error Dashboard.\n\nUpdate Time: ${currentDate}\n\nExecution Time: ${executionTime} milliseconds`;

    // Log success message
    console.log(successMessage);

    // Send email notification for successful update
    sendEmail(recipientList, "Successful AET Update", successMessage);
  } catch (err) {
    // Set "Last Failed Updated" timestamp in the README tab
    readmeTab.getRange("K2:K2").setValue("Last Failed Updated");

    // Prepare failure message for email notification
    const failureMessage = `Unsuccessful Update of Global Academic Error Dashboard.\n\nError Message: ${err}\n\nWill try again in ${frequencyOfUpdate} Hour(s).`;

    // Log error message
    console.log(err);
    console.log(failureMessage);

    // Send email notification for unsuccessful update
    sendEmail(recipientList, "Unsuccessful AET Update", failureMessage);
  } finally {
    const endTime = new Date().getTime(); // Get the end time of the execution
    executionTime = endTime - startTime; // Calculate total execution time if it wasn't set in the try block
    readmeTab.getRange("N2:N2").setValue(currentDate); // Set the current date in the README tab
  }
};

//======================================================================================================================
const sendEmail = (recipientList = "patrick.bracken@newglobe.education", subjectLine, message) => {
  const recipients = recipientList.length > 1 ? recipientList.join(",") : "patrick.bracken@newglobe.education";
  const subject = subjectLine;
  const body = message;

  GmailApp.sendEmail(recipients, subject, body);
};

// =====================================================================================================================
function importSpecificProgramErrors() {
  let spreadsheets = [
    { id: "18cK8hMlNjC8JclnqQt-Wm-oeBbrrLNyWErvePEc_Z1E", program: "BayelsaPRIME" },
    { id: "1InSSjWL_OCXrK6KvYZFweREmCOh9ayEuPbC0IpfTnXQ", program: "Bridge Andhra Pradesh" },
    { id: "1Z49FAwkq8c0bSFPpG2Ki3Gh_t9XjiqvWov2D_1-rNgk", program: "Bridge Kenya" },
    { id: "1tVqoavTYriWY50JAz7ludDW0wttW6Si6YnqkOTNZlbk", program: "Bridge Liberia" },
    { id: "1YIUIxMtfpVgRSgaOp6BwNFs3nbDPVwXANfib1uzrc-c", program: "Bridge Nigeria" },
    { id: "1AgntRauSd70NYGNcU_tgNxAKo-hfuub8NKaptsUbMTM", program: "Bridge Uganda" },
    { id: "12uL3uodPrXZpoQ6ZZ3Bmo_em9ODzq1nyLsXFm2axGwY", program: "EdoBEST" },
    { id: "1ImTQcgqV3gY4aNe_o1w33MOXwg-DyrhvCYgfIXMoXfQ", program: "EKOEXCEL" },
    { id: "1JrnrVwDf8kdzko1NXFriJ6FfdOLW5bH9vDQIuC2_ZHE", program: "KwaraLEARN" },
    { id: "1-ItQ14rgJAIMYN3fWW1t08ciT0Qp54jrB8zZGsWFJmk", program: "RwandaEQUIP" },
    { id: "1hoQl1qeK7C0C7hx1IDRMgMyiUvsreN7ji5wBVxDqDrM", program: "STAR Education" },
  ];

  const destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Error Database [Aggregate]");
  destinationSheet.getRange("A2:S").clearContent();

  let startRow = 2;

  for (var i = 0; i < spreadsheets.length; i++) {
    var spreadsheetId = spreadsheets[i].id;
    var programName = spreadsheets[i].program;

    var sourceSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Error Tracker");

    var lastRow = sourceSheet.getRange("A:A").getValues().filter(String).length;

    // Skip the sheet if there are no data rows starting from A2
    if (lastRow < 2) {
      continue;
    }

    var dataRange = sourceSheet.getRange("A2:Q" + lastRow);
    var values = dataRange.getValues();

    var newData = values.map(function (row) {
      return [programName].concat(row);
    });

    destinationSheet.getRange(startRow, 1, newData.length, newData[0].length).setValues(newData);
    startRow += newData.length;
  }
}

// =====================================================================================================================
function updateGradeNames() {
  /**
   * Trims white space for each cell in the grade column &
   * Reduces the grade names to common groups as specified and returns these condensed subject names
   */

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Error Database [Aggregate]");
  const range = sheet.getRange("G2:G"); // Start from G2 to skip the header
  const values = range.getValues();

  // Create an array to hold updated values for the new column
  const updatedValues = values.map((row) => {
    const cellValue = row[0].trim();
    let newValue = cellValue; // Initialize with the original value

    if (cellValue === "Standard 1" || cellValue === "Grade 1") {
      newValue = "Primary 1";
    } else if (cellValue === "Standard 2" || cellValue === "Grade 2") {
      newValue = "Primary 2";
    } else if (cellValue === "Standard 3" || cellValue === "Grade 3") {
      newValue = "Primary 3";
    } else if (cellValue === "Standard 4" || cellValue === "Grade 4") {
      newValue = "Primary 4";
    } else if (cellValue === "Standard 5" || cellValue === "Grade 5") {
      newValue = "Primary 5";
    } else if (cellValue === "Grade 6" || cellValue === "Class 6") {
      newValue = "Primary 6";
    } else if (cellValue === "Class 7" || cellValue === "Grade 7" || cellValue === "Primary 7") {
      newValue = "JSS 1";
    } else if (cellValue === "Class 8") {
      newValue = "JSS 2";
    } else if (cellValue === "Class 9") {
      newValue = "JSS 3";
    } else if (cellValue === "Class 10" || cellValue === "Grade 10") {
      newValue = "Class 10";
    }

    return [newValue];
  });

  // Write the updated values back to the sheet in the same column (G), starting from G2 to match the original range
  sheet.getRange(2, 7, updatedValues.length, 1).setValues(updatedValues);
}

// =====================================================================================================================
function condenseSubjectNames() {
  /**
   * Trims white space for each cell in the subject column &
   * Reduces the subject names to common groups as specified and returns these condensed subject names
   */

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Error Database [Aggregate]");
  let range = sheet.getRange("H:H");
  let gradeRange = sheet.getRange("G:G").getValues().flat();
  let values = range.getValues().map((row) => [row[0].trim()]);

  const gradesFourAndUp = ["Primary 4", "Primary 5", "Primary 6", "JSS 1", "JSS 2", "JSS 3", "Class 10"];

  for (let i = 0; i < values.length; i++) {
    let cellValue = values[i][0];
    let gradeValue = gradeRange[i];

    if (
      gradesFourAndUp.includes(gradeValue) &&
      (cellValue === "Basic Science and Technology" ||
        cellValue === "Basic Science and Technology - Basic Science" ||
        cellValue === "Science" ||
        cellValue === "Science & Technology")
    ) {
      sheet.getRange(i + 1, 8).setValue("Science (P4+)");
    } else if (
      cellValue === "BECE BST Prep" ||
      cellValue === "BECE English Prep" ||
      cellValue === "BECE Mathematics Prep" ||
      cellValue === "BECE National Values Prep" ||
      cellValue === "BECE Pre-Vocational Studies Prep"
    ) {
      sheet.getRange(i + 1, 8).setValue("BECE Prep");
    } else if (
      cellValue === "HSLC Prep - English" ||
      cellValue === "HSLC Prep - Mathematics" ||
      cellValue === "HSLC Prep - Science" ||
      cellValue === "HSLC Prep - Social Science"
    ) {
      sheet.getRange(i + 1, 8).setValue("HSLC Prep");
    } else if (
      cellValue === "KPSEA Prep Creative Arts" ||
      cellValue === "KPSEA Prep Creative Arts and Social Studies" ||
      cellValue === "KPSEA Prep English" ||
      cellValue === "KPSEA Prep Integrated Sciences" ||
      cellValue === "KPSEA Prep Kiswahili" ||
      cellValue === "KPSEA Prep Mathematics" ||
      cellValue === "KPSEA Prep Science & Technology" ||
      cellValue === "KPSEA Prep Social Studies"
    ) {
      sheet.getRange(i + 1, 8).setValue("KPSEA Prep");
    } else if (
      cellValue === "Co-curricular" ||
      cellValue === "Co-Curricular" ||
      cellValue === "Co Curricular" ||
      cellValue === "Cocurricular" ||
      cellValue === "Clubs"
    ) {
      sheet.getRange(i + 1, 8).setValue("Co-Curricular");
    } else if (cellValue.includes("English Studies") && cellValue.includes("Reading")) {
      sheet.getRange(i + 1, 8).setValue("English Studies - Reading");
    } else if (cellValue.includes("English Studies") && cellValue.includes("Language")) {
      sheet.getRange(i + 1, 8).setValue("English Studies - Language");
    } else if (
      cellValue === "Mathematics 1" ||
      cellValue === "Mathematics 2" ||
      cellValue === "Mathematics 3" ||
      cellValue.includes(" Mathematics") ||
      cellValue.includes("Mathematics 1 ") ||
      cellValue.includes("Mathematics 2 ")
    ) {
      sheet.getRange(i + 1, 8).setValue("Mathematics");
    } else if (cellValue === "Maths" || cellValue === "Math") {
      sheet.getRange(i + 1, 8).setValue("Maths");
    } else if (cellValue === "Supplementary English" || cellValue === "Supplemental English") {
      sheet.getRange(i + 1, 8).setValue("Supplementary English");
    } else if (cellValue === "Supplementary Maths" || cellValue === "Supplemental Maths") {
      sheet.getRange(i + 1, 8).setValue("Supplementary Maths");
    } else if (cellValue.includes("Preparatory English")) {
      sheet.getRange(i + 1, 8).setValue("Preparatory English");
    } else if (cellValue.includes("Preparatory Maths")) {
      sheet.getRange(i + 1, 8).setValue("Preparatory Maths");
    } else if (cellValue !== "Social Studies and Science" && cellValue.includes("Social Studies")) {
      sheet.getRange(i + 1, 8).setValue("Social Studies");
    } else if (
      cellValue.includes("Day") ||
      cellValue.includes("day") ||
      cellValue.includes(" Day") ||
      cellValue.includes(" Day ") ||
      cellValue.includes(" holiday") ||
      cellValue.includes(" holiday ")
    ) {
      sheet.getRange(i + 1, 8).setValue("Holiday Lesson(s)");
    }
  }
}

// =====================================================================================================================
function classifyCoreSubjects() {
  /**
   * Looks at Column I in the Error Database Tab (Lesson Code) and determines the appropriate level value based on the values in the column.
   * Trims leading and trailing whitespace from each cell value before comparison.
   * Compares the trimmed value to a list of predefined values and returns the corresponding level value to Column R (Core Subject / Level).
   *
   * @function classifyCore
   * @returns {void}
   */

  // Access the active spreadsheet and the 'Error Database [Aggregate]' sheet
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Error Database [Aggregate]");

  // Get the range of values in Column I (Lesson Code)
  var lessonCodeColumn = currentSheet.getRange("I2:I");

  // Get the values from the range and trim whitespace from each value
  var lessonCodeColValues = lessonCodeColumn.getValues().map((row) => row[0].trim());

  // Get the range where the output values will be placed (Column R)
  var destinationColumn = currentSheet.getRange("R2:R");

  // Initialize an array to store the output values
  var outputValues = [];

  // Iterate through each value in the Lesson Code column
  for (let i = 0; i < lessonCodeColValues.length; i++) {
    var cellValue = lessonCodeColValues[i];
    var level = "";

    // Determine the level based on the prefix of the Lesson Code
    if (cellValue.startsWith("LAL")) {
      level = "Reading Level A";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LBL")) {
      level = "Reading Level B";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LCL")) {
      level = "Reading Level C";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LDL")) {
      level = "Reading Level D";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LEL")) {
      level = "Reading Level E";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LALG")) {
      level = "Language Level A";
    } else if (cellValue.startsWith("LBLG")) {
      level = "Language Level B";
    } else if (cellValue.startsWith("LAN")) {
      level = "Mathematics Level A";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LBN")) {
      level = "Mathematics Level B";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LCN")) {
      level = "Mathematics Level C";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LDN")) {
      level = "Mathematics Level D";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LEN")) {
      level = "Mathematics Level E";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    }

    // Add the determined level to the outputValues array
    outputValues.push([level]);
  }

  // Set the values in the destination column (Column R) to the outputValues array
  destinationColumn.setValues(outputValues);
}
