// Define a constant for the height of each section in the Answer Sheet.
const SECTION_HEIGHT = 13;
generateAnswerSections();

// Main function to organize the workflow for setting up the Issue List and Answer Sheet.
function generateAnswerSections() {
  // Initialize spreadsheet and sheets.
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const issueSheet = s.getSheetByName("Issue List");
  const answerSheet = s.getSheetByName("Answer Sheet");

  // Call functions to modify the Issue List sheet.
  setExtraTitleOnIssueList(issueSheet);
  setHyperlinkFormulaOnIssueList(issueSheet, answerSheet);
  setAnswerStatusFormulasOnIssueList(issueSheet);
  setFormattingOnIssueList(issueSheet);

  // Call functions to modify the Answer Sheet.
  setQuestionAndAnswerOptionsOnAnswerSheet(issueSheet, answerSheet, s);
  setFormattingOnAnswerSheet(answerSheet);
}

// Adds additional titles for 'Link' and 'Status' columns in the Issue List sheet.
function setExtraTitleOnIssueList(issueSheet) {
  let currentRow = 1;
  issueSheet.getRange(currentRow, 6).setValue("Link");
  issueSheet.getRange(currentRow, 7).setValue("Status");
}

// Sets up a hyperlink formula in the issue list that allows users to click and jump directly to the corresponding answer section.
function setHyperlinkFormulaOnIssueList(issueSheet, answerSheet) {
  const gid = answerSheet.getSheetId();
  const formula = '=ARRAYFORMULA(IF(A2:A="", "", HYPERLINK("#gid=' + gid + '&range=A" & ((ROW(A2:A)-2)*13+1), "Answer")))';
  issueSheet.getRange("F2").setFormula(formula);
}

// Sets up formulas to track the status of answers based on completion in the Issue List sheet.
function setAnswerStatusFormulasOnIssueList(issueSheet) {
  const lastRow = issueSheet.getLastRow(); // Get the last row with data in the issue list to know where to stop.

  for (let i = 2; i <= lastRow; i++) {
    let startRow = (i - 2) * 13 + 3; // Calculate the start row for answer data in the answer sheet based on the issue row.
    let endRow = startRow + 8; // Calculate the end row for checking completeness of answers.

    // Only set the formula if there is a corresponding link.
    if (issueSheet.getRange("F" + i).getValue() !== "") {
      // Construct and set a formula that checks if all required answer cells are filled to determine the status as "DONE" or "WIP".
      let formula = '=IF(COUNTA(INDIRECT("Answer Sheet!D' + startRow + ':D' + endRow + '"))=9, "DONE", "WIP")';
      issueSheet.getRange("G" + i).setFormula(formula);
    }
  }
}

// Sets the visual formatting for the Issue List sheet.
function setFormattingOnIssueList(issueSheet) {
  const lastRow = issueSheet.getLastRow();

  issueSheet.createTextFinder("Tester - (non-subtask level)").replaceAllWith("Tester");;

  const contentRangeCG = issueSheet.getRange("C2:G" + lastRow);
  basicFormatting(contentRangeCG, "center", "middle", "Arial", 11);

  const contentRangeAB = issueSheet.getRange("A2:B" + lastRow);
  basicFormatting(contentRangeAB, "left", "middle", "Arial", 11);

  const headerRangeAG = issueSheet.getRange("A1:G1");
  basicFormatting(headerRangeAG, "center", "middle", "Arial", 12);
  headerRangeAG
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_THICK);
}

// Generates questions and answer options on the Answer Sheet based on the issues listed and template.
function setQuestionAndAnswerOptionsOnAnswerSheet(issueSheet, answerSheet, s) {
  const sectionHeight = SECTION_HEIGHT;
  let currentRow = 1;
  const templateSheet = s.getSheetByName("Template");

  // Get key and summary from issue list and questions from template.
  const issues = issueSheet.getRange("A2:B" + issueSheet.getLastRow()).getValues();
  const templateText = templateSheet.getRange("A2:C11").getValues();

  // Retrieve data validation rules from the template sheet.
  const dataValidationA = templateSheet.getRange("D3").getDataValidation();
  const dataValidationB = templateSheet.getRange("D7").getDataValidation();
  const dataValidationC = templateSheet.getRange("D8").getDataValidation();

  answerSheet.clear();

  // Loop through each issue to create a new section for it in the answer sheet.
  for (let i = 0; i < issues.length; i++) {
    const issueKey = issues[i][0];
    const issueSummary = issues[i][1];

    if (issueKey && issueSummary) {
      // Set the "Issue" header with the key and summary for each issue in the answer sheet.
      answerSheet.getRange(currentRow, 1).setValue("Issue");
      answerSheet.getRange(currentRow, 2).setValue(issueKey);
      answerSheet.getRange(currentRow, 3).setValue(issueSummary);
      answerSheet.getRange(currentRow + 1, 4).setValue("Answer");

      // Fill the template text into the designated range.
      answerSheet.getRange(currentRow + 1, 1, templateText.length, 3).setValues(templateText);

      // Set data validation for answer options based on the template.
      // This part uses loops to apply Yes/No data validation rules to multiple cells.
      for (let j = 2; j <= 5; j++) {
        answerSheet.getRange(currentRow + j, 4).setDataValidation(dataValidationA);
      }

      for (let k = 8; k <= 10; k++) {
        answerSheet.getRange(currentRow + k, 4).setDataValidation(dataValidationA);
      }

      // Apply different data validation rules to specific cells.
      answerSheet.getRange(currentRow + 6, 4).setDataValidation(dataValidationB);
      answerSheet.getRange(currentRow + 7, 4).setDataValidation(dataValidationC);

      // Update the current row to prepare for the next issue, leaving enough space for the current issue's answer section.
      currentRow += sectionHeight;
    }
  }
}

// Sets the visual formatting for the Answer Sheet.
function setFormattingOnAnswerSheet(answerSheet) {
  const lastRow = answerSheet.getLastRow();
  const lastColumn = answerSheet.getLastColumn();

  const allRange = answerSheet.getRange(1, 1, lastRow, lastColumn); // Start row, start column, end row, end column
  basicFormatting(allRange, "left", "middle", "Arial", 11);

  const dCoulmn = answerSheet.getRange("D:D");
  dCoulmn.setHorizontalAlignment("center");

  for (let row = 1; row <= lastRow; row += 13) {
    const twoHeaderRange = answerSheet.getRange(row, 1, 2, lastColumn);
    twoHeaderRange
      .setFontSize(12)
      .setBackground("#d9ead3")
      .setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_THICK);
    for (let titleRow = 2; titleRow <= lastRow; titleRow += 13) {
      const onlyQuestionHeaderRange = answerSheet.getRange(titleRow, 1, 1, lastColumn);
      onlyQuestionHeaderRange.setFontWeight("bold");
    }
  }
}

function basicFormatting(range, horizontal, vertical, font, fontSize) {
  range
    .setHorizontalAlignment(horizontal)
    .setVerticalAlignment(vertical)
    .setFontFamily(font)
    .setFontSize(fontSize);
}