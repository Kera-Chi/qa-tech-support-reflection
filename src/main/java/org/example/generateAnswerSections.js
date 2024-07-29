function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Tripla QA')
        .addItem('Generate Answer List', 'generateAnswerSections')
        .addToUi();
}

function menuItem() {
    SpreadsheetApp.getUi().alert('Generating answer list...');
}

// Define a constant for the height of each section in the Answer Sheet.
const SECTION_HEIGHT = 13;

// Main function to organize the workflow for setting up the Issue List and Answer Sheet.
function generateAnswerSections() {
    // Initialize spreadsheet and sheets.
    const s = SpreadsheetApp.getActiveSpreadsheet();
    const issueSheet = s.getSheetByName("Issue List");
    const answerSheet = s.getSheetByName("Answer Sheet");

    // Call functions to modify the Issue List sheet.
    setExtraTitleOnIssueList(issueSheet);
    setHyperlinkFormulaOnIssueList(issueSheet, answerSheet);
    setDevAnswerStatusFormulasOnIssueList(issueSheet);
    setQaAnswerStatusFormulasOnIssueList(issueSheet);
    setFormattingOnIssueList(issueSheet);

    // Call functions to modify the Answer Sheet.
    setQuestionAndAnswerOptionsOnAnswerSheet(issueSheet, answerSheet, s);
    setFormattingOnAnswerSheet(answerSheet);
}

// Adds additional titles for 'Link' and 'Status' columns in the Issue List sheet.
function setExtraTitleOnIssueList(issueSheet) {
    let currentRow = 1;
    issueSheet.getRange(currentRow, 7).setValue("Link");
    issueSheet.getRange(currentRow, 8).setValue("Dev Status");
    issueSheet.getRange(currentRow, 9).setValue("QA Status");
}

// Sets up a hyperlink formula in the issue list that allows users to click and jump directly to the corresponding answer section.
function setHyperlinkFormulaOnIssueList(issueSheet, answerSheet) {
    const gid = answerSheet.getSheetId();
    const formula = '=ARRAYFORMULA(IF(A2:A="", "", HYPERLINK("#gid=' + gid + '&range=A" & ((ROW(A2:A)-2)*13+1), "Answer")))';
    issueSheet.getRange("G2").setFormula(formula);
}

// Sets up formulas to track the status of answers based on completion in the Issue List sheet.
function setDevAnswerStatusFormulasOnIssueList(issueSheet) {
    const lastRow = issueSheet.getLastRow(); // Get the last row with data in the issue list to know where to stop.

    for (let i = 2; i <= lastRow; i++) {
        let startRow = (i - 2) * 13 + 3; // Calculate the start row for answer data in the answer sheet based on the issue row.
        let endRow = startRow + 5; // Calculate the end row for checking completeness of answers.

        // Only set the formula if there is a corresponding link.
        if (issueSheet.getRange("G" + i).getValue() !== "") {
            // Construct and set a formula that checks if all required answer cells are filled to determine the status as "DONE" or "WIP".
            let formula = '=IF(COUNTA(INDIRECT("Answer Sheet!D' + startRow + ':D' + endRow + '"))=6, "DONE", "WIP")';
            issueSheet.getRange("H" + i).setFormula(formula);
        }
    }
}

function setQaAnswerStatusFormulasOnIssueList(issueSheet) {
    const lastRow = issueSheet.getLastRow();

    for (let i = 2; i <= lastRow; i++) {
        let startRow = (i - 2) * 13 + 9;
        let endRow = startRow + 2;

        if (issueSheet.getRange("G" + i).getValue() !== "") {
            let formula = '=IF(COUNTA(INDIRECT("Answer Sheet!D' + startRow + ':D' + endRow + '"))=3, "DONE", "WIP")';
            issueSheet.getRange("I" + i).setFormula(formula);
        }
    }
}

// Sets the visual formatting for the Issue List sheet.
function setFormattingOnIssueList(issueSheet) {
    const lastRow = issueSheet.getLastRow();

    issueSheet.createTextFinder("Tester - (non-subtask level)").replaceAllWith("Tester");;

    const contentRangeLeft = issueSheet.getRange("A2:E" + lastRow);
    basicFormatting(contentRangeLeft, "left", "middle", "Arial", 11);

    const contentRangeCenter = issueSheet.getRange("F2:I" + lastRow);
    basicFormatting(contentRangeCenter, "center", "middle", "Arial", 11);

    // Set date format for F column starting from row 2
    const dateRange = issueSheet.getRange("F2:F" + lastRow);
    dateRange.setNumberFormat("yyyy-mm-dd");

    const headerRange = issueSheet.getRange("A1:I1");
    basicFormatting(headerRange, "center", "middle", "Arial", 12);
    headerRange
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
    const issues = issueSheet.getRange("A2:D" + issueSheet.getLastRow()).getValues();
    const templateText = templateSheet.getRange("A2:C11").getValues();

    // Split additional template text into three parts
    const devRootCauseTemplateText = templateSheet.getRange("E3:E5").getValues();
    const devPreventionTemplateText = templateSheet.getRange("E6:E8").getValues();
    const qaPreventionTemplateText = templateSheet.getRange("E9:E11").getValues();

    // Retrieve data validation rules from the template sheet.
    const dataValidationA = templateSheet.getRange("D3").getDataValidation();

    answerSheet.clear();

    // Loop through each issue to create a new section for it in the answer sheet.
    for (let i = 0; i < issues.length; i++) {
        const issueKey = issues[i][0];
        const mainIssueType = issues[i][1];
        const issueSummary = issues[i][2];

        if (issueKey && issueSummary) {
            // Set the "Issue" header with the key and summary for each issue in the answer sheet.
            answerSheet.getRange(currentRow, 1).setValue("Issue");
            answerSheet.getRange(currentRow, 2).setValue(issueKey);
            answerSheet.getRange(currentRow, 3).setValue(issueSummary);
            answerSheet.getRange(currentRow + 1, 4).setValue("Answer");
            answerSheet.getRange(currentRow + 1, 5).setValue("Note - Root Cause & Prevention");
            answerSheet.getRange(currentRow , 5).setValue(mainIssueType);

            // Fill the question template text into the designated range.
            answerSheet.getRange(currentRow + 1, 1, templateText.length, 3).setValues(templateText);

            // Fill the additional template text into designated groups and merge cells
            answerSheet.getRange(currentRow + 2, 5, 3, 1).merge().setValue(devRootCauseTemplateText);
            answerSheet.getRange(currentRow + 5, 5, 3, 1).merge().setValue(devPreventionTemplateText);
            answerSheet.getRange(currentRow + 8, 5, 3, 1).merge().setValue(qaPreventionTemplateText);

            // Set data validation for answer options based on the template.
            // This part uses loops to apply Yes/No data validation rules to multiple cells.
            for (let j = 2; j <= 10; j++) {
                answerSheet.getRange(currentRow + j, 4).setDataValidation(dataValidationA);
            }

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

    const dCoulmn = answerSheet.getRange("D:E");
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

            // Align text to left for the merged cells under "Note - Root Cause & Prevention"
            answerSheet.getRange(titleRow + 1, 5, 3, 1).setHorizontalAlignment("left");
            answerSheet.getRange(titleRow + 4, 5, 3, 1).setHorizontalAlignment("left");
            answerSheet.getRange(titleRow + 7, 5, 3, 1).setHorizontalAlignment("left");
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