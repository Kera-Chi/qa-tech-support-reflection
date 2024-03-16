## Spreadsheet Issue Tracker and Answer Sheet Automation

### Overview

The script generates hyperlinks in an Issue List sheet that direct to corresponding detailed answers in an Answer Sheet. It also includes features for tracking the completion status of each issue's answers and applies uniform formatting across the sheets for better readability and organization.

### Features

- **Automated Issue List to Answer Sheet Links**: Dynamically creates hyperlinks in the `Issue List`, when clicked, navigate to the respective detailed answer section in the `Answer Sheet`.
- **Status Tracking**: Implements formulas in the Issue List to track whether the answers for a given issue are complete, marking them as **"DONE"** or **"WIP"** (Work In Progress).
- **Template-Based Answer Sections**: Generates structured answer sections in the Answer Sheet for each issue listed in the Issue List, using a predefined template for consistency.
- **Customizable Formatting**: Applies a uniform visual style across both sheets to enhance readability and maintain consistency.

### Examples

You may view the [Example Google Sheets file](https://docs.google.com/spreadsheets/d/17bXaQZhOm4XIp8eGwpQQSJWO9Atp8A3DjL-HAAfv_jE/edit#gid=997947630).
Please note that the file is view-only and you should make a copy of it if you wish to interact with it.

### Setup Instructions

1. **Initial Setup**:
   - Open your Google Spreadsheet.
   - Ensure your spreadsheet contains three sheets named `Issue List`, `nswer Sheet`, and `Template`.
   - The `Issue List` should have titles in the first row and issues listed from row 2 onwards, with columns A and B containing the issue key and summary, respectively.
   - The `Template` sheet should contain predefined questions and data validation rules for the answers.

2. **Script Deployment**:
   - Open the Script Editor from the Extensions menu in Google Sheets.
   - Copy the provided script into the Script Editor.
   - Save the script with a suitable name (e.g., "IssueTrackerAutomation").

3. **Executing the Script**:
   - Run the `generateAnswerSections` function from the Script Editor to execute the automation.
   - Grant the necessary permissions when prompted.

### Detailed Function Descriptions

- **generateAnswerSections**: Main function orchestrating the workflow. Initializes the spreadsheet and calls specific functions to modify the Issue List and Answer Sheet.

- **setExtraTitleOnIssueList**: Adds additional titles for **"Link"** and **"Status"** columns in the Issue List.

- **setHyperlinkFormulaOnIssueList**: Sets up hyperlinks to the Answer Sheet, enabling direct navigation.

- **setAnswerStatusFormulasOnIssueList**: Implements formulas to track the completion status of each issue's answers.

- **setFormattingOnIssueList**: Applies consistent formatting to the Issue List for improved readability.

- **setQuestionAndAnswerOptionsOnAnswerSheet**: Generates questions and answer options for each issue in the Answer Sheet using a predefined template.

- **setFormattingOnAnswerSheet**: Applies uniform formatting to the Answer Sheet, enhancing visual consistency and readability.

- **basicFormatting**: Utility function used to apply basic formatting options to a range of cells.

### Notes

- Ensure the column A and B in `Issue List` sheet contain data for tracking issues.
- Ensure the `Template` sheet is properly set up with the desired questions, answer options, and data validation rules.
- The **SECTION_HEIGHT** constant is configurable and should be adjusted based on the number of rows required for each answer section in the Answer Sheet.
- Originally designed to function with Jira Cloud for Sheets, it automatically generates the data and executes the remaining steps with the provided script.