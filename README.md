# Jira Issues Report Generator with Charts

This project generates an Excel file containing Jira issue data,  . The project is implemented in JavaScript using the **ExcelJS** library.

## Prerequisites

### Required Tools and Software
- [Node.js](https://nodejs.org/)  
- A Jira account with an active project
- Jira API Token (for authentication)
- **ExcelJS** library installed

## Setup Instructions

### 1. Clone the Repository
```bash
git clone https://github.com/BasselZeidan80/JIRA-Script-Report.git
cd JiraReportScript
```

### 2. Install Dependencies
Run the following command to install the required npm packages:
```bash
npm install
```

### 3. Configure Jira Credentials
Update the `index.js` file with your Jira credentials:

- **JIRA_BASE_URL**: Your Jira instance URL (  `https://yourdomain.atlassian.net`)
- **JIRA_EMAIL**: Your Jira email
- **JIRA_API_TOKEN**: Your Jira API token from url https://id.atlassian.com/manage-profile/security/api-tokens
- **JIRA_PROJECT_KEY**: Your Jira PROJECT_KEY

### 4. Run the Script
Generate the Excel report by running the following command:
```bash
node index.js
```

This will create a file named `jira_issues.xlsx` containing Jira issue data.

## Adding Charts


### 1. Open the Generated File
After running the script, open the `jira_issues.xlsx` file using Microsoft Excel.

### 2. Add a Pie Chart
You can use a VBA macro to automatically generate the chart. Follow these steps:

1. Open the Excel file.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module: **Insert > Module**.
4. Paste the following VBA code:


## Output
- **Excel File**: `jiraReport.xlsx`
- Includes issue data and an optional pie chart summarizing statuses.

## Notes
- The script uses the Jira API to fetch issues. Ensure you have the necessary permissions to access the project data.
- If you encounter authentication errors, double-check your email , API token and Project Key.

## License
This project is licensed under the MIT License.

## Created By:
Bassel A Zeidan...