  
const axios = require('axios');
const ExcelJS = require('exceljs');

// Jira credentials
const JIRA_BASE_URL = 'https://amanacap.atlassian.net'; // Replace with your Jira URL
const JIRA_EMAIL = 'Bassel.Zeidan@amanacapital.com'; // Replace with your Jira email
const JIRA_API_TOKEN = 'ATATT3xFfGF0DCcoJY-m9xqicVgy7-GjVgj4AsxyFojLN1nE0n_RECZxh4bzE4OLu7DIaeOQ8ZtvE6h2KLZJAmx5Cw4-gXMNEI_Y26_eccfAQXBovl0m1UtzpTATe0s2sPQmnM2SKCZLvEHMBEE6IFlZGOsG0PtvElOmAonQeHXD_VJuO0wEVG8=321D5B40'; // Replace with your fresh API token

// Fetch Jira issues
async function fetchJiraIssues() {
    try {
        const response = await axios.get(
            `${JIRA_BASE_URL}/rest/api/3/search?jql=project=AMT`, // Replace AMT with your project key
            {
                headers: {
                    Authorization: `Basic ${Buffer.from(`${JIRA_EMAIL}:${JIRA_API_TOKEN}`).toString('base64')}`,
                    'Accept': 'application/json',
                },
            }
        );
        return response.data.issues;
    } catch (error) {
        console.error('Error fetching Jira issues:', error.response ? error.response.data : error.message);
        return [];
    }
}

// Export issues to Excel with Pie Chart Data
async function exportToExcelWithChart() {
    const issues = await fetchJiraIssues();
    if (issues.length === 0) {
        console.log('No issues found.');
        return;
    }

    // Map issues to a flat structure for Excel
    const data = issues.map(issue => ({
        Key: issue.key,
        Summary: issue.fields.summary,
        Status: issue.fields.status.name,
        Assignee: issue.fields.assignee ? issue.fields.assignee.displayName : 'Unassigned',
        Reporter: issue.fields.reporter.displayName,
        Created: issue.fields.created,
    }));

    // Aggregate data for pie chart (e.g., count issues by status)
    const statusCounts = data.reduce((counts, issue) => {
        counts[issue.Status] = (counts[issue.Status] || 0) + 1;
        return counts;
    }, {});

    // Prepare data for pie chart
    const pieChartLabels = Object.keys(statusCounts);
    const pieChartValues = Object.values(statusCounts);

    // Create workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Jira Issues');

    // Add data to worksheet
    worksheet.columns = [
        { header: 'Key', key: 'Key', width: 15 },
        { header: 'Summary', key: 'Summary', width: 50 },
        { header: 'Status', key: 'Status', width: 15 },
        { header: 'Assignee', key: 'Assignee', width: 20 },
        { header: 'Reporter', key: 'Reporter', width: 20 },
        { header: 'Created', key: 'Created', width: 25 },
    ];
    worksheet.addRows(data);

    // Add a new sheet for chart data
    const chartWorksheet = workbook.addWorksheet('Chart Data');
    chartWorksheet.addRow(['Status', 'Count']); // Add header row for the chart data
    pieChartLabels.forEach((label, index) => {
        chartWorksheet.addRow([label, pieChartValues[index]]);
    });

    // Write Excel file
    const fileName = 'jira_issues_with_chart.xlsx';
    await workbook.xlsx.writeFile(fileName);

    console.log(`Jira issues exported to ${fileName}`);
    console.log('You can create a pie chart in Excel using the data in the "Chart Data" sheet.');
}

// Execute the function
exportToExcelWithChart();
