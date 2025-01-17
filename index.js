
const axios = require('axios');
const ExcelJS = require('exceljs');

// set your domain BaseUrl
const JIRA_BASE_URL = 'https://XXXXXXXX.atlassian.net'; 
// set your email 
const JIRA_EMAIL = 'Bassel.Zeidan@XXXXXXXXX.com'; 
// set your API token (check readme file to know how you can get the API token)
const JIRA_API_TOKEN = 'XXXXXXXXXXX';  

async function fetchJiraIssues() {
    try {
        const response = await axios.get(
            `${JIRA_BASE_URL}/rest/api/3/search?jql=project=XXXXXX`, // Replace XXXXXX with Your Project Key
            {
                headers: {
                    Authorization: `Basic ${Buffer.from(`${JIRA_EMAIL}:${JIRA_API_TOKEN}`).toString('base64')}`,
                    Accept: 'application/json',
                },
            }
        );
        return response.data.issues;
    } catch (error) {
        console.error('Error fetching Jira issues:', error.response ? error.response.data : error.message);
        return [];
    }
}

async function exportToExcelWithChart() {
    const issues = await fetchJiraIssues();
    if (issues.length === 0) {
        console.log('No issues found.');
        return;
    }

    const data = issues.map(issue => ({
        Key: issue.key,
        Summary: issue.fields.summary || 'N/A',
        Status: issue.fields.status ? issue.fields.status.name : 'Unknown',
        Assignee: issue.fields.assignee ? issue.fields.assignee.displayName : 'Unassigned',
        Reporter: issue.fields.reporter ? issue.fields.reporter.displayName : 'Unknown',
        Created: issue.fields.created || 'N/A',
        FixVersions: issue.fields.fixVersions.map(fv => fv.name).join(', ') || 'None',
    }));

    const statusCounts = data.reduce((counts, issue) => {
        counts[issue.Status] = (counts[issue.Status] || 0) + 1;
        return counts;
    }, {});

    const pieChartLabels = Object.keys(statusCounts);
    const pieChartValues = Object.values(statusCounts);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Jira Issues');

    worksheet.columns = [
        { header: 'Key', key: 'Key', width: 15 },
        { header: 'Summary', key: 'Summary', width: 50 },
        { header: 'Status', key: 'Status', width: 15 },
        { header: 'Assignee', key: 'Assignee', width: 20 },
        { header: 'Reporter', key: 'Reporter', width: 20 },
        { header: 'Fix Versions', key: 'FixVersions', width: 25 },
        { header: 'Created', key: 'Created', width: 30 },
    ];
    worksheet.addRows(data);

     
    const balancedGreen = '37906D';
    const lightGreen = 'D4EBE4';

     
    worksheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: balancedGreen },
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell(cell => {
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
            };
            if (rowNumber > 1 && rowNumber % 2 === 0) {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: lightGreen },
                };
            }
        });
    });

    const chartWorksheet = workbook.addWorksheet('Pivot Table');
    chartWorksheet.addRow(['Status', 'Count']);
    pieChartLabels.forEach((label, index) => {
        chartWorksheet.addRow([label, pieChartValues[index]]);
    });

    chartWorksheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: balancedGreen },
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    chartWorksheet.eachRow((row, rowNumber) => {
        row.eachCell(cell => {
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
            };
            if (rowNumber > 1 && rowNumber % 2 === 0) {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: lightGreen },
                };
            }
        });
    });

    const fileName = 'jira_Report.xlsx';
    await workbook.xlsx.writeFile(fileName);

    console.log(`Jira issues exported to ${fileName}`);
}

exportToExcelWithChart();
