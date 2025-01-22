
const axios = require('axios');

const ExcelJS = require('exceljs');

const { createCanvas } = require('canvas');

const { Chart, registerables } = require('chart.js');

 

//dont touch this ------

Chart.register(...registerables);

 

// Set Jira credentials===





const JIRA_BASE_URL = "Domain_URL";

const JIRA_EMAIL = "Your_email";

const JIRA_API_TOKEN = "Your_Api_Token";

const Project_Key = "Project_Key";

 

async function fetchJiraIssues() {

    try {

        const response = await axios.get(

            `${JIRA_BASE_URL}/rest/api/3/search?jql=project=${Project_Key}`,

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

 

async function createPieChart(labels, data) {

    const width = 400;  

    const height = 300;  

    const canvas = createCanvas(width, height);

    const ctx = canvas.getContext('2d');

 

    const total = data.reduce((sum, value) => sum + value, 0);

    const percentages = data.map(value => ((value / total) * 100).toFixed(2) + '%');

 

    const configuration = {

        type: 'pie',

        data: {

            labels: labels.map((label, index) => `${label} (${data[index]} - ${percentages[index]})`),

            datasets: [{

                label: 'Issue Types',

                data: data,

                backgroundColor: [

                    'rgba(255, 99, 133, 0.4)',

                    'rgba(54, 163, 235, 0.42)',

                    'rgba(255, 207, 86, 0.51)',

                    'rgba(75, 192, 192, 0.37)',

                    'rgba(153, 102, 255, 0.33)',

                    'rgba(255, 160, 64, 0.34)'

                ],

                borderColor: [

                    'rgba(255, 99, 132, 1)',

                    'rgba(54, 162, 235, 1)',

                    'rgba(255, 206, 86, 1)',

                    'rgba(75, 192, 192, 1)',

                    'rgba(153, 102, 255, 1)',

                    'rgba(255, 159, 64, 1)'

                ],

                borderWidth: 1

            }]

        },

        options: {

            responsive: true,

            plugins: {

                legend: {

                    position: 'top',

                },

                title: {

                    display: true,

                    text: 'Issue Types Distribution'

                }

            }

        }

    };

 

    new Chart(ctx, configuration);

    return canvas.toBuffer();

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

        Priority: issue.fields.priority ? issue.fields.priority.name : 'None',

        Labels: issue.fields.labels.join(', ') || 'None',

        Assignee: issue.fields.assignee ? issue.fields.assignee.displayName : 'Unassigned',

        Reporter: issue.fields.reporter ? issue.fields.reporter.displayName : 'Unknown',

        Created: issue.fields.created || 'N/A',

        FixVersions: issue.fields.fixVersions.map(fv => fv.name).join(', ') || 'None',

        IssueType: issue.fields.issuetype ? issue.fields.issuetype.name : 'Unknown',

    }));

 

console.log(data);




    const statusCounts = data.reduce((counts, issue) => {

        counts[issue.Status] = (counts[issue.Status] || 0) + 1;

        return counts;

    }, {});

 

console.log(statusCounts);

 

    const priorityCounts = data.reduce((counts, issue) => {

        counts[issue.Priority] = (counts[issue.Priority] || 0) + 1;

        return counts;

    }, {});

 

console.log(priorityCounts);

 

    const issueTypeCounts = data.reduce((counts, issue) => {

        counts[issue.IssueType] = (counts[issue.IssueType] || 0) + 1;

        return counts;

    }, {});

    console.log(issueTypeCounts);

   

 

    const pieChartLabels = Object.keys(statusCounts);

    const pieChartValues = Object.values(statusCounts);

 

    const priorityChartLabels = Object.keys(priorityCounts);

    const priorityChartValues = Object.values(priorityCounts);

 

    const issueTypeChartLabels = Object.keys(issueTypeCounts);

    const issueTypeChartValues = Object.values(issueTypeCounts);

 

    const workbook = new ExcelJS.Workbook();

    const worksheet = workbook.addWorksheet('Jira Issues');

 

    worksheet.columns = [

        { header: 'Key', key: 'Key', width: 15 },

        { header: 'Summary', key: 'Summary', width: 50 },

        { header: 'Status', key: 'Status', width: 15 },

        { header: 'Assignee', key: 'Assignee', width: 20 },

        { header: 'Priority', key: 'Priority', width: 20 },

        { header: 'Reporter', key: 'Reporter', width: 20 },

        { header: 'Fix Versions', key: 'FixVersions', width: 25 },

        { header: 'Labels', key: 'Labels', width: 15 },

        { header: 'Issue Type', key: 'IssueType', width: 15 },

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

 

    const chartWorksheet = workbook.addWorksheet('Charts');

    chartWorksheet.addRow(['Status', 'Count']);

    pieChartLabels.forEach((label, index) => {

        chartWorksheet.addRow([label, pieChartValues[index]]);

    });

 

    chartWorksheet.addRow([]);

    chartWorksheet.addRow(['Priority', 'Count']);

    priorityChartLabels.forEach((label, index) => {

        chartWorksheet.addRow([label, priorityChartValues[index]]);

    });

 

    chartWorksheet.addRow([]);

    chartWorksheet.addRow(['Issue Type', 'Count']);

    issueTypeChartLabels.forEach((label, index) => {

        chartWorksheet.addRow([label, issueTypeChartValues[index]]);

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

 

    const summaryWorksheet = workbook.addWorksheet('Summary');

    summaryWorksheet.addRow(['Issue Type', 'Count']);

    issueTypeChartLabels.forEach((label, index) => {

        summaryWorksheet.addRow([label, issueTypeChartValues[index]]);

    });

 

    summaryWorksheet.getRow(1).eachCell(cell => {

        cell.font = { bold: true, color: { argb: 'FFFFFF' } };

        cell.fill = {

            type: 'pattern',

            pattern: 'solid',

            fgColor: { argb: balancedGreen },

        };

        cell.alignment = { horizontal: 'center', vertical: 'middle' };

    });

 

    summaryWorksheet.eachRow((row, rowNumber) => {

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

 

    const pieChartImage = await createPieChart(issueTypeChartLabels, issueTypeChartValues);

    const imageId = workbook.addImage({

        buffer: pieChartImage,

        extension: 'png',

    });

 

    chartWorksheet.addImage(imageId, {

        tl: { col: 3, row: issueTypeChartLabels.length + 3 },

        ext: { width: 400, height: 400 },

    });

 

    const fileName = 'jira_Report.xlsx';

    await workbook.xlsx.writeFile(fileName);

 

    console.log(`Jira issues exported to ${fileName}`);

}

 

exportToExcelWithChart();