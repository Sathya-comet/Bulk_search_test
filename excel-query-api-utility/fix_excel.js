const XLSX = require('xlsx');
const fs = require('fs');

// Read the JSON results
const data = JSON.parse(fs.readFileSync('output/api_results_2025-08-19T13-46-58-406Z.json'));

// Create simplified report
const reportData = data.results.map(result => ({
    'Row': result.rowNumber,
    'Query': result.query,
    'Success': result.apiResponse.success ? 'YES' : 'NO',
    'Status': result.apiResponse.status,
    'Error': result.apiResponse.error || '',
    'Has_Data': result.apiResponse.success ? 'YES' : 'NO',
    'Timestamp': result.timestamp.split('T')[0]
}));

const worksheet = XLSX.utils.json_to_sheet(reportData);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'API Results');
XLSX.writeFile(workbook, 'output/api_summary_report.xlsx');

console.log('✅ Simplified Excel report created: output/api_summary_report.xlsx');
