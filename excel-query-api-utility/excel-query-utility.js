const XLSX = require('xlsx');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

class ExcelQueryAPIUtility {
    constructor() {
        this.apiConfig = {
            url: 'https://platform.kore.ai/api/public/bot/st-432f6262-f1ce-5a64-86d8-bdd58f5479be/advancedSearch',
            headers: {
                'auth': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJhcHBJZCI6ImNzLTA2MGQ0MDgwLTk5MzgtNWM0Mi04MjQ1LWZmNDNlM2EyNTUwMyJ9.DhXZkC_cfG4c_Qj7r2HyCHaYiBfmRFP6mRTxKS9JZcM',
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'User-Agent': 'Excel-Query-API-Utility/1.0'
            }
        };
        this.testAppId = 'st-cf495d40-5aca-5bde-9d8b-3fa910dcf9af';
        this.delay = 1000; // 1 second delay between requests
    }

    /**
     * Read Excel file and extract queries
     */
    readExcelQueries(filePath, sheetName = null, queryColumn = 'query') {
        try {
            console.log(`Reading Excel file: ${filePath}`);
            
            const workbook = XLSX.readFile(filePath);
            const sheet = sheetName || workbook.SheetNames[0];
            console.log(`Processing sheet: ${sheet}`);
            
            const worksheet = workbook.Sheets[sheet];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`Found ${jsonData.length} rows in Excel file`);
            
            const queries = jsonData.map((row, index) => ({
                rowNumber: index + 2,
                query: row[queryColumn],
                originalRow: row
            })).filter(item => item.query && item.query.trim() !== '');
            
            console.log(`Extracted ${queries.length} valid queries`);
            return queries;
            
        } catch (error) {
            console.error('Error reading Excel file:', error.message);
            throw error;
        }
    }

    /**
     * Call the API for a single query
     */
    async callAPI(query) {
        try {
            console.log(`Calling API for query: "${query}"`);
            
            const response = await axios.post(this.apiConfig.url, {
                query: query
            }, {
                headers: this.apiConfig.headers,
                timeout: 30000
            });
            
            console.log(`✓ API call successful for query: "${query}"`);
            return {
                success: true,
                data: response.data,
                status: response.status,
                query: query
            };
            
        } catch (error) {
            console.error(`✗ API call failed for query: "${query}"`, error.message);
            return {
                success: false,
                error: error.message,
                status: error.response?.status || 'NETWORK_ERROR',
                query: query
            };
        }
    }

    /**
     * Process all queries sequentially
     */
    async processQueries(queries) {
        const results = [];
        
        console.log(`\nStarting sequential processing of ${queries.length} queries...`);
        console.log(`Test App ID: ${this.testAppId}`);
        console.log(`Delay between requests: ${this.delay}ms\n`);
        
        for (let i = 0; i < queries.length; i++) {
            const queryObj = queries[i];
            
            console.log(`[${i + 1}/${queries.length}] Processing row ${queryObj.rowNumber}`);
            
            const apiResult = await this.callAPI(queryObj.query);
            
            const result = {
                rowNumber: queryObj.rowNumber,
                query: queryObj.query,
                originalRow: queryObj.originalRow,
                apiResponse: apiResult,
                timestamp: new Date().toISOString(),
                testAppId: this.testAppId
            };
            
            results.push(result);
            
            if (i < queries.length - 1) {
                console.log(`Waiting ${this.delay}ms before next request...\n`);
                await new Promise(resolve => setTimeout(resolve, this.delay));
            }
        }
        
        console.log(`\nCompleted processing all ${queries.length} queries`);
        return results;
    }

    /**
     * Save results to JSON file
     */
    saveResults(results, outputPath) {
        try {
            const outputData = {
                metadata: {
                    totalQueries: results.length,
                    successfulCalls: results.filter(r => r.apiResponse.success).length,
                    failedCalls: results.filter(r => !r.apiResponse.success).length,
                    testAppId: this.testAppId,
                    processedAt: new Date().toISOString()
                },
                results: results
            };
            
            fs.writeFileSync(outputPath, JSON.stringify(outputData, null, 2));
            console.log(`Results saved to: ${outputPath}`);
            
        } catch (error) {
            console.error('Error saving results:', error.message);
            throw error;
        }
    }

    /**
     * Create Excel report
     */
    createExcelReport(results, outputPath) {
        try {
            const reportData = results.map(result => {
                // Truncate response data to fit Excel limits
                let responseData = '';
                if (result.apiResponse.success && result.apiResponse.data) {
                    responseData = JSON.stringify(result.apiResponse.data);
                    if (responseData.length > 30000) {
                        responseData = responseData.substring(0, 30000) + '... [TRUNCATED]';
                    }
                }
                
                return {
                    'Row Number': result.rowNumber,
                    'Query': result.query,
                    'API Success': result.apiResponse.success ? 'YES' : 'NO',
                    'HTTP Status': result.apiResponse.status,
                    'Error Message': (result.apiResponse.error || '').substring(0, 1000),
                    'Response Data': responseData,
                    'Timestamp': result.timestamp,
                    'Test App ID': result.testAppId
                };
            });
            
            const worksheet = XLSX.utils.json_to_sheet(reportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'API Results');
            
            XLSX.writeFile(workbook, outputPath);
            console.log(`Excel report saved to: ${outputPath}`);
            
        } catch (error) {
            console.error('Error creating Excel report:', error.message);
            throw error;
        }
    }

    /**
     * Main execution function
     */
    async execute(inputExcelPath, options = {}) {
        const {
            sheetName = null,
            queryColumn = 'query',
            outputDir = './output',
            delayMs = 1000
        } = options;
        
        this.delay = delayMs;
        
        try {
            if (!fs.existsSync(outputDir)) {
                fs.mkdirSync(outputDir, { recursive: true });
            }
            
            const queries = this.readExcelQueries(inputExcelPath, sheetName, queryColumn);
            
            if (queries.length === 0) {
                console.log('No queries found in the Excel file');
                return;
            }
            
            const results = await this.processQueries(queries);
            
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
            const jsonOutputPath = path.join(outputDir, `api_results_${timestamp}.json`);
            const excelOutputPath = path.join(outputDir, `api_results_${timestamp}.xlsx`);
            
            this.saveResults(results, jsonOutputPath);
            this.createExcelReport(results, excelOutputPath);
            
            const successful = results.filter(r => r.apiResponse.success).length;
            const failed = results.filter(r => !r.apiResponse.success).length;
            
            console.log('\n=== EXECUTION SUMMARY ===');
            console.log(`Total queries processed: ${results.length}`);
            console.log(`Successful API calls: ${successful}`);
            console.log(`Failed API calls: ${failed}`);
            console.log(`Success rate: ${((successful / results.length) * 100).toFixed(1)}%`);
            console.log(`Test App ID: ${this.testAppId}`);
            console.log(`JSON results: ${jsonOutputPath}`);
            console.log(`Excel report: ${excelOutputPath}`);
            
        } catch (error) {
            console.error('Execution failed:', error.message);
            throw error;
        }
    }
}

async function main() {
    const utility = new ExcelQueryAPIUtility();
    
    const inputFile = './input/queries.xlsx';
    const options = {
        sheetName: null,
        queryColumn: 'query',
        outputDir: './output',
        delayMs: 1000
    };
    
    try {
        await utility.execute(inputFile, options);
    } catch (error) {
        console.error('Application failed:', error.message);
        process.exit(1);
    }
}

if (require.main === module) {
    main();
}

module.exports = ExcelQueryAPIUtility;
