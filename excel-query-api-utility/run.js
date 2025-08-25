const ExcelQueryAPIUtility = require('./excel-query-utility');
const path = require('path');
const fs = require('fs');

async function runUtility() {
    const utility = new ExcelQueryAPIUtility();
    
    // Configuration
    const inputFile = './input/queries.xlsx';
    const options = {
        sheetName: null,        // Use first sheet
        queryColumn: 'query',   // Column name with queries
        outputDir: './output',  // Output directory
        delayMs: 1000          // 1 second delay between API calls
    };
    
    try {
        console.log('🚀 Starting Excel Query API Utility...');
        console.log(`📁 Input file: ${path.resolve(inputFile)}`);
        console.log(`⚙️  Configuration:`, options);
        console.log('─'.repeat(50));
        
        // Check if input file exists
        if (!fs.existsSync(inputFile)) {
            console.error(`❌ Input file not found: ${inputFile}`);
            console.log('\n📋 Please create an Excel file with the following structure:');
            console.log('   | query |');
            console.log('   |-------|');
            console.log('   | how to remove admin user |');
            console.log('   | reset password procedure |');
            console.log('   | create new workspace |');
            console.log('\n   Save it as: ./input/queries.xlsx');
            return;
        }
        
        await utility.execute(inputFile, options);
        
        console.log('─'.repeat(50));
        console.log('✅ Utility completed successfully!');
        console.log('\n📂 Check the output folder for results:');
        console.log('   - JSON file with detailed API responses');
        console.log('   - Excel file with summary report');
        
    } catch (error) {
        console.error('❌ Error:', error.message);
        process.exit(1);
    }
}

// Run the utility
runUtility();
