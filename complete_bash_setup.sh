#!/bin/bash

# Excel Query API Utility - Complete Setup and Run Script
# For Mac/Linux systems

set -e  # Exit on any error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Configuration
PROJECT_NAME="excel-query-api-utility"
CURRENT_DIR=$(pwd)
PROJECT_DIR="$CURRENT_DIR/$PROJECT_NAME"

echo -e "${BLUE}🚀 Excel Query API Utility Setup Script${NC}"
echo -e "${BLUE}======================================${NC}\n"

# Function to print colored output
print_status() {
    echo -e "${GREEN}✅ $1${NC}"
}

print_warning() {
    echo -e "${YELLOW}⚠️  $1${NC}"
}

print_error() {
    echo -e "${RED}❌ $1${NC}"
}

print_info() {
    echo -e "${BLUE}ℹ️  $1${NC}"
}

# Check if Node.js is installed
check_nodejs() {
    if ! command -v node &> /dev/null; then
        print_error "Node.js is not installed!"
        echo "Please install Node.js from https://nodejs.org/"
        echo "Or use Homebrew: brew install node"
        exit 1
    fi
    
    NODE_VERSION=$(node --version)
    print_status "Node.js is installed: $NODE_VERSION"
}

# Check if npm is installed
check_npm() {
    if ! command -v npm &> /dev/null; then
        print_error "npm is not installed!"
        exit 1
    fi
    
    NPM_VERSION=$(npm --version)
    print_status "npm is installed: $NPM_VERSION"
}

# Create project structure
create_project_structure() {
    print_info "Creating project structure..."
    
    # Create main project directory
    mkdir -p "$PROJECT_DIR"
    cd "$PROJECT_DIR"
    
    # Create subdirectories
    mkdir -p input output
    
    print_status "Project structure created at: $PROJECT_DIR"
}

# Create package.json
create_package_json() {
    print_info "Creating package.json..."
    
    cat > package.json << 'EOF'
{
  "name": "excel-query-api-utility",
  "version": "1.0.0",
  "description": "Utility to process Excel queries through Kore.ai API",
  "main": "excel-query-utility.js",
  "scripts": {
    "start": "node run.js",
    "test": "node excel-query-utility.js",
    "dev": "node --inspect run.js"
  },
  "keywords": ["excel", "api", "kore.ai", "automation", "testing"],
  "author": "API Testing Team",
  "license": "MIT",
  "dependencies": {
    "xlsx": "^0.18.5",
    "axios": "^1.6.0"
  }
}
EOF
    
    print_status "package.json created"
}

# Create the main utility file
create_utility_file() {
    print_info "Creating excel-query-utility.js..."
    
    cat > excel-query-utility.js << 'EOF'
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
            const reportData = results.map(result => ({
                'Row Number': result.rowNumber,
                'Query': result.query,
                'API Success': result.apiResponse.success ? 'YES' : 'NO',
                'HTTP Status': result.apiResponse.status,
                'Error Message': result.apiResponse.error || '',
                'Response Data': result.apiResponse.success ? JSON.stringify(result.apiResponse.data) : '',
                'Timestamp': result.timestamp,
                'Test App ID': result.testAppId
            }));
            
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
EOF
    
    print_status "excel-query-utility.js created"
}

# Create runner script
create_runner_script() {
    print_info "Creating run.js..."
    
    cat > run.js << 'EOF'
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
EOF
    
    print_status "run.js created"
}

# Create sample Excel file
create_sample_excel() {
    print_info "Creating sample Excel file..."
    
    cat > create_sample_excel.js << 'EOF'
const XLSX = require('xlsx');

// Sample queries for testing
const sampleQueries = [
    { query: "how to remove admin user" },
    { query: "reset password procedure" },
    { query: "create new workspace" },
    { query: "delete user account" },
    { query: "configure SSO settings" },
    { query: "manage user permissions" },
    { query: "backup data export" },
    { query: "restore system settings" }
];

// Create workbook and worksheet
const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(sampleQueries);

// Add worksheet to workbook
XLSX.utils.book_append_sheet(workbook, worksheet, 'Queries');

// Write to file
XLSX.writeFile(workbook, './input/queries.xlsx');

console.log('✅ Sample Excel file created: ./input/queries.xlsx');
EOF
    
    print_status "Sample Excel creator script created"
}

# Install dependencies
install_dependencies() {
    print_info "Installing Node.js dependencies..."
    
    npm install
    
    print_status "Dependencies installed successfully"
}

# Create sample Excel file
generate_sample_excel() {
    print_info "Generating sample Excel file..."
    
    node create_sample_excel.js
    rm create_sample_excel.js  # Clean up the temporary script
    
    print_status "Sample Excel file created"
}

# Create README file
create_readme() {
    print_info "Creating README.md..."
    
    cat > README.md << 'EOF'
# Excel Query API Utility

A Node.js utility to process Excel files containing search queries and test them against the Kore.ai API sequentially.

## Quick Start

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Prepare your Excel file:**
   - Place your Excel file in the `input/` folder
   - Name it `queries.xlsx`
   - Ensure it has a column named "query"

3. **Run the utility:**
   ```bash
   npm start
   ```

## Project Structure

```
excel-query-api-utility/
├── excel-query-utility.js    # Main utility
├── run.js                    # Runner script
├── package.json              # Dependencies
├── input/
│   └── queries.xlsx         # Your Excel file
└── output/                  # Results (auto-generated)
```

## Excel Format

Your Excel file should have at least one column named "query":

| query |
|-------|
| how to remove admin user |
| reset password procedure |
| create new workspace |

## Output

The utility generates:
- **JSON file**: Detailed API responses
- **Excel report**: Summary with success/failure status

## Configuration

Edit `run.js` to customize:
- Input file path
- Output directory
- Delay between API calls
- Column name for queries

## API Details

- **Endpoint**: Kore.ai Advanced Search API
- **Authentication**: JWT token (pre-configured)
- **Test App ID**: st-cf495d40-5aca-5bde-9d8b-3fa910dcf9af
EOF
    
    print_status "README.md created"
}

# Main execution
main() {
    echo "Starting setup in: $CURRENT_DIR"
    
    # Check prerequisites
    check_nodejs
    check_npm
    
    # Create project
    create_project_structure
    create_package_json
    create_utility_file
    create_runner_script
    create_sample_excel
    create_readme
    
    # Install and setup
    install_dependencies
    generate_sample_excel
    
    echo ""
    print_status "Setup completed successfully!"
    echo ""
    print_info "Project created at: $PROJECT_DIR"
    print_info "Sample Excel file created with test queries"
    echo ""
    echo -e "${YELLOW}Next steps:${NC}"
    echo "1. cd $PROJECT_NAME"
    echo "2. Edit input/queries.xlsx with your test queries"
    echo "3. npm start"
    echo ""
    echo -e "${GREEN}🎉 Ready to test your Kore.ai API queries!${NC}"
}

# Run the script
main
EOF

I've created a complete bash script that will set up everything for you automatically! Here's what this script does:

## 🚀 Complete Setup Script Features:

1. **Checks prerequisites** (Node.js, npm)
2. **Creates project structure** with all folders
3. **Generates all code files** (utility, runner, package.json)
4. **Installs dependencies** automatically
5. **Creates sample Excel file** with test queries
6. **Sets up README** documentation

## 📋 How to Use:

1. **Save the script** as `setup.sh` in your "Bulk search test" folder
2. **Make it executable**:
   ```bash
   chmod +x setup.sh
   ```
3. **Run the script**:
   ```bash
   ./setup.sh
   ```

## 🎯 What You Get:

After running the script, you'll have:
```
excel-query-api-utility/
├── excel-query-utility.js    # Main utility
├── run.js                    # Runner script  
├── package.json              # Dependencies
├── input/
│   └── queries.xlsx         # Sample Excel with test queries
└── output/                  # Results folder (auto-created)
```

## 🏃‍♂️ Quick Run Commands:

```bash
cd excel-query-api-utility
npm start                     # Run with sample queries
```

The script includes **error checking**, **colored output**, and **automatic sample data generation**. It's a complete one-click setup for your Kore.ai API testing environment!