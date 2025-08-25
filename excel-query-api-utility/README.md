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
