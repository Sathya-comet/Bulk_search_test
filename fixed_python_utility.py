import requests
import pandas as pd
import time
import json
import argparse
from datetime import datetime
import os
import re
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

def perform_api_request(query):
    """
    Send a query to the Answer Generation API and return the response
    """
    url = "https://platform.kore.ai/api/1.1/builder/streams/st-cf495d40-5aca-5bde-9d8b-3fa910dcf9af/answer"
    
    # Get authentication from environment variables
    account_id = os.getenv('KORE_ACCOUNT_ID', '65eb1114c6a81530c0e3232c')
    bearer_token = os.getenv('KORE_BEARER_TOKEN', 'dt6csksSqU8PjKGXT09t-hTf3azPFmfvG5OmMr3hn6_oiVeG56K5HyI-JGEqN3Js')
    
    # Debug: Print token info (first and last 10 chars only for security)
    print(f"  Using token: {bearer_token[:10]}...{bearer_token[-10:]}")
    
    headers = {
        'Accountid': account_id,
        'Authorization': f'bearer {bearer_token}',
        'sec-ch-ua-platform': '"macOS"',
        'Referer': 'https://platform.kore.ai/builder/app/answergeneration',
        'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
        'App-Language': 'en',
        'sec-ch-ua-mobile': '?0',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
        'Accept': 'application/json, text/plain, */*',
        'state': 'configured',
        'Content-Type': 'application/json;charset=UTF-8'
    }
    
    payload = {
        "query": query
    }
    
    try:
        print(f"  Calling Answer Generation API for: '{query[:50]}...'")
        print(f"  URL: {url}")
        print(f"  Account ID: {account_id}")
        
        response = requests.post(url, headers=headers, json=payload, timeout=30)
        
        # Debug: Print response status
        print(f"  HTTP Status: {response.status_code}")
        
        if response.status_code == 401:
            print(f"  Response headers: {dict(response.headers)}")
            print(f"  Response text: {response.text[:500]}")
            
        response.raise_for_status()
        print(f"  API call successful")
        return {
            "success": True,
            "data": response.json(),
            "status_code": response.status_code
        }
    except requests.exceptions.Timeout:
        print(f"  API call timed out")
        return {
            "success": False,
            "error": "Request timed out after 30 seconds",
            "status_code": "TIMEOUT"
        }
    except requests.exceptions.RequestException as e:
        print(f"  API call failed: {str(e)}")
        if hasattr(e, 'response') and e.response is not None:
            print(f"  Response status: {e.response.status_code}")
            print(f"  Response text: {e.response.text[:500]}")
        return {
            "success": False,
            "error": str(e),
            "status_code": getattr(e.response, 'status_code', 'NETWORK_ERROR')
        }

def extract_answer_and_titles(response):
    """
    Extract the answer text and source document titles from the API response
    """
    try:
        if not response.get("success", False):
            return response.get("error", "Unknown error"), ""
        
        data = response.get("data", {})
        answer = ""
        titles = []
        
        # Extract answer from response.answer field
        if "response" in data and "answer" in data["response"]:
            answer = data["response"]["answer"]
        elif "answer" in data:
            answer = data["answer"]
        else:
            answer = "No answer found in response"
        
        # Extract document titles from source citations
        try:
            if "response" in data and "answer_payload" in data["response"]:
                center_panel = data["response"]["answer_payload"].get("center_panel", {})
                if "data" in center_panel:
                    for item in center_panel["data"]:
                        if "snippet_content" in item:
                            for content in item["snippet_content"]:
                                if "sources" in content:
                                    for source in content["sources"]:
                                        title = source.get("title", "")
                                        if title and title not in titles:
                                            titles.append(title)
        except Exception as e:
            print(f"  Warning: Could not extract titles - {e}")
        
        # Join titles with semicolon separator
        titles_string = "; ".join(titles) if titles else "No titles found"
        
        return answer, titles_string
            
    except Exception as e:
        return f"Error extracting answer: {str(e)}", ""

def compare_sources(expected_source, response_titles):
    """
    Compare expected source with response document titles using contains logic
    Returns 'YES' if any response title contains the expected document name, 'NO' otherwise
    """
    if not expected_source or expected_source.strip() == "" or expected_source == "No titles found":
        return "N/A"
    
    if not response_titles or response_titles == "No titles found":
        return "NO"
    
    # Extract document name from expected source (remove page numbers and extra info)
    expected_clean = expected_source.lower().strip()
    
    # Remove page number references like "(pp. 5-7)", "(p. 11)", etc.
    expected_clean = re.sub(r'\s*\([p\.]*\s*\d+[-–]*\d*\)', '', expected_clean)
    expected_clean = expected_clean.strip()
    
    # Extract key document identifiers
    doc_keywords = []
    if "kore.ai" in expected_clean or "kore" in expected_clean:
        doc_keywords.append("kore")
    if "handbook" in expected_clean:
        doc_keywords.append("handbook")
    if "us" in expected_clean:
        doc_keywords.append("us")
    if "uk" in expected_clean:
        doc_keywords.append("uk")
    if "staff" in expected_clean:
        doc_keywords.append("staff")
    
    # Split response titles and check each one
    title_list = [title.strip() for title in response_titles.split(';') if title.strip()]
    
    for title in title_list:
        title_clean = title.lower().strip()
        
        # Check if key document keywords appear in the response title
        if len(doc_keywords) >= 2:
            matches = sum(1 for keyword in doc_keywords if keyword in title_clean)
            if matches >= 2:
                return "YES"
        
        # Also check for direct filename matches (without page numbers)
        if expected_clean.replace(".pdf", "") in title_clean:
            return "YES"
        
        # Check if any significant part of expected source appears in title
        expected_parts = expected_clean.replace(".pdf", "").split("-")
        if len(expected_parts) > 1:
            for part in expected_parts:
                if len(part) > 3 and part in title_clean:
                    return "YES"
    
    return "NO"

def batch_process(input_file, output_file, delay=1):
    """
    Process queries from input Excel file and write results to output Excel file
    """
    try:
        if not os.path.exists(input_file):
            print(f"Error: Input file '{input_file}' not found")
            return False
        
        print(f"Reading queries from: {input_file}")
        
        # Read Excel file with proper engine
        if input_file.endswith('.xlsx'):
            df = pd.read_excel(input_file, engine='openpyxl')
        elif input_file.endswith('.csv'):
            df = pd.read_csv(input_file)
        else:
            print("Error: Input file must be .xlsx or .csv format")
            return False
        
        # Check for required columns (case insensitive)
        query_column = None
        source_column = None
        
        for col in df.columns:
            if col.lower() in ['query', 'queries']:
                query_column = col
            elif col.lower() in ['source_url', 'source_urls', 'sourceurl', 'source']:
                source_column = col
        
        if query_column is None:
            print("Error: Input file must contain a 'query' column")
            print(f"Found columns: {list(df.columns)}")
            return False
        
        print(f"Found query column: '{query_column}'")
        
        if source_column is None:
            print("Warning: No 'source_url' column found. Source comparison will be skipped.")
            print(f"Found columns: {list(df.columns)}")
        else:
            print(f"Found source column: '{source_column}'")
        
        # Filter out empty queries
        df_filtered = df[df[query_column].notna() & (df[query_column] != "")]
        
        if len(df_filtered) == 0:
            print("Error: No valid queries found in the input file")
            return False
        
        # Add results columns
        df_filtered = df_filtered.copy()
        df_filtered['API_Success'] = None
        df_filtered['Status_Code'] = None
        df_filtered['Answer'] = None
        df_filtered['Response_Titles'] = None
        df_filtered['Source_Match'] = None
        df_filtered['Timestamp'] = None
        df_filtered['Processing_Time_Seconds'] = None
        
        print(f"\nProcessing {len(df_filtered)} queries with a {delay}-second delay between each...")
        print("Using Answer Generation API endpoint")
        print("=" * 70)
        
        # Test environment variables
        account_id = os.getenv('KORE_ACCOUNT_ID')
        bearer_token = os.getenv('KORE_BEARER_TOKEN')
        print(f"Environment check - Account ID: {account_id}")
        print(f"Environment check - Token: {'Found' if bearer_token else 'NOT FOUND'}")
        print("=" * 70)
        
        # Process each query
        for index, row in df_filtered.iterrows():
            query = row[query_column]
            expected_source = row.get(source_column, "") if source_column else ""
            query_num = index + 1
            
            print(f"\n[{query_num}/{len(df_filtered)}] Processing: {query}")
            if expected_source:
                print(f"  Expected source: {expected_source}")
            
            start_time = time.time()
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Make API request
            response = perform_api_request(query)
            
            # Extract answer and titles
            answer, titles = extract_answer_and_titles(response)
            processing_time = round(time.time() - start_time, 2)
            
            # Compare sources if source column exists
            source_match = "N/A"
            if source_column and expected_source:
                source_match = compare_sources(expected_source, titles)
            
            # Create answer preview (first 150 chars for console display)
            if answer and len(answer) > 150:
                answer_preview = answer[:150] + "..."
            else:
                answer_preview = answer
            
            # Update DataFrame with results
            df_filtered.at[index, 'API_Success'] = "YES" if response.get("success") else "NO"
            df_filtered.at[index, 'Status_Code'] = response.get("status_code", "Unknown")
            df_filtered.at[index, 'Answer'] = answer
            df_filtered.at[index, 'Response_Titles'] = titles
            df_filtered.at[index, 'Source_Match'] = source_match
            df_filtered.at[index, 'Timestamp'] = timestamp
            df_filtered.at[index, 'Processing_Time_Seconds'] = processing_time
            
            print(f"  Status: {'SUCCESS' if response.get('success') else 'FAILED'}")
            if response.get("success") and answer:
                print(f"  Answer preview: {answer_preview}")
                title_count = len(titles.split('; ')) if titles != "No titles found" else 0
                print(f"  Response titles found: {title_count}")
                if source_column:
                    print(f"  Source Match: {source_match}")
            print(f"  Processing time: {processing_time}s")
            
            # Wait for the specified delay before the next query
            if index < len(df_filtered) - 1:
                print(f"  Waiting {delay} seconds before next query...")
                time.sleep(delay)
        
        # Write results to output Excel file
        print(f"\n" + "=" * 70)
        
        # Create output directory if it doesn't exist
        os.makedirs("output", exist_ok=True)
        
        print(f"Writing results to: {output_file}")
        df_filtered.to_excel(output_file, index=False, engine='openpyxl')
        
        # Print summary
        successful = len(df_filtered[df_filtered['API_Success'] == 'YES'])
        failed = len(df_filtered[df_filtered['API_Success'] == 'NO'])
        
        # Source matching summary
        source_matches = 0
        source_no_matches = 0
        if source_column:
            source_matches = len(df_filtered[df_filtered['Source_Match'] == 'YES'])
            source_no_matches = len(df_filtered[df_filtered['Source_Match'] == 'NO'])
        
        print(f"\nPROCESSING SUMMARY:")
        print(f"   API Endpoint: Answer Generation")
        print(f"   Test App ID: st-cf495d40-5aca-5bde-9d8b-3fa910dcf9af")
        print(f"   Total queries processed: {len(df_filtered)}")
        print(f"   Successful API calls: {successful}")
        print(f"   Failed API calls: {failed}")
        print(f"   Success rate: {(successful/len(df_filtered)*100):.1f}%")
        
        if source_column:
            print(f"   Source Matches (YES): {source_matches}")
            print(f"   Source Non-matches (NO): {source_no_matches}")
            if (source_matches + source_no_matches) > 0:
                print(f"   Source Match rate: {(source_matches/(source_matches+source_no_matches)*100):.1f}%")
        
        print(f"   Results saved to: {output_file}")
        
        if failed > 0:
            print(f"\n   Failed queries:")
            failed_queries = df_filtered[df_filtered['API_Success'] == 'NO']
            for _, row in failed_queries.iterrows():
                print(f"     - \"{row[query_column]}\" ({row['Status_Code']})")
        
        return True
    
    except Exception as e:
        print(f"Error during batch processing: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def create_sample_input():
    """
    Create a sample input Excel file for testing
    """
    sample_data = [
        {"query": "What is the company's code of conduct?", "source": "Kore.ai-US-Handbook.pdf (pp. 5-7)"},
        {"query": "How do I report harassment?", "source": "Kore.ai-US-Handbook.pdf (pp. 12-14)"},
        {"query": "What are equal opportunity policies?", "source": "Kore.ai-US-Handbook.pdf (pp. 8-10)"},
        {"query": "Outside employment policy?", "source": "Kore.ai-US-Handbook.pdf (p. 11)"},
        {"query": "Confidentiality rules?", "source": "Kore.ai-US-Handbook.pdf (pp. 15-16)"}
    ]
    
    df = pd.DataFrame(sample_data)
    
    # Create input directory if it doesn't exist
    os.makedirs("input", exist_ok=True)
    
    filename = "input/sample_queries.xlsx"
    df.to_excel(filename, index=False, engine='openpyxl')
    print(f"Sample input file created: {filename}")
    return filename

def main():
    parser = argparse.ArgumentParser(description='Answer Generation API Testing Utility')
    parser.add_argument('--input', '-i', help='Input Excel file with queries (e.g., input/queries.xlsx)')
    parser.add_argument('--output', '-o', help='Output Excel file for results')
    parser.add_argument('--delay', '-d', type=int, default=1, help='Delay between queries in seconds (default: 1)')
    parser.add_argument('--create-sample', action='store_true', help='Create a sample input file')
    
    args = parser.parse_args()
    
    # Create sample file if requested
    if args.create_sample:
        create_sample_input()
        print("\nTo run the utility with the sample file:")
        print("python fixed_python_utility.py --input input/sample_queries.xlsx")
        return
    
    # Check for input file
    if not args.input:
        print("Error: Please provide an input file with --input or use --create-sample to create one")
        print("\nUsage examples:")
        print("  python fixed_python_utility.py --create-sample")
        print("  python fixed_python_utility.py --input input/queries.xlsx")
        print("  python fixed_python_utility.py --input input/queries.xlsx --output output/results.xlsx --delay 2")
        return
    
    # Generate default output filename if not provided
    output_file = args.output
    if not output_file:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"output/answer_results_{timestamp}.xlsx"
    
    print("Kore.ai Answer Generation API Testing Utility")
    print("=" * 55)
    
    success = batch_process(args.input, output_file, args.delay)
    
    if success:
        print("\nBatch processing completed!")
    else:
        print("\nBatch processing failed!")

if __name__ == "__main__":
    main()