#!/usr/bin/env python3
"""Fetch company information with OpenAI GPT-5 using web search.

This script reads company names from an Excel file, uses OpenAI GPT-5 with web search
to fetch company revenue, CEO name, and CEO bio, and adds the results to new columns.
"""

import argparse
import json
import logging
import os
import sys
import time
from datetime import datetime
from typing import Dict, Optional

import pandas as pd


def load_env_file(env_path: Optional[str] = None) -> Dict[str, str]:
    """Load environment variables from a .env file.
    
    Args:
        env_path: Path to the .env file. If None, looks for .env in the repository root.
        
    Returns:
        Dictionary of environment variables.
    """
    if env_path is None:
        env_path = os.path.join(os.path.dirname(__file__), "..", ".env")
    
    env_vars: Dict[str, str] = {}
    
    if not os.path.exists(env_path):
        return env_vars
    
    try:
        with open(env_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    env_vars[key.strip()] = value.strip()
    except Exception:
        # Silently ignore errors reading .env file
        pass
    
    return env_vars


def get_openai_api_key() -> str:
    """Get the OpenAI API key from environment variables or .env file.
    
    Returns:
        The OpenAI API key, or empty string if not found.
    """
    # First check environment variables
    api_key = os.environ.get('OPENAI_API_KEY', '')
    if api_key:
        return api_key
    
    # Then check .env file
    env_vars = load_env_file()
    return env_vars.get('OPENAI_API_KEY', '')


def fetch_company_info(company_name: str, api_key: str) -> Dict[str, str]:
    """Fetch company information using OpenAI GPT-5 with web search.
    
    Args:
        company_name: Name of the company
        api_key: OpenAI API key
        
    Returns:
        Dictionary with fetched company information
    """
    if not company_name or not api_key:
        return {
            "AI_Revenue": "",
            "AI_CEO_Name": "",
            "AI_CEO_Bio": "",
            "AI_LinkedIn_URL": "",
            "AI_Error": "Missing company name or API key"
        }
    
    try:
        # Import OpenAI here to avoid dependency issues
        try:
            from openai import OpenAI
        except ImportError:
            return {
                "AI_Revenue": "",
                "AI_CEO_Name": "",
                "AI_CEO_Bio": "",
                "AI_LinkedIn_URL": "",
                "AI_Error": "OpenAI package not installed"
            }
        
        client = OpenAI(api_key=api_key)
        logging.debug(f"OpenAI client created for company info fetching")
        
        # Prepare the search query
        developer_prompt = """Your are financial analyst. User gives you companies one by one and your task is to find information: revenue, CEO name, CEO Bio, and LinkedIn profile URL
Look on finder.fi with company name, Linkedin for CEO Bio and profile URL. 
Response only JSON, no references, background data, nothing else."""

        user_prompt = f"""Find information defined in response JSON below.
{{
  "companyName": "{company_name}",
  "revenue": "X€",
  "ceoName": "N.N.",
  "ceoBioInLinkedin": "He is .....",
  "linkedInProfileUrl": "https://www.linkedin.com/in/ceo-name"
}}"""
        
        logging.info(f"Making OpenAI GPT-5 API call with web search for {company_name}")
        logging.debug(f"Developer prompt prepared")
        logging.debug(f"User prompt: {user_prompt[:200]}...")
        
        try:
            response = client.responses.create(
                model="gpt-5-mini",
                input=[
                    {
                        "role": "developer",
                        "content": [
                            {
                                "type": "input_text",
                                "text": developer_prompt
                            }
                        ]
                    },
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "input_text",
                                "text": user_prompt
                            }
                        ]
                    }
                ],
                text={
                    "format": {
                        "type": "text"
                    },
                    "verbosity": "low"
                },
                reasoning={
                    "effort": "low",
                    "summary": None
                },
                tools=[
                    {
                        "type": "web_search",
                        "user_location": {
                            "type": "approximate",
                            "country": "FI",
                            "city": "Helsinki"
                        },
                        "search_context_size": "low"
                    }
                ],
                store=False,
                include=[
                    "reasoning.encrypted_content"
                ]
            )
            logging.info(f"OpenAI GPT-5 API call successful for {company_name}")
        except Exception as api_exc:
            logging.error(f"OpenAI GPT-5 API call failed for {company_name}: {api_exc}")
            return {
                "AI_Revenue": "",
                "AI_CEO_Name": "",
                "AI_CEO_Bio": "",
                "AI_LinkedIn_URL": "",
                "AI_Error": f"API call failed: {api_exc}"
            }
        
        # Extract and parse the response
        response_text = ""
        if hasattr(response, 'output_text') and response.output_text:
            response_text = response.output_text
            logging.debug(f"Got output_text: {response_text[:200] if response_text else 'Empty response'}...")
        elif hasattr(response, 'text') and response.text:
            response_data = response.text
            if isinstance(response_data, str):
                response_text = response_data
                logging.debug(f"Got string response: {response_text[:200]}...")
            else:
                response_text = str(response_data)
                logging.debug(f"Converted response.text to string: {response_text[:200]}...")
        else:
            logging.error(f"Unable to extract text from response. Response type: {type(response)}")
            logging.error(f"Available attributes: {[attr for attr in dir(response) if not attr.startswith('_')]}")
            return {
                "AI_Revenue": "",
                "AI_CEO_Name": "",
                "AI_CEO_Bio": "",
                "AI_LinkedIn_URL": "",
                "AI_Error": "Cannot extract response text"
            }
            
        if not response_text:
            logging.error(f"Empty response text for {company_name}")
            return {
                "AI_Revenue": "",
                "AI_CEO_Name": "",
                "AI_CEO_Bio": "",
                "AI_LinkedIn_URL": "",
                "AI_Error": "Empty response"
            }
            
        logging.info(f"OpenAI response received for {company_name}: {len(response_text)} characters")
        logging.debug(f"Raw response: {response_text}")

        # Parse JSON response
        try:
            logging.info(f"Attempting to parse JSON response for {company_name}")
            data = json.loads(response_text)
            logging.info(f"JSON parsing successful for {company_name}")
            logging.debug(f"Parsed data keys: {list(data.keys())}")
            
            # Extract information from the structured response
            company_name_resp = data.get('companyName', company_name)
            revenue = data.get('revenue', '')
            ceo_name = data.get('ceoName', '')
            ceo_bio = data.get('ceoBioInLinkedin', '')
            linkedin_url = data.get('linkedInProfileUrl', '')
            
            # Clean up the data
            if revenue and revenue.lower() in ['x€', 'n.n.', 'unknown', 'not available']:
                revenue = ''
            if ceo_name and ceo_name.lower() in ['n.n.', 'unknown', 'not available']:
                ceo_name = ''
            if ceo_bio and ceo_bio.lower() in ['he is .....', 'not available', 'unknown']:
                ceo_bio = ''
            if linkedin_url and linkedin_url.lower() in ['https://www.linkedin.com/in/ceo-name', 'not available', 'unknown']:
                linkedin_url = ''
            
            return {
                "AI_Revenue": revenue,
                "AI_CEO_Name": ceo_name,
                "AI_CEO_Bio": ceo_bio,
                "AI_LinkedIn_URL": linkedin_url,
                "AI_Error": ""
            }
            
        except json.JSONDecodeError as json_err:
            # If JSON parsing fails, try to extract information from raw text
            logging.error(f"JSON parsing failed for {company_name}: {json_err}")
            logging.error(f"Raw response that failed to parse: {response_text}")
            
            # Try to extract basic information from non-JSON response
            revenue = ""
            ceo_name = ""
            ceo_bio = ""
            
            # Simple text extraction as fallback
            if "revenue" in response_text.lower() or "€" in response_text:
                lines = response_text.split('\n')
                for line in lines:
                    if "revenue" in line.lower() and ("€" in line or "million" in line.lower() or "million" in line.lower()):
                        revenue = line.strip()
                        break
            
            if "ceo" in response_text.lower():
                lines = response_text.split('\n')
                for line in lines:
                    if "ceo" in line.lower() and len(line.split()) < 10:  # Likely a name
                        ceo_name = line.strip()
                        break
            
            truncated_response = response_text[:100] + "..." if len(response_text) > 100 else response_text
            return {
                "AI_Revenue": revenue,
                "AI_CEO_Name": ceo_name,
                "AI_CEO_Bio": ceo_bio,
                "AI_LinkedIn_URL": "",
                "AI_Error": f"JSON parsing failed, extracted from text: {truncated_response}"
            }
            
    except Exception as exc:
        logging.error(f"Unexpected error fetching info for {company_name}: {exc}")
        logging.error(f"Exception type: {type(exc).__name__}")
        return {
            "AI_Revenue": "",
            "AI_CEO_Name": "",
            "AI_CEO_Bio": "",
            "AI_LinkedIn_URL": "",
            "AI_Error": str(exc)
        }


def generate_timestamp():
    """Generate timestamp in YYYYMMDD_HHMMSS format"""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def parse_args():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(
        description="Fetch company information with OpenAI GPT-5 using web search"
    )
    parser.add_argument(
        "filename",
        help="Input Excel filename (e.g., Startup_Finland_Digital_Mature.xlsx)"
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=2.0,
        help="Delay in seconds between OpenAI API calls (default: 2.0)"
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Print detailed progress and results"
    )
    parser.add_argument(
        "--start-row",
        type=int,
        default=1,
        help="Row to start processing from (default: 1)"
    )
    parser.add_argument(
        "--max-rows",
        type=int,
        help="Maximum number of rows to process (for testing)"
    )
    
    return parser.parse_args()


def print_verbose_result(row_num: int, company: str, result: Dict[str, str]):
    """Print verbose results for a single company"""
    print(f"\n{'='*80}")
    print(f"ROW {row_num} COMPANY INFO FETCH")
    print(f"{'='*80}")
    print(f"Company: {company}")
    print(f"\nFETCHED INFORMATION:")
    print(f"{'─'*40}")
    print(f"Revenue: {result['AI_Revenue']}")
    print(f"CEO Name: {result['AI_CEO_Name']}")
    print(f"CEO Bio: {result['AI_CEO_Bio']}")
    print(f"LinkedIn URL: {result['AI_LinkedIn_URL']}")
    if result['AI_Error']:
        print(f"Error: {result['AI_Error']}")
    print(f"{'='*80}")


def main():
    """Main function"""
    args = parse_args()
    
    # Setup logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    logging.info("Starting company information fetching script")
    logging.info(f"Input file: {args.filename}")
    logging.info(f"Verbose mode: {args.verbose}")
    
    # Get API key
    api_key = get_openai_api_key()
    if not api_key:
        print("❌ Error: No OpenAI API key found")
        print("💡 Make sure to set OPENAI_API_KEY in environment or .env file")
        logging.error("No OpenAI API key found")
        return 1
    
    print(f"✅ OpenAI API key loaded (length: {len(api_key)} characters)")
    logging.info(f"OpenAI API key loaded successfully")
    
    # Check if input file exists
    if not os.path.exists(args.filename):
        print(f"❌ Error: File not found: {args.filename}")
        logging.error(f"Input file not found: {args.filename}")
        return 1
    
    try:
        # Read Excel file
        print(f"📂 Reading Excel file: {args.filename}")
        df = pd.read_excel(args.filename)
        
        print(f"📋 Found {len(df)} rows")
        print(f"📋 Columns: {list(df.columns)}")
        
        # Find company name column (first column or one containing 'company' and 'name')
        company_name_col = None
        
        # Check for exact matches first
        for col in df.columns:
            if col.lower() == 'company_name':
                company_name_col = col
                break
        
        # Fallback to first column if no exact match
        if not company_name_col:
            company_name_col = df.columns[0]
        
        print(f"📋 Company Name column: '{company_name_col}'")
        
        # Determine rows to process
        start_row = args.start_row - 1  # Convert to 0-based index
        end_row = len(df)
        if args.max_rows:
            end_row = min(start_row + args.max_rows, len(df))
        
        rows_to_process = end_row - start_row
        print(f"🔍 Processing rows {start_row + 1} to {end_row} ({rows_to_process} total)")
        
        # Add new columns for AI fetched information
        ai_columns = ['AI_Revenue', 'AI_CEO_Name', 'AI_CEO_Bio', 'AI_LinkedIn_URL', 'AI_Error']
        for col in ai_columns:
            if col not in df.columns:
                df[col] = ""
        
        # Process each row
        processed = 0
        fetched = 0
        
        for idx in range(start_row, end_row):
            company = str(df.iloc[idx][company_name_col]).strip()
            
            print(f"\n🔍 Processing row {idx + 1}: {company}")
            logging.info(f"Processing row {idx + 1}: {company}")
            
            if company and company.lower() != 'nan':
                # Fetch company information
                logging.info(f"Starting OpenAI company info fetch for: {company}")
                result = fetch_company_info(company, api_key)
                logging.info(f"Company info fetch completed for {company}")
                
                if result.get('AI_Error'):
                    logging.warning(f"Fetch error for {company}: {result['AI_Error']}")
                
                # Update the dataframe
                for col in ai_columns:
                    df.at[idx, col] = result[col]
                
                # Print verbose results if requested
                if args.verbose:
                    print_verbose_result(idx + 1, company, result)
                
                fetched += 1
                print(f"✅ Fetched info for row {idx + 1}")
                logging.info(f"Successfully fetched info for row {idx + 1}")
                
                # Add delay between API calls
                if args.delay > 0 and idx < end_row - 1:
                    print(f"⏳ Waiting {args.delay} seconds...")
                    logging.debug(f"Applying delay of {args.delay} seconds")
                    time.sleep(args.delay)
            else:
                print(f"⏭️ Skipping row {idx + 1}: No company name")
                logging.info(f"Skipping row {idx + 1}: No company name")
            
            processed += 1
        
        # Create Results directory if it doesn't exist
        results_dir = os.path.join(os.path.dirname(args.filename), "Results")
        if not os.path.exists(results_dir):
            os.makedirs(results_dir)
            print(f"📁 Created Results directory: {results_dir}")
        
        # Generate output filename
        timestamp = generate_timestamp()
        base_name = os.path.splitext(os.path.basename(args.filename))[0]
        output_filename = os.path.join(results_dir, f"{timestamp}_{base_name}.xlsx")
        
        # Save results
        print(f"\n💾 Saving fetched data to: {output_filename}")
        df.to_excel(output_filename, index=False)
        
        print(f"\n✅ Information fetching complete!")
        print(f"📊 Processed {processed} rows")
        print(f"🌐 Fetched info for {fetched} companies")
        print(f"📁 Output saved to: {output_filename}")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error: {e}")
        logging.error(f"Main function error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())