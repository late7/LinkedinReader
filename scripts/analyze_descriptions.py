#!/usr/bin/env python3
"""Analyze investor descriptions with OpenAI to extract structured investment data.

This script reads investor descriptions from an Excel file, uses OpenAI GPT-4o-mini
to extract structured investment information, and adds the results to new columns.
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


def analyze_description(description: str, existing_ticket_size: str, api_key: str) -> Dict[str, str]:
    """Analyze investor description using OpenAI to extract structured data.
    
    Args:
        description: Company description text
        existing_ticket_size: Existing ticket size data (if any)
        api_key: OpenAI API key
        
    Returns:
        Dictionary with extracted investment data
    """
    if not description or len(description) <= 50 or not api_key:
        return {
            "AI_SectorFocus": "",
            "AI_Stage": "",
            "AI_TicketSize_Min": "",
            "AI_TicketSize_Max": "",
            "AI_Website": "",
            "AI_Error": "Description too short or missing API key"
        }
    
    try:
        # Import OpenAI here to avoid dependency issues
        try:
            from openai import OpenAI
        except ImportError:
            return {
                "AI_SectorFocus": "",
                "AI_Stage": "",
                "AI_TicketSize_Min": "",
                "AI_TicketSize_Max": "",
                "AI_Website": "",
                "AI_Error": "OpenAI package not installed"
            }
        
        client = OpenAI(api_key=api_key)
        logging.debug(f"OpenAI client created for description analysis")
        
        # Prepare the analysis query
        developer_prompt = """Analyze this investor description and extract structured investment information.

Please return ONLY a JSON object with the following structure:
{
  "SectorFocus": ["Technology", "FinTech", "Healthcare", "etc"],
  "Stage": ["Pre-Seed", "Seed", "Series A", "Series B", "Growth", "etc"],
  "TicketSize": {
    "Min": "€100K",
    "Max": "€5M"
  },
  "Website": "www.example.com"
}

Important:
- SectorFocus: List of investment sectors/industries
- Stage: List of investment stages they focus on
- TicketSize: Extract or estimate investment amounts with currency
- Website: Only include if explicitly mentioned or well-known
- Use empty strings for unknown fields
- Return ONLY the JSON object, no other text"""

        user_prompt = f"""Description: {description}
Existing Ticket Size Info: {existing_ticket_size if existing_ticket_size else 'Not provided'}"""
        
        logging.info(f"Making OpenAI Responses API call for description analysis")
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
                tools=[],
                store=False,
                include=[
                    "reasoning.encrypted_content"
                ]
            )
            logging.info(f"OpenAI Responses API call successful for description analysis")
        except Exception as api_exc:
            logging.error(f"OpenAI Responses API call failed for description analysis: {api_exc}")
            return {
                "AI_SectorFocus": "",
                "AI_Stage": "",
                "AI_TicketSize_Min": "",
                "AI_TicketSize_Max": "",
                "AI_Website": "",
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
                "AI_SectorFocus": "",
                "AI_Stage": "",
                "AI_TicketSize_Min": "",
                "AI_TicketSize_Max": "",
                "AI_Website": "",
                "AI_Error": "Cannot extract response text"
            }
            
        if not response_text:
            logging.error(f"Empty response text")
            return {
                "AI_SectorFocus": "",
                "AI_Stage": "",
                "AI_TicketSize_Min": "",
                "AI_TicketSize_Max": "",
                "AI_Website": "",
                "AI_Error": "Empty response"
            }
            
        logging.info(f"OpenAI response received: {len(response_text)} characters")
        logging.debug(f"Raw response: {response_text}")

        # Parse JSON response
        try:
            logging.info(f"Attempting to parse JSON response")
            data = json.loads(response_text)
            logging.info(f"JSON parsing successful")
            logging.debug(f"Parsed data keys: {list(data.keys())}")
            
            # Extract information from the structured response
            sector_focus = data.get('SectorFocus', [])
            if isinstance(sector_focus, list):
                sector_str = ', '.join(sector_focus)
            else:
                sector_str = str(sector_focus)
            
            stage = data.get('Stage', [])
            if isinstance(stage, list):
                stage_str = ', '.join(stage)
            else:
                stage_str = str(stage)
            
            # Extract ticket size
            ticket_info = data.get('TicketSize', {})
            if isinstance(ticket_info, dict):
                min_val = ticket_info.get('Min', '')
                max_val = ticket_info.get('Max', '')
            else:
                min_val = str(ticket_info)
                max_val = ''
            
            website = data.get('Website', '')
            
            return {
                "AI_SectorFocus": sector_str,
                "AI_Stage": stage_str,
                "AI_TicketSize_Min": min_val,
                "AI_TicketSize_Max": max_val,
                "AI_Website": website,
                "AI_Error": ""
            }
            
        except json.JSONDecodeError as json_err:
            # If JSON parsing fails, return error with truncated response
            logging.error(f"JSON parsing failed: {json_err}")
            logging.error(f"Raw response that failed to parse: {response_text}")
            truncated_response = response_text[:100] + "..." if len(response_text) > 100 else response_text
            return {
                "AI_SectorFocus": "",
                "AI_Stage": "",
                "AI_TicketSize_Min": "",
                "AI_TicketSize_Max": "",
                "AI_Website": "",
                "AI_Error": f"JSON parsing failed: {truncated_response}"
            }
            
    except Exception as exc:
        logging.error(f"Unexpected error in description analysis: {exc}")
        logging.error(f"Exception type: {type(exc).__name__}")
        return {
            "AI_SectorFocus": "",
            "AI_Stage": "",
            "AI_TicketSize_Min": "",
            "AI_TicketSize_Max": "",
            "AI_Website": "",
            "AI_Error": str(exc)
        }


def generate_timestamp():
    """Generate timestamp in YYYYMMDD_HHMMSS format"""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def parse_args():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(
        description="Analyze investor descriptions with OpenAI to extract structured data"
    )
    parser.add_argument(
        "filename",
        help="Input Excel filename (e.g., Investors2025.xlsx)"
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=1.0,
        help="Delay in seconds between OpenAI API calls (default: 1.0)"
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


def print_verbose_result(row_num: int, company: str, description: str, result: Dict[str, str]):
    """Print verbose results for a single company"""
    print(f"\n{'='*80}")
    print(f"ROW {row_num} ANALYSIS")
    print(f"{'='*80}")
    print(f"Company: {company}")
    print(f"Description: {description[:100]}{'...' if len(description) > 100 else ''}")
    print(f"\nAI ANALYSIS RESULTS:")
    print(f"{'─'*40}")
    print(f"Sector Focus: {result['AI_SectorFocus']}")
    print(f"Investment Stage: {result['AI_Stage']}")
    print(f"Ticket Size Min: {result['AI_TicketSize_Min']}")
    print(f"Ticket Size Max: {result['AI_TicketSize_Max']}")
    print(f"Website: {result['AI_Website']}")
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
    
    logging.info("Starting description analysis script")
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
        
        # Find Description and Ticket Size columns
        description_col = None
        ticket_size_col = None
        company_name_col = None
        
        for col in df.columns:
            if 'description' in col.lower():
                description_col = col
            elif 'ticket' in col.lower() and 'size' in col.lower():
                ticket_size_col = col
            elif 'company' in col.lower() and 'name' in col.lower():
                company_name_col = col
        
        if not description_col:
            print("❌ Error: No 'Description' column found")
            return 1
        
        if not company_name_col:
            company_name_col = df.columns[0]  # Use first column as company name
        
        print(f"📋 Description column: '{description_col}'")
        print(f"📋 Ticket Size column: '{ticket_size_col or 'Not found'}'")
        print(f"📋 Company Name column: '{company_name_col}'")
        
        # Determine rows to process
        start_row = args.start_row - 1  # Convert to 0-based index
        end_row = len(df)
        if args.max_rows:
            end_row = min(start_row + args.max_rows, len(df))
        
        rows_to_process = end_row - start_row
        print(f"🔍 Processing rows {start_row + 1} to {end_row} ({rows_to_process} total)")
        
        # Add new columns for AI analysis results
        ai_columns = ['AI_SectorFocus', 'AI_Stage', 'AI_TicketSize_Min', 'AI_TicketSize_Max', 'AI_Website', 'AI_Error']
        for col in ai_columns:
            if col not in df.columns:
                df[col] = ""
        
        # Process each row
        processed = 0
        analyzed = 0
        
        for idx in range(start_row, end_row):
            company = str(df.iloc[idx][company_name_col]).strip() if company_name_col else f"Row {idx + 1}"
            description = str(df.iloc[idx][description_col]).strip()
            existing_ticket_size = str(df.iloc[idx][ticket_size_col]).strip() if ticket_size_col else ""
            
            print(f"\n🔍 Processing row {idx + 1}: {company}")
            logging.info(f"Processing row {idx + 1}: {company}")
            
            if description and description.lower() != 'nan' and len(description) > 50:
                print(f"   📝 Description length: {len(description)} characters")
                print(f"   📊 Existing ticket size: {existing_ticket_size or 'None'}")
                
                # Analyze the description
                logging.info(f"Starting OpenAI analysis for: {company}")
                result = analyze_description(description, existing_ticket_size, api_key)
                logging.info(f"Analysis completed for {company}")
                
                if result.get('AI_Error'):
                    logging.warning(f"Analysis error for {company}: {result['AI_Error']}")
                
                # Update the dataframe
                for col in ai_columns:
                    df.at[idx, col] = result[col]
                
                # Print verbose results if requested
                if args.verbose:
                    print_verbose_result(idx + 1, company, description, result)
                
                analyzed += 1
                print(f"✅ Analyzed row {idx + 1}")
                logging.info(f"Successfully analyzed row {idx + 1}")
                
                # Add delay between API calls
                if args.delay > 0 and idx < end_row - 1:
                    print(f"⏳ Waiting {args.delay} seconds...")
                    logging.debug(f"Applying delay of {args.delay} seconds")
                    time.sleep(args.delay)
            else:
                reason = "No description" if not description or description.lower() == 'nan' else f"Description too short ({len(description)} chars)"
                print(f"⏭️ Skipping row {idx + 1}: {reason}")
                logging.info(f"Skipping row {idx + 1}: {reason}")
            
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
        print(f"\n💾 Saving analyzed data to: {output_filename}")
        df.to_excel(output_filename, index=False)
        
        print(f"\n✅ Analysis complete!")
        print(f"📊 Processed {processed} rows")
        print(f"🤖 Analyzed {analyzed} descriptions with AI")
        print(f"📁 Output saved to: {output_filename}")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error: {e}")
        logging.error(f"Main function error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())