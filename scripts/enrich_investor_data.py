#!/usr/bin/env python3
"""Enrich Excel investor data with OpenAI research.

This script reads investor company names and cities from an Excel file (columns A and D),
uses OpenAI to research detailed investment information, and adds the results to new columns.
"""

import argparse
import json
import logging
import os
import sys
import time
from datetime import datetime
from typing import Dict, List, Optional, Any

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


def research_investor_with_web(company_name: str, city: str, api_key: str) -> Dict[str, str]:
    """Research investor information using OpenAI with web search enabled.
    
    Args:
        company_name: Name of the investor company
        city: City where the company is located
        api_key: OpenAI API key
        
    Returns:
        Dictionary with investor research results
    """
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        
        # Prepare the query with web search
        query_text = f"Find information defined in response JSON below. Make InvestmentStrategy very short. Company: {company_name}"
        if city:
            query_text += f", City: {city}"
            
        query_text += """
{
  "Investor": "[Company Name]",
  "www": "[website.com]",
  "InvestmentProfile": {
    "Stage": ["Seed", "Series A", "etc"],
    "TicketSize": {
      "Currency": "EUR/USD",
      "Range": "€X - €Y",
      "Typical": "Around €X"
    },
    "SectorFocus": [
      "Technology",
      "B2B SaaS", 
      "etc"
    ],
    "InvestmentStrategy": "Brief strategy description"
  }
}"""
        
        logging.info(f"Making web-enabled OpenAI API call for {company_name}")
        response = client.responses.create(
            model="gpt-5",
            input=[
                {
                    "role": "developer",
                    "content": [
                        {
                            "type": "input_text",
                            "text": "Your are financial analyst. User gives you companies one by one and your task is to find information. Answer only JSON. No Sources, explanation, summary, nothing but just JSON."
                        }
                    ]
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": query_text
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
                        "country": "US",
                        "region": "New York",
                        "city": "New York"
                    },
                    "search_context_size": "medium"
                }
            ],
            store=False,
            include=[
                "reasoning.encrypted_content"
            ]
        )
        logging.info(f"Web-enabled OpenAI API call successful for {company_name}")
        print(f"Web-enabled OpenAI API call successful for {company_name}")
        
        # Extract response text
        response_text = ""
        if hasattr(response, 'output_text') and response.output_text:
            response_text = response.output_text
            logging.info(f"Web search response received for {company_name}: {len(response_text)} characters")
            logging.debug(f"Web search raw response: {response_text[:500]}{'...' if len(response_text) > 500 else ''}")
        else:
            logging.error(f"No output_text in web search response for {company_name}")
            return {
                "Website": "ERROR: No web search response",
                "Investment_Stage": "",
                "Ticket_Size": "",
                "Sector_Focus": "",
                "Investment_Strategy": "",
                "Error": "No web search response"
            }

        # Parse JSON response
        try:
            logging.info(f"Attempting to parse web search JSON response for {company_name}")
            data = json.loads(response_text)
            logging.info(f"Web search JSON parsing successful for {company_name}")
            
            # Extract information from the structured response
            website = data.get('www', '')
            investment_profile = data.get('InvestmentProfile', {})
            
            # Format stages
            stages = investment_profile.get('Stage', [])
            stage_str = ', '.join(stages) if isinstance(stages, list) else str(stages)
            
            # Format ticket size
            ticket_info = investment_profile.get('TicketSize', {})
            if isinstance(ticket_info, dict):
                currency = ticket_info.get('Currency', '')
                range_val = ticket_info.get('Range', '')
                typical = ticket_info.get('Typical', '')
                ticket_str = f"{range_val} ({typical})" if range_val and typical else range_val or typical
            else:
                ticket_str = str(ticket_info)
            
            # Format sector focus
            sectors = investment_profile.get('SectorFocus', [])
            sector_str = ', '.join(sectors) if isinstance(sectors, list) else str(sectors)
            
            # Get strategy
            strategy = investment_profile.get('InvestmentStrategy', '')
            
            return {
                "Website": website,
                "Investment_Stage": stage_str,
                "Ticket_Size": ticket_str,
                "Sector_Focus": sector_str,
                "Investment_Strategy": strategy,
                "Error": ""
            }
            
        except json.JSONDecodeError as json_err:
            logging.error(f"Web search JSON parsing failed for {company_name}: {json_err}")
            logging.error(f"Web search raw response that failed to parse: {response_text}")
            truncated_response = response_text[:200] + "..." if len(response_text) > 200 else response_text
            return {
                "Website": "Web Search JSON Parse Error",
                "Investment_Stage": "",
                "Ticket_Size": "",
                "Sector_Focus": "",
                "Investment_Strategy": truncated_response,
                "Error": f"Web search JSON parsing failed: {json_err}"
            }
            
    except Exception as exc:
        logging.error(f"Web search error for {company_name}: {exc}")
        return {
            "Website": f"Web Search ERROR: {exc}",
            "Investment_Stage": "",
            "Ticket_Size": "",
            "Sector_Focus": "",
            "Investment_Strategy": "",
            "Error": str(exc)
        }


def research_investor(company_name: str, city: str, api_key: str) -> Dict[str, str]:
    """Research investor information using OpenAI.
    
    Args:
        company_name: Name of the investor company
        city: City where the company is located
        api_key: OpenAI API key
        
    Returns:
        Dictionary with investor research results
    """
    if not company_name or not api_key:
        return {
            "Website": "ERROR: Missing company name or API key",
            "Investment_Stage": "",
            "Ticket_Size": "",
            "Sector_Focus": "",
            "Investment_Strategy": "",
            "Error": "Missing data"
        }
    
    try:
        # Import OpenAI here to avoid dependency issues
        try:
            from openai import OpenAI
        except ImportError:
            return {
                "Website": "ERROR: OpenAI package not installed",
                "Investment_Stage": "",
                "Ticket_Size": "",
                "Sector_Focus": "",
                "Investment_Strategy": "",
                "Error": "Missing OpenAI package"
            }
        
        client = OpenAI(api_key=api_key)
        logging.debug(f"OpenAI client created for {company_name}")
        
        # Prepare the query
        query_text = f"Find information defined in response JSON below. Make InvestmentStrategy very short. Company: {company_name}"
        if city:
            query_text += f", City: {city}"
            
        logging.debug(f"Query prepared for {company_name}: {query_text[:100]}...")
        
        query_text += """
{
  "Investor": "[Company Name]",
  "www": "[website.com]",
  "InvestmentProfile": {
    "Stage": ["Seed", "Series A", "etc"],
    "TicketSize": {
      "Currency": "EUR/USD",
      "Range": "€X - €Y",
      "Typical": "Around €X"
    },
    "SectorFocus": [
      "Technology",
      "B2B SaaS", 
      "etc"
    ],
    "InvestmentStrategy": "Brief strategy description"
  }
}"""
        
        logging.info(f"Making OpenAI API call for {company_name}")
        try:
            response = client.responses.create(
            model="gpt-5-mini",  # Using gpt-4o instead of gpt-5-mini for better availability
            input=[
                {
                    "role": "developer",
                    "content": [
                        {
                            "type": "input_text",
                            "text": "You are a financial analyst. User gives you companies one by one and your task is to find investment information"
                        }
                    ]
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": query_text
                        }
                    ]
                }
            ],
            text={
                "format": {
                    "type": "json_object"
                },
                "verbosity": "medium"
            },
            reasoning={
                "effort": "medium",
                "summary": None
            },
            store=False,
            include=[
                "reasoning.encrypted_content"
            ]
        )
            logging.info(f"OpenAI API call successful for {company_name}")
        except Exception as api_exc:
            logging.error(f"OpenAI API call failed for {company_name}: {api_exc}")
            return {
                "Website": f"API Error: {api_exc}",
                "Investment_Stage": "",
                "Ticket_Size": "",
                "Sector_Focus": "",
                "Investment_Strategy": "",
                "Error": f"API call failed: {api_exc}"
            }
        
        # Extract and parse the response
        response_text = ""
        logging.debug(f"Raw response object type: {type(response)}")
        
        # Try different ways to extract the response text
        if hasattr(response, 'output_text') and response.output_text:
            response_text = response.output_text
            logging.debug(f"Got output_text: {response_text[:100]}...")
        elif hasattr(response, 'text') and response.text:
            response_data = response.text
            if isinstance(response_data, str):
                response_text = response_data
                logging.debug(f"Got string response: {response_text[:100]}...")
            else:
                response_text = str(response_data)
                logging.debug(f"Converted response.text to string: {response_text[:100]}...")
        elif hasattr(response, 'output') and response.output:
            response_text = str(response.output)
            logging.debug(f"Got output from response: {response_text[:100]}...")
        else:
            logging.error(f"Unable to extract text from response. Response type: {type(response)}")
            logging.error(f"Available attributes: {[attr for attr in dir(response) if not attr.startswith('_')]}")
            return {
                "Website": "ERROR: Cannot extract response text",
                "Investment_Stage": "",
                "Ticket_Size": "",
                "Sector_Focus": "",
                "Investment_Strategy": "",
                "Error": "Cannot extract response text"
            }
            
        if not response_text:
            logging.error(f"Empty response text for {company_name}")
            return {
                "Website": "ERROR: Empty response",
                "Investment_Stage": "",
                "Ticket_Size": "",
                "Sector_Focus": "",
                "Investment_Strategy": "",
                "Error": "Empty response"
            }
            
        logging.info(f"OpenAI response received for {company_name}: {len(response_text)} characters")
        logging.debug(f"Raw response: {response_text[:500]}{'...' if len(response_text) > 500 else ''}")

        # Parse JSON response
        try:
            logging.info(f"Attempting to parse JSON response for {company_name}")
            data = json.loads(response_text)
            logging.info(f"JSON parsing successful for {company_name}")
            logging.debug(f"Parsed data keys: {list(data.keys())}")            # Extract information from the structured response
            website = data.get('www', '')
            investment_profile = data.get('InvestmentProfile', {})
            
            # Format stages
            stages = investment_profile.get('Stage', [])
            stage_str = ', '.join(stages) if isinstance(stages, list) else str(stages)
            
            # Format ticket size
            ticket_info = investment_profile.get('TicketSize', {})
            if isinstance(ticket_info, dict):
                currency = ticket_info.get('Currency', '')
                range_val = ticket_info.get('Range', '')
                typical = ticket_info.get('Typical', '')
                ticket_str = f"{range_val} ({typical})" if range_val and typical else range_val or typical
            else:
                ticket_str = str(ticket_info)
            
            # Format sector focus
            sectors = investment_profile.get('SectorFocus', [])
            sector_str = ', '.join(sectors) if isinstance(sectors, list) else str(sectors)
            
            # Get strategy
            strategy = investment_profile.get('InvestmentStrategy', '')
            
            # Check if the response contains meaningful data or just placeholders
            is_empty_response = (
                not website or website in ["[website.com]", ""] or
                not stage_str or stage_str in ["etc", ""] or
                not ticket_str or ticket_str in ["€X - €Y", "Around €X", ""] or
                not sector_str or sector_str in ["etc", ""] or
                not strategy or strategy in ["Brief strategy description", ""]
            )
            
            if is_empty_response:
                logging.warning(f"Empty or placeholder response for {company_name}, trying web search fallback")
                web_result = research_investor_with_web(company_name, city, api_key)
                
                # Check if web search provided better results
                if web_result.get('Website') and not web_result.get('Error'):
                    logging.info(f"Web search provided better results for {company_name}")
                    return web_result
                else:
                    logging.warning(f"Web search also failed for {company_name}, returning original response")
            
            return {
                "Website": website,
                "Investment_Stage": stage_str,
                "Ticket_Size": ticket_str,
                "Sector_Focus": sector_str,
                "Investment_Strategy": strategy,
                "Error": ""
            }
            
        except json.JSONDecodeError as json_err:
            # If JSON parsing fails, return raw response
            logging.error(f"JSON parsing failed for {company_name}: {json_err}")
            logging.error(f"Raw response that failed to parse: {response_text}")
            truncated_response = response_text[:200] + "..." if len(response_text) > 200 else response_text
            return {
                "Website": "JSON Parse Error",
                "Investment_Stage": "",
                "Ticket_Size": "",
                "Sector_Focus": "",
                "Investment_Strategy": truncated_response,
                "Error": f"JSON parsing failed: {json_err}"
            }
            
    except Exception as exc:
        logging.error(f"Unexpected error researching {company_name}: {exc}")
        logging.error(f"Exception type: {type(exc).__name__}")
        return {
            "Website": f"ERROR: {exc}",
            "Investment_Stage": "",
            "Ticket_Size": "",
            "Sector_Focus": "",
            "Investment_Strategy": "",
            "Error": str(exc)
        }


def generate_timestamp():
    """Generate timestamp in YYYYMMDD_HHMMSS format"""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def parse_args():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(
        description="Enrich Excel investor data with OpenAI research"
    )
    parser.add_argument(
        "filename",
        help="Input Excel filename (e.g., data.xlsx)"
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
        default=2,
        help="Row to start processing from (default: 2, assuming row 1 is headers)"
    )
    parser.add_argument(
        "--max-rows",
        type=int,
        help="Maximum number of rows to process (for testing)"
    )
    
    return parser.parse_args()


def print_verbose_result(row_num: int, company: str, city: str, result: Dict[str, str]):
    """Print verbose results for a single company"""
    print(f"\n{'='*80}")
    print(f"ROW {row_num} RESULTS")
    print(f"{'='*80}")
    print(f"Company: {company}")
    print(f"City: {city}")
    print(f"\nRESEARCH RESULTS:")
    print(f"{'─'*40}")
    print(f"Website: {result['Website']}")
    print(f"Investment Stage: {result['Investment_Stage']}")
    print(f"Ticket Size: {result['Ticket_Size']}")
    print(f"Sector Focus: {result['Sector_Focus']}")
    print(f"Strategy: {result['Investment_Strategy']}")
    if result['Error']:
        print(f"Error: {result['Error']}")
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
    
    logging.info("Starting investor data enrichment script")
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
        
        if len(df.columns) < 2:
            print("❌ Error: Excel file must have at least 2 columns (Company Name and City)")
            return 1
        
        # Get company and city columns (A and D)
        company_col = df.columns[0]  # Column A
        city_col = df.columns[3]     # Column D
        
        print(f"📋 Found {len(df)} rows")
        print(f"📋 Company column: '{company_col}'")
        print(f"📋 City column: '{city_col}'")
        
        # Determine rows to process
        start_row = args.start_row - 1  # Convert to 0-based index
        end_row = len(df)
        if args.max_rows:
            end_row = min(start_row + args.max_rows, len(df))
        
        rows_to_process = end_row - start_row
        print(f"🔍 Processing rows {start_row + 1} to {end_row} ({rows_to_process} total)")
        
        # Add new columns for results
        new_columns = ['Website', 'Investment_Stage', 'Ticket_Size', 'Sector_Focus', 'Investment_Strategy']
        for col in new_columns:
            if col not in df.columns:
                df[col] = ""
        
        # Process each row
        processed = 0
        for idx in range(start_row, end_row):
            company = str(df.iloc[idx][company_col]).strip()
            city = str(df.iloc[idx][city_col]).strip()
            
            if company and company.lower() != 'nan':
                print(f"\n🔍 Processing row {idx + 1}: {company}")
                logging.info(f"Processing row {idx + 1}: {company} (City: {city})")
                if city and city.lower() != 'nan':
                    print(f"   📍 City: {city}")
                
                # Research the investor
                logging.info(f"Starting OpenAI research for: {company}")
                result = research_investor(company, city, api_key)
                logging.info(f"Research completed for {company}. Result keys: {list(result.keys())}")
                
                if result.get('Error'):
                    logging.warning(f"Research error for {company}: {result['Error']}")
                
                # Update the dataframe using .at for cleaner assignment
                df.at[idx, 'Website'] = result['Website']
                df.at[idx, 'Investment_Stage'] = result['Investment_Stage']
                df.at[idx, 'Ticket_Size'] = result['Ticket_Size']
                df.at[idx, 'Sector_Focus'] = result['Sector_Focus']
                df.at[idx, 'Investment_Strategy'] = result['Investment_Strategy']
                
                # Print verbose results if requested
                if args.verbose:
                    print_verbose_result(idx + 1, company, city, result)
                
                processed += 1
                print(f"✅ Completed row {idx + 1}")
                logging.info(f"Successfully processed row {idx + 1}")
                
                # Add delay between API calls
                if args.delay > 0 and idx < end_row - 1:
                    print(f"⏳ Waiting {args.delay} seconds...")
                    logging.debug(f"Applying delay of {args.delay} seconds")
                    time.sleep(args.delay)
            else:
                print(f"⏭️ Skipping row {idx + 1}: No company name")
        
        # Generate output filename
        timestamp = generate_timestamp()
        base_name = os.path.splitext(args.filename)[0]
        output_filename = f"{base_name}_enriched_{timestamp}.xlsx"
        
        # Save results
        print(f"\n💾 Saving enriched data to: {output_filename}")
        df.to_excel(output_filename, index=False)
        
        print(f"\n✅ Processing complete!")
        print(f"📊 Processed {processed} companies")
        print(f"📁 Output saved to: {output_filename}")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())