import re
import pandas as pd
import argparse
import os
from datetime import datetime

def generate_timestamp():
    """Generate timestamp in YYYYMMDD_HHMMSS format"""
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def read_input_file(filename):
    """Read data from input folder"""
    input_path = os.path.join("input", filename)
    
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"File not found: {input_path}")
    
    with open(input_path, 'r', encoding='utf-8') as file:
        return file.read()

def get_all_input_files():
    """Get all files from the input folder"""
    input_folder = "input"
    
    if not os.path.exists(input_folder):
        raise FileNotFoundError(f"Input folder not found: {input_folder}")
    
    files = []
    for filename in os.listdir(input_folder):
        filepath = os.path.join(input_folder, filename)
        if os.path.isfile(filepath):
            files.append(filename)
    
    if not files:
        raise ValueError(f"No files found in {input_folder} folder")
    
    return files

def parse_args():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(
        description="Extract investor data from text file and convert to Excel"
    )
    
    # Create mutually exclusive group for filename vs all files
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "filename", 
        nargs="?",
        help="Input filename (must be in 'input' folder)"
    )
    group.add_argument(
        "-a", "--all",
        action="store_true",
        help="Process all files in the 'input' folder"
    )
    
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Print extracted data to terminal"
    )
    
    return parser.parse_args()

def extract_investor_data(data):
    """Extract investor data from text"""
    # ---- SPLIT EACH COMPANY BLOCK ----
    blocks = re.split(r'\n\s*View company\s*\n', data)
    investors = []

    for block in blocks:
        if not block.strip():
            continue

        # Split into lines and clean
        lines = [l.strip() for l in block.splitlines() if l.strip()]
        if len(lines) < 3:  # Need at least company type, name, and location
            continue
        
        # Extract company type (first line after "View company")
        company_type = lines[0] if lines[0] else ""
        
        # Skip logo line and find company name
        company_name = ""
        location_line = ""
        remaining_lines = []
        
        # Find company name (skip logo lines)
        name_start_idx = 1
        for i in range(1, len(lines)):
            if not lines[i].endswith(' logo') and not lines[i].startswith('+'):
                # Check if this looks like a location line (contains country code and year)
                if re.search(r',\s[A-Z]{2}\s•\s\d{4}', lines[i]):
                    location_line = lines[i]
                    remaining_lines = lines[i+1:]
                    break
                else:
                    # This should be the company name
                    company_name = lines[i]
                    # Look for location line next
                    if i+1 < len(lines) and re.search(r',\s[A-Z]{2}\s•\s\d{4}', lines[i+1]):
                        location_line = lines[i+1]
                        remaining_lines = lines[i+2:]
                        break
        
        # Parse location and year
        location = ""
        year = ""
        if location_line:
            loc_match = re.search(r'([^•]+)\s•\s(\d{4})', location_line)
            if loc_match:
                location = loc_match.group(1).strip()
                year = loc_match.group(2).strip()
        
        # Join remaining lines for further parsing
        remaining_text = " ".join(remaining_lines)
        
        # Extract focus areas (B2B / B2C / B2G)
        focus = ", ".join(re.findall(r'\bB2[BGC]\b', remaining_text))
        
        # Extract description (find text between B2X and "Team of" or other specific patterns)
        description = ""
        # Remove +numbers from text
        clean_text = re.sub(r'\+\d+', '', remaining_text)
        
        # Find description between focus areas and team info
        desc_pattern = r'(?:B2[BGC]\s*(?:\+\d+\s*)?)(.*?)(?:Team of|\€|\$|funding|Next raising|$)'
        desc_match = re.search(desc_pattern, clean_text, re.DOTALL)
        if desc_match:
            description = desc_match.group(1).strip()
        
        # Extract team info
        team_size = ""
        team_members = ""
        team_match = re.search(r'Team of\s+(\d+)\s+•\s+([^€$\n]+)', remaining_text)
        if team_match:
            team_size = team_match.group(1)
            team_members = team_match.group(2).strip()
        
        # Extract funding information
        funding_info = ""
        funding_match = re.search(r'(€\d+[KMB]?|$\d+[KMB]?)\s*(?:funding from\s*([^€$\n]+))?', remaining_text)
        if funding_match:
            amount = funding_match.group(1)
            investors_list = funding_match.group(2) if funding_match.group(2) else ""
            funding_info = f"{amount} from {investors_list}".strip()
        
        # Extract next raising info
        next_raising = ""
        next_match = re.search(r'Next raising\s+([^€$\n]+)\s+of\s+(€\d+[KMB]?|$\d+[KMB]?)', remaining_text)
        if next_match:
            round_type = next_match.group(1).strip()
            amount = next_match.group(2).strip()
            next_raising = f"{round_type} of {amount}"

        investors.append({
            "Filename": "",  # Will be set by process_single_file function
            "Company Type": company_type,
            "Company Name": company_name,
            "Location": location,
            "Founded": year,
            "Focus Areas": focus,
            "Description": description,
            "Team Size": team_size,
            "Team Members": team_members,
            "Funding": funding_info,
            "Next Raising": next_raising
        })

    return investors

def print_verbose_results(investors):
    """Print extracted data in verbose mode"""
    print("\n" + "="*80)
    print("EXTRACTED INVESTOR DATA")
    print("="*80)
    
    for i, investor in enumerate(investors, 1):
        print(f"\n--- INVESTOR {i} ---")
        for key, value in investor.items():
            print(f"{key}: {value}")
    
    print(f"\n{'='*80}")
    print(f"TOTAL INVESTORS EXTRACTED: {len(investors)}")
    print("="*80)

def process_single_file(filename, verbose=False):
    """Process a single file and return extracted data"""
    print(f"📂 Processing file: input/{filename}")
    
    try:
        # Read input data
        data = read_input_file(filename)
        print(f"✅ Successfully read {len(data)} characters from {filename}")
        
        # Extract investor data
        print("🔍 Extracting investor data...")
        investors = extract_investor_data(data)
        print(f"✅ Extracted {len(investors)} investors from {filename}")
        
        # Add filename (without .txt extension) to each record
        filename_without_ext = filename.replace('.txt', '') if filename.endswith('.txt') else filename
        for investor in investors:
            investor["Filename"] = filename_without_ext
        
        # Print verbose output if requested
        if verbose:
            print(f"\n--- VERBOSE OUTPUT FOR {filename} ---")
            print_verbose_results(investors)
        
        return investors, filename
        
    except Exception as e:
        print(f"❌ Error processing {filename}: {e}")
        return [], filename

def main():
    """Main function"""
    args = parse_args()
    
    all_investors = []
    processed_files = []
    
    try:
        if args.all:
            # Process all files in input folder
            print("📂 Processing ALL files in input folder...")
            files = get_all_input_files()
            print(f"📁 Found {len(files)} files: {', '.join(files)}")
            
            for filename in files:
                investors, processed_file = process_single_file(filename, args.verbose)
                all_investors.extend(investors)
                processed_files.append(processed_file)
                print()  # Add spacing between files
        else:
            # Process single file
            print(f"📂 Processing single file: {args.filename}")
            investors, processed_file = process_single_file(args.filename, args.verbose)
            all_investors.extend(investors)
            processed_files.append(processed_file)
        
        if not all_investors:
            print("❌ No investor data found in any files")
            return
        
        # Create output filename with timestamp
        timestamp = generate_timestamp()
        output_filename = f"{timestamp}_export.xlsx"
        
        # Save to Excel
        print(f"💾 Saving combined data to Excel: {output_filename}")
        df = pd.DataFrame(all_investors)
        df.to_excel(output_filename, index=False)
        
        print(f"✅ Data extracted and saved to {output_filename}")
        print(f"📊 Total records: {len(all_investors)}")
        print(f"📁 Files processed: {', '.join(processed_files)}")
        
    except FileNotFoundError as e:
        print(f"❌ Error: {e}")
        print("💡 Make sure the 'input' folder exists and contains your file(s)")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")

if __name__ == "__main__":
    main()
