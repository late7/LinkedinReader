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

        # Extract lines
        lines = [l.strip() for l in block.splitlines() if l.strip()]
        text = " ".join(lines)

        # ---- Extract fields ----
        name_match = re.search(r'([A-Za-z0-9&.,\-\s]+?)\s+(?:logo|B2B|B2C|B2G)', text)
        company_name = name_match.group(1).strip() if name_match else ""
        
        # Remove "investor" prefix if present (case insensitive)
        if company_name.lower().startswith('investor '):
            company_name = company_name[9:].strip()

        # Location & year
        loc_match = re.search(r'([A-Za-z\s,]+)\s•\s(\d{4})', text)
        location = loc_match.group(1).strip() if loc_match else ""
        year = loc_match.group(2).strip() if loc_match else ""
        
        # Remove "investor" prefix from location if present (case insensitive)
        if location.lower().startswith('investor '):
            location = location[9:].strip()

        # Focus areas (B2B / B2C / B2G)
        focus = ", ".join(re.findall(r'\bB2[BGC]\b', text))

        # Description (between focus area(s) and "Team of")
        desc_match = re.search(r'(?:B2[BGC]\s*)+(.*?)(?:Team of)', text)
        description = desc_match.group(1).strip() if desc_match else ""

        # Team info
        team_match = re.search(r'Team of\s+(\d+)\s+•\s+([A-Za-z\s,]+)', text)
        team_size = team_match.group(1) if team_match else ""
        team_members = team_match.group(2).strip() if team_match else ""

        # Notable investments
        invest_match = re.search(r'Notable Investments\s*(.*?)\s*(?:Ticket Size|$)', text)
        notable_investments = invest_match.group(1).strip() if invest_match else ""

        # Ticket size
        ticket_match = re.search(r'Ticket Size\s*([0-9kM\-–]+)', text)
        ticket_size = ticket_match.group(1).strip() if ticket_match else ""

        investors.append({
            "Company Name": company_name,
            "Location": location,
            "Founded": year,
            "Focus Areas": focus,
            "Description": description,
            "Team Size": team_size,
            "Team Members": team_members,
            "Notable Investments": notable_investments,
            "Ticket Size": ticket_size
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
        
        # Add source file information to each investor record
        for i, investor in enumerate(all_investors):
            # Find which file this investor came from (simplified approach)
            investor["Source File"] = "Combined" if args.all else processed_files[0]
        
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
