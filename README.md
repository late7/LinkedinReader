# LinkedIn Reader

A collection of Python scripts for processing LinkedIn data and investor information.

## Scripts

### 1. LinkedIn Bio Fetcher (`fetch_linkedin_bios.py`)
Fetches profile bios from LinkedIn URLs stored in an Excel workbook.

**Features:**
- **Bio Extraction**: Extracts bio/description from LinkedIn profile pages
- **Background Check (Optional)**: AI-powered background checks using OpenAI API
- **Company Lookup (Optional)**: AI-powered company information lookup including email, phone, company type, industry, and revenue
- **Verbose Mode**: Prints detailed results to terminal as well as saving to Excel
- **Timestamped Output**: Automatically timestamped Excel files to avoid conflicts

### 2. Investor Data Enricher (`enrich_investor_data.py`)
Enriches Excel investor data with detailed research using OpenAI.

**Features:**
- **Investor Research**: Uses OpenAI to research detailed investment information
- **Company Website**: Finds official website URLs
- **Investment Profile**: Extracts stage, ticket size, sector focus, and strategy
- **Batch Processing**: Processes multiple companies from Excel file
- **Error Handling**: Graceful handling of API failures and missing data

### 3. Excel Data Processor (`makeExcel.py`)
Extracts investor data from text files and converts to Excel format.

**Features:**
- **Text Parsing**: Extracts structured data from unstructured text
- **Company Information**: Parses company type, name, location, funding details
- **Batch Processing**: Process single files or all files in input folder
- **Timestamped Output**: Avoids file conflicts with timestamps

## Usage

### LinkedIn Bio Fetcher

#### Basic Usage
```bash
python scripts/fetch_linkedin_bios.py
```

#### With Background Check
```bash
python scripts/fetch_linkedin_bios.py --bg
```

#### With Company Information Lookup
```bash
python scripts/fetch_linkedin_bios.py --company
```

#### With Verbose Output
```bash
python scripts/fetch_linkedin_bios.py --verbose
```

#### With All Features
```bash
python scripts/fetch_linkedin_bios.py --bg --company --verbose
```

### Investor Data Enricher

#### Basic Usage
```bash
python scripts/enrich_investor_data.py data.xlsx
```

#### With Verbose Output
```bash
python scripts/enrich_investor_data.py data.xlsx --verbose
```

#### Process Limited Rows (for testing)
```bash
python scripts/enrich_investor_data.py data.xlsx --max-rows 5 --verbose
```

#### Custom Delay Between API Calls
```bash
python scripts/enrich_investor_data.py data.xlsx --delay 3.0
```

### Excel Data Processor

#### Process Single File
```bash
python scripts/makeExcel.py data1.txt
```

#### Process All Files
```bash
python scripts/makeExcel.py -a
```

#### With Verbose Output
```bash
python scripts/makeExcel.py data1.txt --verbose
```

### Options

#### LinkedIn Bio Fetcher
- `--input INPUT`: Path to input Excel workbook (default: LinkedIN.xlsx)
- `--output OUTPUT`: Path for output workbook (default: Results/LinkedIn_Bios_{timestamp}.xlsx)
- `--delay DELAY`: Delay in seconds between requests (default: 0)
- `--bg`: Enable AI-powered background checks (requires OpenAI API key)
- `--company`: Enable AI-powered company information lookup (requires OpenAI API key)
- `--verbose`: Enable verbose mode - print detailed results to terminal as well as saving to Excel

#### Investor Data Enricher
- `--delay DELAY`: Delay in seconds between OpenAI API calls (default: 2.0)
- `--verbose, -v`: Print detailed progress and results
- `--start-row START_ROW`: Row to start processing from (default: 2, assuming row 1 is headers)
- `--max-rows MAX_ROWS`: Maximum number of rows to process (for testing)

#### Excel Data Processor
- `-a, --all`: Process all files in the 'input' folder
- `--verbose, -v`: Print extracted data to terminal

## Input/Output Formats

### LinkedIn Bio Fetcher
**Input:** Excel file with LinkedIn profile URLs in column A  
**Output:** Excel file with original URLs plus extracted bio information and optional AI enhancements

### Investor Data Enricher  
**Input:** Excel file with:
- Column A: Company/investor names
- Column B: City/location information

**Output:** Excel file with original data plus AI research columns:
- Website
- Investment Stage  
- Typical Ticket Size
- Sector Focus
- Investment Strategy

### Excel Data Processor
**Input:** Text files in 'input' folder containing investor/company information  
**Output:** Excel file with extracted structured data:
- Company names
- Company types (e.g., "Private Equity", "Venture Capital")  
- Structured tabular format

## Setup

### Environment Setup
1. Create a `.env` file in the repository root
2. Add your OpenAI API key:
   ```
   OPENAI_API_KEY=your_api_key_here
   ```

### Dependencies
- Standard library only for basic functionality
- Additional packages for full functionality:
  ```bash
  pip install pandas openpyxl requests beautifulsoup4 openai python-dotenv
  ```

### Folder Structure
Create the following folder structure:
```
LinkedinReader/
├── scripts/
├── Results/
├── input/          # For text files (Excel processor)
├── .env
└── your_excel_files
```

## Legacy Input Format (LinkedIn Bio Fetcher)

The input Excel file should contain a column named "LinkedIn Page" with LinkedIn profile URLs.

## Output Details

### LinkedIn Bio Fetcher Output

The script generates timestamped Excel files with:
- **Original columns** from input
- **Bio column** with extracted profile descriptions
- **Background Check column** (when --bg flag is used) with AI-generated background information
- **Company Info column** (when --company flag is used) with AI-generated company information:
  - Email (Sähköposti)
  - Phone number (puhelinnumero)
  - Company type (yrityksen tyyppi)
  - Industry (toimiala)
  - Latest revenue (viimeisin liikevaihto)

### Output File Naming
Files are automatically timestamped to prevent conflicts:
- Format: `LinkedIn_Bios_YYYYMMDD_HHMMSS.xlsx`
- Example: `LinkedIn_Bios_20251014_113001.xlsx`

### Verbose Mode Output
When using `--verbose`, the script displays detailed results in the terminal:
- **Processing status** for each profile
- **Formatted results** with clear section separators
- **Error handling** showing specific error messages
- **Processing summary** with statistics and enabled features
- **Visual formatting** with borders and section headers

Example verbose output:
```
================================================================================
ROW 2 RESULTS
================================================================================
URL: https://www.linkedin.com/in/example

BIO:
----------------------------------------
[Bio content here]

BACKGROUND CHECK:
----------------------------------------
[Background check results here]

COMPANY INFORMATION:
----------------------------------------
[Company information here]
================================================================================
```

## Security

- The `.env` file is automatically ignored by git
- API keys are never displayed in console output
- Only the length of the API key is shown for verification