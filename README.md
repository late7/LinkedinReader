# LinkedIn Reader

A Python script that fetches profile bios from LinkedIn URLs stored in an Excel workbook.

## Features

- **Bio Extraction**: Extracts bio/description from LinkedIn profile pages
- **Background Check (Optional)**: AI-powered background checks using OpenAI API
- **Company Lookup (Optional)**: AI-powered company information lookup including email, phone, company type, industry, and revenue
- **Timestamped Output**: Automatically timestamped Excel files to avoid conflicts

## Usage

### Basic Usage
```bash
python scripts/fetch_linkedin_bios.py
```

### With Background Check
```bash
python scripts/fetch_linkedin_bios.py --bg
```

### With Company Information Lookup
```bash
python scripts/fetch_linkedin_bios.py --company
```

### With Verbose Output
```bash
python scripts/fetch_linkedin_bios.py --verbose
```

### With All Features
```bash
python scripts/fetch_linkedin_bios.py --bg --company --verbose
```

### Options
- `--input INPUT`: Path to input Excel workbook (default: LinkedIN.xlsx)
- `--output OUTPUT`: Path for output workbook (default: Results/LinkedIn_Bios_{timestamp}.xlsx)
- `--delay DELAY`: Delay in seconds between requests (default: 0)
- `--bg`: Enable AI-powered background checks (requires OpenAI API key)
- `--company`: Enable AI-powered company information lookup (requires OpenAI API key)
- `--verbose`: Enable verbose mode - print detailed results to terminal as well as saving to Excel

## Setup

### Environment Setup
1. Create a `.env` file in the repository root
2. Add your OpenAI API key:
   ```
   OPENAI_API_KEY=your_api_key_here
   ```

### Dependencies
- Standard library only for basic functionality
- `openai` package required for background checks and company lookup:
  ```bash
  pip install openai
  ```

## Input Format

The input Excel file should contain a column named "LinkedIn Page" with LinkedIn profile URLs.

## Output

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