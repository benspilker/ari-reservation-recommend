# Azure Reservation Recommendation Tool

This tool analyzes Azure Virtual Machines and generates recommendations for cost savings through Reserved Instances. It combines data from Azure Resource Inventory and Azure Advisor to provide comprehensive pricing analysis.

## Features

- ✅ **Automatic dependency management** - Installs required packages automatically
- ✅ **VM recommendation analysis** - Identifies VMs that would benefit from reservations
- ✅ **Parallel processing** - Runs Azure API and web scraping simultaneously for faster execution
- ✅ **Multiple output formats** - JSON and Excel spreadsheets with detailed pricing
- ✅ **Savings calculations** - Compares Pay-as-you-go vs 1-year and 3-year reservations

## Requirements

### Software Requirements
- **Python 3.8+**
- **Google Chrome browser** (for web scraping)
- **ChromeDriver** (matching your Chrome version)
  - Download: https://chromedriver.chromium.org/
  - Or install via: `choco install chromedriver` (Windows) or `brew install chromedriver` (Mac)

### Python Dependencies
The script will automatically install these if missing:
- pandas >= 2.0.0
- openpyxl >= 3.0.0
- requests >= 2.28.0
- selenium >= 4.0.0

Or install manually:
```bash
pip install -r requirements.txt
```

### Input File Required
- **Azure Resource Inventory Report** (format: `AzureResourceInventory_Report_*.xlsx`)
  - Must contain "Advisor" and "Virtual Machines" sheets
  - The script will automatically find the most recent report in the current directory

## Usage

### Quick Start (Put python script in the same folder as your AzureResourceInventory_Report_*.xlsx)
```bash
python azure-reservation-analysis.py
```

### Workflow

**Section 1: VM Recommendations**
1. Script loads the Azure Resource Inventory file
2. Filters running VMs and high-impact cost recommendations
3. Displays summary and impact analysis
4. Creates `output.json` and `inputs.json`
5. **Pauses and asks if you want to continue**

**Section 2: Pricing Analysis** (if you choose to continue)
1. Queries Azure Retail Prices API
2. Scrapes Windows pricing from vantage.sh (in parallel)
3. Generates detailed Excel spreadsheets
4. Calculates savings for 1-year and 3-year reservations

## Output Files

### JSON Files
- **output.json** - High-impact VM recommendations with annual savings
- **inputs.json** - Input data for pricing analysis
- **azure_windows_pricing_data.json** - Cached Windows pricing from vantage.sh
- **skus-regions-windows.json** - List of SKU-region pairs processed

### Excel Spreadsheets
- **azure_compute_estimate.xlsx** - Detailed **compute-only** pricing with Azure API data (NO SOFTWARE OS LICENSING COSTS)
  - Pay-as-you-go pricing
  - 1-year and 3-year reservation pricing
  - Monthly cost breakdowns
  - Totals section
  - **Note:** Includes automatic failover to vantage.sh when Azure API has no data

- **azure_savings_estimate.xlsx** - Enhanced pricing with vantage.sh data to include estimated compute and software OS costs
  - Sheet1: Formatted with blank rows (3 lines + blank pattern)
  - FilterMe: Flattened data (no blank rows for easy filtering)
  - Annual savings calculations
  - 3-year total cost comparisons
  - **Bold highlights on key savings metrics**

- **ranked_vms.xlsx** - VMs sorted by pay-as-you-go cost (highest to lowest)

## Example Output

```
=== Impact Summary ===
{
    "Total VMs": 847,
    "Savings (USD) According to ARI": 2859253.26,
    "Average Savings per VM (USD)": 3376.27
}
```

## Configuration

### Savings Threshold
By default, the script filters VMs with annual savings >= $100. To change this, edit line 87:
```python
savings_threshold = 100  # Change this value (if desired)
```

### Minimum Recommendations
The script ensures at least 20 VMs are included in the analysis. To change this, edit lines 207 and 211:
```python
if len(impact_recs) < 20:  # Change this value (if desired)
```

## Troubleshooting

### ChromeDriver Issues
If you see errors like "chromedriver not found" or "Session not created":
1. Verify Chrome is installed
2. Download ChromeDriver matching your Chrome version
3. Add ChromeDriver to your system PATH
4. Or place chromedriver.exe in the same folder as the script

### Missing Azure Resource Inventory
If you see "No file found matching AzureResourceInventory_Report_*.xlsx":
1. Ensure the file is in the same directory as the script
2. Verify the filename matches the pattern
3. Check that it contains "Advisor" and "Virtual Machines" sheets

### Import Errors
The script should auto-install missing packages. If it fails:
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

## Performance Notes

- **Section 1** typically completes in < 1 minute
- **Section 2** duration depends on:
  - Number of VMs to analyze
  - Number of unique SKU-region combinations
  - Network speed
  - Typically 3-5 minutes for 500-1000 VMs

## Contributing

Feel free to submit issues or pull requests for improvements!

## License

This tool is provided as-is for Azure cost analysis purposes.

