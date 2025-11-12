# Amazon Product Details Scraper

A Python tool to scrape detailed product information from Amazon India and export the data to beautifully formatted Excel files.

## Features

- Smart extraction of product details from Amazon India
- Extracts: Title, Price, Rating, Review Count, Bestseller Rank, Main Category, Sub Category
- Batch processing of multiple URLs
- Formatted Excel output with headers and styling
- Headless browser operation (no window opens)
- Cross-platform support (Windows, Linux, macOS)
- Automatic ChromeDriver management

## Data Extracted

| Field | Description | Example |
|-------|-------------|---------|
| ASIN | Amazon Standard Identification Number | B0CZL9BM4S |
| Product Title | Full product name | Logitech Signature M650 L Wireless Mouse |
| Price | Current price | 2,395 |
| Rating | Star rating | 4.2 |
| Number of Ratings | Count of customer ratings | 2,214 ratings |
| Bestseller Rank | Product rank in category | #538 |
| Main Category | Primary category | Electronics |
| Sub Category | Secondary category | Computer Peripherals |

## Requirements

- Python 3.8 or higher
- Google Chrome or Chromium browser
- Internet connection

## Installation

### 1. Install Python (if not already installed)

Download from https://www.python.org/downloads/

### 2. Install Dependencies

Create a `requirements.txt` file with:

```
selenium==4.15.0
openpyxl==3.11.0
webdriver-manager==4.0.1
```

Then install:

```bash
pip install -r requirements.txt
```

### 3. Verify Installation

```bash
python --version
```

## Quick Start

### Single URL

```bash
python amazon_exporter.py --input "https://www.amazon.in/dp/B0CZL9BM4S"
```

### Multiple URLs from File

Create `urls.txt`:

```
https://www.amazon.in/dp/B0CZL9BM4S
https://www.amazon.in/dp/ANOTHER_ASIN
https://www.amazon.in/dp/THIRD_ASIN
```

Then run:

```bash
python amazon_exporter.py --input "urls.txt"
```

### Custom Output File

```bash
python amazon_exporter.py --input "urls.txt" --output "my_results.xlsx"
```

### Auto-generated Output (default)

If no `--output` specified, file is saved as:

```
amazon_products_20240115_143022.xlsx
```

## Usage Examples

### Example 1: Single Product

```bash
python amazon_exporter.py --input "https://www.amazon.in/dp/B0CZL9BM4S" --output "polo_shirt.xlsx"
```

### Example 2: Batch Processing

Create `electronics.txt`:

```
https://www.amazon.in/dp/B08ABC123
https://www.amazon.in/dp/B08ABC456
https://www.amazon.in/dp/B08ABC789
```

Then run:

```bash
python amazon_exporter.py --input "electronics.txt" --output "electronics_data.xlsx"
```

## Output Format

### Excel Structure

| ASIN | Product Title | Price | Rating | # Ratings | Bestseller Rank | Main Category | Sub Category |
|------|---------------|-------|--------|-----------|-----------------|---------------|--------------|
| B0CZL9BM4S | Logitech Mouse | 2,395 | 4.2 | 2,214 ratings | #538 | Electronics | Peripherals |

### Features of Generated Excel

- Header with timestamp
- Color-coded column headers
- Properly sized columns for readability
- Frozen header rows for easy scrolling
- Professional formatting

## Command Line Arguments

```bash
python amazon_exporter.py --input <URL_OR_FILE> [--output <OUTPUT_FILE>]
```

| Argument | Required | Description |
|----------|----------|-------------|
| `--input` | Yes | Single URL or file path containing URLs (one per line) |
| `--output` | No | Output Excel file path (auto-generated if not specified) |

## Troubleshooting

| Issue | Solution |
|-------|----------|
| ModuleNotFoundError | Run `pip install -r requirements.txt` |
| ChromeDriver not found | webdriver-manager will auto-download it |
| Connection timeout | Check internet connection, Amazon may be rate-limiting |
| Empty data extracted | Amazon may have changed page structure, try again later |
| Rating shows N/A | Product may not have ratings yet |

## Tips

- Add delays between multiple URLs to avoid blocking
- Check the directory where you run the script for output files
- URL must be Amazon India (`amazon.in`)
- Some products may have missing fields (N/A) if not available
- Script runs in headless mode (no browser window)

## File Naming Convention

```
amazon_products_YYYYMMDD_HHMMSS.xlsx

Example:
amazon_products_20251111_114956.xlsx
```

## URL Format

Must be direct Amazon product URLs:

Valid:
```
https://www.amazon.in/dp/B0CZL9BM4S
https://www.amazon.in/Logitech-Signature-Wireless-Mouse/dp/B0CZL9BM4S
```

Invalid:
```
https://www.amazon.in/s?k=laptop
https://amazon.in/some-product
```

## Performance

- Single URL: ~10-15 seconds
- 10 URLs: ~2-3 minutes
- 50 URLs: ~8-12 minutes

Speed depends on your internet connection and Amazon's response time.

## Limitations

- Only works with Amazon India (amazon.in)
- Requires active internet connection
- Some data may be N/A if not available on product page
- Amazon may rate-limit or block automated requests

## License

This project is licensed under the MIT License.

## Disclaimer

This tool is for educational and personal research purposes only. Please respect Amazon's Terms of Service and robots.txt. Use responsibly and avoid excessive scraping.

- Do not use for commercial purposes without permission
- Respect rate limits and implement delays between requests
- Use for personal research and learning only
- The authors are not responsible for misuse

## Contact

For issues or suggestions, please reach out.

## Contributing

Contributions are welcome! Feel free to submit issues and pull requests.

---

**Last Updated**: November 2025