# 🛒 Amazon Product Scraper

A powerful and user-friendly tool to scrape product information from Amazon India and export the data to beautifully formatted Excel files.

## ✨ Features

- 🔍 **Smart Search** - Search any product on Amazon India
- 📊 **Customizable Results** - Choose number of products to scrape (1-50)
- 📁 **Flexible Storage** - Save Excel files to any location
- 🎨 **Beautiful Excel Output** - Formatted with headers, clickable URLs, and metadata
- 🚀 **Headless Operation** - Runs in background without opening browser windows
- 💾 **Data Extraction**:
  - Product Name
  - Price
  - Rating
  - Number of Reviews
  - Product URL (clickable)
- ⏱️ **Timestamped Files** - Never overwrite previous searches
- 🖥️ **Cross-Platform** - Works on Windows, Linux, and macOS

## 📸 Screenshots

### Console Interface
```
============================================================
        🛒 Amazon India Product Scraper 🛒
============================================================

🔍 Enter product keyword to search: laptop

🔢 How many products would you like to scrape?
   Number of products: 10

📁 Save location: C:\Users\YourName\Documents

✓ Found 63 products. Extracting top 10...
[1/10] ✓ Dell Inspiron 15 Laptop...
        Price: ₹45,990 | Rating: 4.3 | Reviews: 1,234
```

### Excel Output
Beautiful, formatted spreadsheet with:
- Header with search keyword
- Timestamp
- Color-coded columns
- Clickable product URLs
- Frozen header panes

## 🚀 Quick Start

### Prerequisites

- Python 3.8 or higher
- Google Chrome or Chromium browser
- Internet connection

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/amazon-product-scraper.git
   cd amazon-product-scraper
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the scraper**
   ```bash
   python product_scrap.py
   ```

### Using the Windows EXE (No Python Required)

1. Download the latest `AmazonScraper.exe` from [Releases](https://github.com/yourusername/amazon-product-scraper/releases)
2. Double-click to run
3. Follow the on-screen prompts
4. Find your Excel file in the chosen location

## 📦 Installation Guide

### Method 1: Python Script

```bash
# Clone repository
git clone https://github.com/yourusername/amazon-product-scraper.git
cd amazon-product-scraper

# Create virtual environment (recommended)
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On Linux/Mac:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run
python product_scrap.py
```

### Method 2: Windows Executable

Download from [Releases](https://github.com/yourusername/amazon-product-scraper/releases) - no installation needed!

## 🎯 Usage

### Running the Script

```bash
python product_scrap.py
```

### Interactive Prompts

1. **Enter Search Keyword**
   ```
   🔍 Enter product keyword to search: wireless mouse
   ```

2. **Choose Number of Products**
   ```
   🔢 How many products would you like to scrape?
      Number of products: 20
   ```

3. **Select Save Location**
   ```
   📁 Where would you like to save the Excel file?
      Enter full path: C:\Users\YourName\Documents
   ```

4. **Confirm and Scrape**
   ```
   ▶ Proceed with scraping? (y/n): y
   ```

### Example Output

```
============================================================
✅ SUCCESS!
============================================================
📊 Scraped products: 20
📁 File saved to: C:\Users\YourName\Documents\wireless_mouse_amazon_20251004_143022.xlsx
📝 File name: wireless_mouse_amazon_20251004_143022.xlsx
============================================================
```

## 📋 Requirements

```txt
selenium>=4.15.0
openpyxl>=3.1.0
```

### System Requirements

- **Python**: 3.8 or higher
- **Browser**: Google Chrome or Chromium
- **OS**: Windows 10/11, Linux (Ubuntu 20.04+), macOS 10.15+
- **RAM**: Minimum 2GB recommended
- **Disk Space**: ~100MB for dependencies

## 🛠️ Advanced Usage

### Customizing the Script

Edit `product_scrap.py` to modify:

- **Default product count**: Change line with `return 10`
- **Max products limit**: Modify `if 1 <= count <= 50:`
- **Wait times**: Adjust `time.sleep(3)` for slower/faster scraping
- **User agent**: Update the user-agent string for different browsers

### Running in Silent Mode

For automation, you can modify the script to accept command-line arguments:

```python
import sys

if len(sys.argv) >= 2:
    keyword = sys.argv[1]
    product_count = int(sys.argv[2]) if len(sys.argv) >= 3 else 10
```

Then run:
```bash
python product_scrap.py "laptop" 20
```

## 🔧 Troubleshooting

### Common Issues

#### 1. "Cannot find Chrome binary"

**Solution**: Install Google Chrome or Chromium

```bash
# Ubuntu/Debian
sudo apt install chromium-browser

# Windows
# Download from: https://www.google.com/chrome/
```

#### 2. "No products found"

**Possible Causes**:
- Amazon blocking automated access (try again later)
- Network connectivity issues
- Invalid search keyword
- Page structure changed

**Solution**: 
- Wait a few minutes and retry
- Try a different keyword
- Check `debug.html` file for details

#### 3. "Module not found"

**Solution**: Install missing dependencies
```bash
pip install -r requirements.txt
```

#### 4. Slow Scraping

**Solution**: This is normal. The script waits to avoid detection and ensure data loads properly.

## 🏗️ Building from Source

### Create Windows EXE

```bash
# Install PyInstaller
pip install pyinstaller

# Build executable
pyinstaller --onefile --console --name "AmazonScraper" product_scrap.py

# Find output in: dist/AmazonScraper.exe
```

### Build Script (Automated)

**Windows** (`build.bat`):
```batch
pip install pyinstaller
pyinstaller --onefile --console --name "AmazonScraper" product_scrap.py
echo Build complete! Check dist folder.
pause
```

**Linux/Mac** (`build.sh`):
```bash
pip install pyinstaller
pyinstaller --onefile --console --name "AmazonScraper" product_scrap.py
echo "Build complete! Check dist folder."
```

## 📊 Output Format

### Excel Structure

| # | Product Name | Price | Rating | Reviews | Product URL |
|---|--------------|-------|--------|---------|-------------|
| 1 | Product Name Here | ₹1,999 | 4.5 | 1,234 | [Link] |
| 2 | Another Product | ₹2,499 | 4.2 | 567 | [Link] |

### File Naming Convention

```
{keyword}_amazon_{YYYYMMDD_HHMMSS}.xlsx

Example:
laptop_amazon_20251004_143022.xlsx
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚠️ Disclaimer

This tool is for educational purposes only. Please respect Amazon's Terms of Service and robots.txt. Use responsibly and avoid excessive scraping that could impact Amazon's servers.

- Do not use for commercial purposes without permission
- Respect rate limits and implement delays
- Use for personal research and learning only
- The authors are not responsible for misuse

## 🙏 Acknowledgments

- Built with [Selenium](https://www.selenium.dev/) for web automation
- Excel export powered by [openpyxl](https://openpyxl.readthedocs.io/)
- Inspired by the need for easy product research

## 📧 Contact

- **Issues**: [Report a bug](https://github.com/yourusername/amazon-product-scraper/issues)

## 🗺️ Roadmap

- [ ] Add support for multiple Amazon regions (US, UK, etc.)
- [ ] Export to CSV format
- [ ] Price tracking over time
- [ ] Email notifications for price drops
- [ ] GUI version with Tkinter
- [ ] Scheduled scraping with cron/Task Scheduler
- [ ] Database storage option (SQLite)
- [ ] Product comparison features

## ⭐ Star History

If you find this project useful, please consider giving it a star! ⭐

---