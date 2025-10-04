import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime

def get_save_location():
    """
    Get save location from user, with default to current directory
    """
    print("\n📁 Where would you like to save the Excel file?")
    print("   (Press Enter for current directory)")
    save_path = input("   Enter full path: ").strip()
    
    if not save_path:
        save_path = os.getcwd()
    
    # Validate path
    if not os.path.exists(save_path):
        print(f"\n⚠ Path doesn't exist: {save_path}")
        create = input("   Create this directory? (y/n): ").strip().lower()
        if create == 'y':
            try:
                os.makedirs(save_path, exist_ok=True)
                print(f"✓ Created directory: {save_path}")
            except Exception as e:
                print(f"❌ Error creating directory: {e}")
                print("   Using current directory instead.")
                save_path = os.getcwd()
        else:
            print("   Using current directory instead.")
            save_path = os.getcwd()
    
    return save_path

def get_product_count():
    """
    Get number of products to scrape from user
    """
    while True:
        print("\n🔢 How many products would you like to scrape?")
        print("   (Enter a number between 1-50, default is 10)")
        count_input = input("   Number of products: ").strip()
        
        if not count_input:
            return 10
        
        try:
            count = int(count_input)
            if 1 <= count <= 50:
                return count
            else:
                print("   ⚠ Please enter a number between 1 and 50")
        except ValueError:
            print("   ⚠ Please enter a valid number")

def scrape_amazon_products(keyword, max_products=10):
    """
    Scrape Amazon India for products based on keyword
    Returns list of products with title, price, rating, and URL
    """
    options = webdriver.ChromeOptions()
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-extensions')
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    # Disable automation flags
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    driver = None
    try:
        print("\n🔧 Initializing Chrome WebDriver in headless mode...")
        driver = webdriver.Chrome(options=options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        # Build Search URL
        search_url = f"https://www.amazon.in/s?k={keyword.replace(' ', '+')}"
        print(f"🌐 Navigating to: {search_url}")
        driver.get(search_url)
        
        # Wait for search results to load
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-component-type='s-search-result']")))
        
        time.sleep(3)  # Additional wait for dynamic content
        
        products = []
        results = driver.find_elements(By.XPATH, "//div[@data-component-type='s-search-result']")
        
        total_found = len(results)
        products_to_scrape = min(max_products, total_found)
        
        print(f"✓ Found {total_found} products. Extracting top {products_to_scrape}...")
        print("-" * 60)
        
        for idx, result in enumerate(results[:max_products], 1):
            product_data = {}
            
            # Extract Title - try multiple selectors
            try:
                title_element = result.find_element(By.CSS_SELECTOR, "h2 a span")
                product_data['title'] = title_element.text.strip()
            except:
                try:
                    title_element = result.find_element(By.CSS_SELECTOR, "h2.a-size-mini span")
                    product_data['title'] = title_element.text.strip()
                except:
                    try:
                        title_element = result.find_element(By.XPATH, ".//h2//span")
                        product_data['title'] = title_element.text.strip()
                    except:
                        product_data['title'] = "N/A"
            
            # Extract Price - try multiple methods
            try:
                price_whole = result.find_element(By.CSS_SELECTOR, "span.a-price-whole").text
                try:
                    price_fraction = result.find_element(By.CSS_SELECTOR, "span.a-price-fraction").text
                    product_data['price'] = f"₹{price_whole}{price_fraction}"
                except:
                    product_data['price'] = f"₹{price_whole}"
            except:
                try:
                    price = result.find_element(By.CSS_SELECTOR, "span.a-price span.a-offscreen").text
                    product_data['price'] = price
                except:
                    try:
                        price = result.find_element(By.XPATH, ".//span[@class='a-price']//span[@class='a-offscreen']").text
                        product_data['price'] = price
                    except:
                        product_data['price'] = "N/A"
            
            # Extract Rating
            try:
                rating = result.find_element(By.CSS_SELECTOR, "span.a-icon-alt").text
                product_data['rating'] = rating.split()[0] if rating else "N/A"
            except:
                try:
                    rating = result.find_element(By.XPATH, ".//i[contains(@class, 'a-icon-star-small')]//span").text
                    product_data['rating'] = rating.split()[0] if rating else "N/A"
                except:
                    product_data['rating'] = "N/A"
            
            # Extract Review Count
            try:
                reviews = result.find_element(By.CSS_SELECTOR, "span.a-size-base.s-underline-text").text
                product_data['reviews'] = reviews
            except:
                try:
                    reviews = result.find_element(By.XPATH, ".//span[contains(@aria-label, 'ratings')]").text
                    product_data['reviews'] = reviews if reviews else "N/A"
                except:
                    product_data['reviews'] = "N/A"
            
            # Extract Product URL
            try:
                link_element = result.find_element(By.CSS_SELECTOR, "h2 a")
                product_url = link_element.get_attribute('href')
                if '?' in product_url and 'amazon.in' in product_url:
                    base_url = product_url.split('?')[0]
                    product_data['url'] = base_url
                else:
                    product_data['url'] = product_url
            except:
                try:
                    link_element = result.find_element(By.XPATH, ".//h2//a")
                    product_url = link_element.get_attribute('href')
                    if '?' in product_url:
                        product_url = product_url.split('?')[0]
                    product_data['url'] = product_url
                except:
                    product_data['url'] = "N/A"
            
            # Progress indicator
            if product_data['title'] != "N/A":
                title_preview = product_data['title'][:50] + "..." if len(product_data['title']) > 50 else product_data['title']
                print(f"[{idx:2d}/{products_to_scrape}] ✓ {title_preview}")
                print(f"        Price: {product_data['price']} | Rating: {product_data['rating']} | Reviews: {product_data['reviews']}")
            
            # Add product if we got at least title OR url
            if product_data['title'] != "N/A" or product_data['url'] != "N/A":
                products.append(product_data)
        
        print("-" * 60)
        
        # If no products found, save page source for debugging
        if not products:
            debug_file = 'debug.html'
            print(f"\n⚠ Debug: Saving page source to {debug_file} for inspection...")
            with open(debug_file, 'w', encoding='utf-8') as f:
                f.write(driver.page_source)
            print(f"✓ Page source saved to: {os.path.abspath(debug_file)}")
        
        return products
    
    except Exception as e:
        print(f"\n❌ Error during scraping: {str(e)}")
        import traceback
        traceback.print_exc()
        return []
    
    finally:
        if driver:
            driver.quit()

def save_to_excel(products, filename, keyword):
    """
    Save products to Excel with formatting
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Amazon Products"
    
    # Define styles
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    
    # Add metadata
    ws.merge_cells('A1:F1')
    meta_cell = ws['A1']
    meta_cell.value = f"Amazon India Search Results - '{keyword}'"
    meta_cell.font = Font(bold=True, size=14)
    meta_cell.alignment = Alignment(horizontal="center", vertical="center")
    meta_cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    
    ws.merge_cells('A2:F2')
    date_cell = ws['A2']
    date_cell.value = f"Scraped on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    date_cell.alignment = Alignment(horizontal="center", vertical="center")
    date_cell.font = Font(italic=True, size=10)
    
    # Add empty row
    ws.append([])
    
    # Add headers
    headers = ["#", "Product Name", "Price", "Rating", "Reviews", "Product URL"]
    ws.append(headers)
    
    # Style headers
    for cell in ws[4]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Add product data
    for idx, prod in enumerate(products, 1):
        ws.append([
            idx,
            prod.get('title', 'N/A'),
            prod.get('price', 'N/A'),
            prod.get('rating', 'N/A'),
            prod.get('reviews', 'N/A'),
            prod.get('url', 'N/A')
        ])
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 80
    
    # Set row heights
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[4].height = 20
    
    # Make URLs clickable
    for row in range(5, len(products) + 5):
        cell = ws.cell(row=row, column=6)
        if cell.value and cell.value != "N/A":
            cell.hyperlink = cell.value
            cell.font = Font(color="0563C1", underline="single")
    
    # Freeze panes (keep headers visible)
    ws.freeze_panes = 'A5'
    
    # Save file
    wb.save(filename)

def main():
    """
    Main function to run the scraper
    """
    print("=" * 60)
    print("        🛒 Amazon India Product Scraper 🛒")
    print("=" * 60)
    print("\nThis tool scrapes product information from Amazon India")
    print("and exports the data to an Excel file.")
    print("=" * 60)
    
    # Get search keyword
    keyword = input("\n🔍 Enter product keyword to search: ").strip()
    
    if not keyword:
        print("❌ Error: Keyword cannot be empty!")
        input("\nPress Enter to exit...")
        return
    
    # Get number of products
    product_count = get_product_count()
    
    # Get save location
    save_path = get_save_location()
    
    # Confirm settings
    print("\n" + "=" * 60)
    print("📋 SCRAPING CONFIGURATION:")
    print(f"   Keyword: {keyword}")
    print(f"   Products to scrape: {product_count}")
    print(f"   Save location: {save_path}")
    print("=" * 60)
    
    proceed = input("\n▶ Proceed with scraping? (y/n): ").strip().lower()
    if proceed != 'y':
        print("\n⚠ Scraping cancelled by user.")
        input("Press Enter to exit...")
        return
    
    print("\n" + "=" * 60)
    print(f"🔍 Searching for '{keyword}' on Amazon India...")
    print("=" * 60)
    
    # Scrape products
    products = scrape_amazon_products(keyword, product_count)
    
    if products:
        # Create safe filename
        safe_keyword = "".join(c if c.isalnum() or c in (' ', '_') else '_' for c in keyword)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_name = f"{safe_keyword.replace(' ', '_')}_amazon_{timestamp}.xlsx"
        full_path = os.path.join(save_path, excel_name)
        
        print(f"\n💾 Saving data to Excel...")
        save_to_excel(products, full_path, keyword)
        
        print("\n" + "=" * 60)
        print("✅ SUCCESS!")
        print("=" * 60)
        print(f"📊 Scraped products: {len(products)}")
        print(f"📁 File saved to: {full_path}")
        print(f"📝 File name: {excel_name}")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("❌ SCRAPING FAILED")
        print("=" * 60)
        print("No products found! This could be due to:")
        print("  1. Amazon blocking automated access")
        print("  2. Network connectivity issues")
        print("  3. Page structure has changed")
        print("  4. Invalid search keyword")
        print("\n💡 Tips:")
        print("  - Try a different keyword")
        print("  - Check your internet connection")
        print("  - Try again after a few minutes")
        print("=" * 60)
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠ Scraping interrupted by user.")
        input("Press Enter to exit...")
    except Exception as e:
        print(f"\n❌ Unexpected error: {str(e)}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")