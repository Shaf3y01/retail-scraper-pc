from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
import time
import pandas as pd
import os
import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# === Chrome Setup ===
options = Options()
# options.add_argument('--headless=new')  # Uncomment for headless mode
options.add_argument('--disable-gpu')
options.add_argument('--window-size=1920,1080')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option('excludeSwitches', ['enable-automation'])
options.add_experimental_option('useAutomationExtension', False)
driver = webdriver.Chrome(options=options)
driver.execute_cdp_cmd('Network.setUserAgentOverride', {
    "userAgent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
})
wait = WebDriverWait(driver, 10)

# === Input Excel ===
input_excel = "rizkalla-targets.xlsx"
df = pd.read_excel(input_excel, header=1)
df.columns = df.columns.str.strip()
df = df.dropna(subset=["Category", "URL"])
category_links = list(zip(df["Category"], df["URL"]))

# === Output Directory ===
output_dir = "rizkalla-outputs"
os.makedirs(output_dir, exist_ok=True)

# === Excel Styling (Same as Btech) ===
def style_excel_file(path, category_name, date_str):
    wb = load_workbook(path)
    ws = wb.active

    header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")  # Black
    header_font = Font(color="FFFFFF", bold=True)
    body_font = Font(color="000000")
    center_align = Alignment(horizontal="center", vertical="center")
    border = Border(bottom=Side(border_style="thin", color="000000"))

    # Insert and merge header row
    ws.insert_rows(1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    merged_cell = ws.cell(row=1, column=1)
    merged_cell.value = f"Rizkalla {category_name} {date_str}"
    merged_cell.font = header_font
    merged_cell.fill = header_fill
    merged_cell.alignment = center_align

    # Style header (row 2) and body
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = center_align
            cell.border = border
            if cell.row == 2:
                cell.fill = header_fill
                cell.font = header_font
            else:
                cell.font = body_font

    # Auto column width + fixed URL width
    for column in ws.columns:
        col_idx = column[0].column
        col_letter = get_column_letter(col_idx)
        header = str(column[1].value).strip() if len(column) > 1 and column[1].value else ""

        if header == "Product URL":
            ws.column_dimensions[col_letter].width = 30
            for cell in column:
                if cell.row == 1:
                    cell.alignment = center_align
                elif cell.row == 2:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(
                        horizontal="left",
                        vertical="center",
                        wrap_text=False,
                        shrink_to_fit=True
                    )
        else:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column[1:])
            ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(path)

# === Helper Functions ===
def normalize_price(text):
    """Extracts only the integer part of the price."""
    if not text:
        return None
    match = re.search(r'[\d,]+(?=\.\d+|$)', text)
    if not match:
        return None
    cleaned = match.group(0).replace(',', '')
    return int(cleaned) if cleaned.isdigit() else None

def extract_sku(name):
    """Btech-style SKU extraction."""
    if not name:
        return ""
    name = name.upper().replace("\u200f", " ")
    pattern = r'(?:-\s*)?([A-Z0-9][A-Z0-9\s\+\-]{2,})$'
    matches = re.findall(pattern, name)
    if not matches:
        pattern2 = r'([A-Z0-9][A-Z0-9\s\+\-]{2,})'
        matches = re.findall(pattern2, name)
    candidates = [m.strip() for m in matches if any(c.isalpha() for c in m) and len(m.strip()) >= 3]
    return candidates[-1] if candidates else ""

def normalize_sku(sku):
    """Removes separators and converts to lowercase."""
    return re.sub(r'[^a-zA-Z0-9]', '', sku).lower() if sku else ""

# === Detect Page Type ===
def is_search_page(driver):
    """Returns True if URL contains 'search' or search results title is present."""
    return 'search' in driver.current_url.lower() or \
           'q=' in driver.current_url.lower() or \
           len(driver.find_elements(By.CSS_SELECTOR, "h3.search-results_title")) > 0

# === Extract Total Product Count (Fallback Logic) ===
def get_total_product_count(driver):
    """Extracts total product count from either category or search page."""
    # Try 1: div.products-showing (category pages)
    try:
        elem = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.products-showing#products-showing"))
        )
        text = elem.text.strip()
        numbers = [int(num) for num in re.findall(r'\d+', text)]
        return max(numbers) if numbers else None
    except:
        pass

    # Try 2: h3.search-results_title (search pages)
    try:
        elem = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h3.search-results_title"))
        )
        text = elem.text.strip()
        numbers = [int(num) for num in re.findall(r'\d+', text)]
        return max(numbers) if numbers else None
    except:
        pass

    return None

# === Get Product Cards Based on Page Type ===
def get_product_cards(driver):
    """Returns list of product cards using correct selector for page type."""
    if is_search_page(driver):
        # Search page: .search-results_inner > .product-product-grid > product-card
        return driver.find_elements(By.CSS_SELECTOR, ".search-results_inner > .product-product-grid > product-card")
    else:
        # Category page: #main-collection-product-grid > product-card
        return driver.find_elements(By.CSS_SELECTOR, "div#main-collection-product-grid > product-card")

# === Start Scraping ===
print("🚀 Starting Rizkalla Scraper")
for category, url in category_links:
    print(f"\n➡️ Scraping Category: {category}")
    print(f"🔗 URL: {url}")
    driver.get(url)
    time.sleep(3)

    # Detect page type
    search_mode = is_search_page(driver)
    products_per_page = 16 if search_mode else 20
    print(f"🔍 Detected mode: {'Search' if search_mode else 'Category'} ({products_per_page} products/page)")

    # Get total product count
    total_count = get_total_product_count(driver)
    if not total_count:
        print("❌ Could not determine total product count. Skipping category.")
        continue
    print(f"📊 Total products: {total_count}")

    # Calculate total pages
    total_pages = (total_count // products_per_page) + (1 if total_count % products_per_page > 0 else 0)
    print(f" totalPages: {total_pages} ({products_per_page} per page)")

    all_data = []

    for page in range(1, total_pages + 1):
        print(f"\n📄 Scraping Page {page} of {total_pages}...")

        # Wait for product container
        try:
            if search_mode:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".search-results_inner")))
            else:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div#main-collection-product-grid")))
        except TimeoutException:
            print("❌ Timeout: Product grid not found.")
            break

        # Extract product cards
        product_cards = get_product_cards(driver)
        print(f"✅ Found {len(product_cards)} product cards on page {page}.")

        page_data = []
        for card in product_cards:
            try:
                # Title & URL
                title_el = card.find_element(By.CSS_SELECTOR, "section > header > div.product-card_vendor-title > h3 > a")
                title = title_el.text.strip()
                product_url = title_el.get_attribute("href").strip()

                # Price container
                price_container = card.find_element(By.CSS_SELECTOR, "footer div.product-price")

                # New Price
                try:
                    new_price_el = price_container.find_element(By.CSS_SELECTOR, "div.price-sale")
                    new_price = normalize_price(new_price_el.text)
                except:
                    new_price = None

                # Old Price
                try:
                    old_price_el = price_container.find_element(By.CSS_SELECTOR, "del.price-compare")
                    old_price = normalize_price(old_price_el.text)
                except:
                    old_price = None

                # SKU Extraction
                product_code = extract_sku(title)
                normalized_code = normalize_sku(product_code)

                page_data.append({
                    "Item Name": title,
                    "Old Price": old_price,
                    "New Price": new_price,
                    "Product Code": product_code,
                    "Normalized Code": normalized_code,
                    "Product URL": product_url
                })

            except Exception as e:
                print(f"⚠️ Skipped product: {e}")
                continue

        all_data.extend(page_data)

        # === Go to Next Page (if not last page) ===
        if page < total_pages:
            next_page_num = page + 1
            print(f"➡️ Navigating to page {next_page_num}...")

            # Try 1: Click "Next" button (class="next")
            try:
                next_button = driver.find_element(By.CSS_SELECTOR, "div.pagination-holder ul.pagination li a.next")
                if next_button.is_displayed():
                    print("👉 Attempting to click 'Next' button...")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                    time.sleep(1)
                    try:
                        next_button.click()
                    except ElementClickInterceptedException:
                        driver.execute_script("arguments[0].click();", next_button)
                    time.sleep(3)
                    continue  # Go to next page
            except Exception as e:
                print(f"⚠️ 'Next' button not found or not clickable: {e}")

            # Try 2: Click numbered page link (fallback)
            try:
                page_link = driver.find_element(
                    By.CSS_SELECTOR,
                    f"div.pagination-holder ul.pagination li > a[href*='page={next_page_num}']"
                )
                if page_link.is_displayed():
                    print(f"👉 Falling back to clicking page {next_page_num} link...")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", page_link)
                    time.sleep(1)
                    try:
                        page_link.click()
                    except ElementClickInterceptedException:
                        driver.execute_script("arguments[0].click();", page_link)
                    time.sleep(3)
                    continue
            except Exception as e:
                print(f"⚠️ Could not navigate to page {next_page_num}: {e}")
                break

        else:
            print("🔚 Last page reached. Pagination complete.")
            break

    # Save to Excel
    if all_data:
        timestamp = datetime.now().strftime("%Y-%m-%d")
        safe_category = re.sub(r'[^\w\s-]', '', category).replace("  ", "_").strip()
        output_file = os.path.join(output_dir, f"rizkalla_{safe_category}_{timestamp}.xlsx")

        df_out = pd.DataFrame(all_data)
        df_out = df_out[[
            "Item Name",
            "Old Price",
            "New Price",
            "Product Code",
            "Normalized Code",
            "Product URL"
        ]]
        df_out.to_excel(output_file, index=False, engine='openpyxl')
        style_excel_file(output_file, category, datetime.now().strftime("%y-%m-%d"))
        print(f"💾 Saved {len(all_data)} products to {output_file}")
    else:
        print("⚠️ No product data collected.")

# === Cleanup ===
driver.quit()
print("🏁 Rizkalla scraping completed.")