from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    ElementClickInterceptedException, TimeoutException, StaleElementReferenceException
)
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
# options.add_argument('--headless=new')
options.add_argument('--disable-gpu')
options.add_argument('--window-size=1920,1080')
options.add_argument('--lang=ar')
options.add_argument('--disable-blink-features=AutomationControlled')  # Try to avoid detection
options.add_experimental_option('excludeSwitches', ['enable-automation'])  # Try to avoid detection
options.add_experimental_option('useAutomationExtension', False)  # Try to avoid detection
driver = webdriver.Chrome(options=options)
driver.execute_cdp_cmd('Network.setUserAgentOverride', {"userAgent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'})
wait = WebDriverWait(driver, 10)

# === Input Excel ===
input_excel = "btech-targets.xlsx"
df = pd.read_excel(input_excel, header=1)
df.columns = df.columns.str.strip()
df = df.dropna(subset=["Category", "URL"])
category_links = list(zip(df["Category"], df["URL"]))

# === Output Directory ===
output_dir = "btech-outputs"
os.makedirs(output_dir, exist_ok=True)

# === Excel Styling ===
def style_excel_file(path, category_name, date_str):
    wb = load_workbook(path)
    ws = wb.active

    header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    body_font = Font(color="000000")
    center_align = Alignment(horizontal="center", vertical="center")
    border = Border(bottom=Side(border_style="thin", color="000000"))

    # Insert a new row at the top for the merged header
    ws.insert_rows(1)
    # Merge the first 6 columns in the first row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    merged_cell = ws.cell(row=1, column=1)
    merged_cell.value = f"Btech {category_name} {date_str}"
    merged_cell.font = header_font
    merged_cell.fill = header_fill
    merged_cell.alignment = center_align

    # Style the rest of the header row (now row 2) and body
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = center_align
            cell.border = border
            if cell.row == 2:
                cell.fill = header_fill
                cell.font = header_font
            else:
                cell.font = body_font

    for column in ws.columns:
        col_idx = column[0].column
        col_letter = get_column_letter(col_idx)
        header = str(column[1].value).strip() if len(column) > 1 and column[1].value else ""
        if header == "Product URL":
            ws.column_dimensions[col_letter].width = 30
            for cell in column:
                if cell.row > 1:
                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=False,
                        shrink_to_fit=True
                    )
        else:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column[1:])
            ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(path)

# === Helpers ===
def normalize_price(price_text):
    return int(price_text.replace(",", "").strip()) if price_text else None

def extract_sku(name):
    if not name:
        return ""

    name = name.upper().replace("\u200f", "")  # remove RTL char

    # Regex: match blocks of 3+ alphanum (optionally separated by space, dash, plus), at end or after dash
    pattern = r'(?:-\s*)?([A-Z0-9][A-Z0-9\s\+\-]{2,})$'
    matches = re.findall(pattern, name)
    # Fallback: also match any block of 3+ alphanum (with optional spaces/pluses/dashes)
    if not matches:
        pattern2 = r'([A-Z0-9][A-Z0-9\s\+\-]{2,})'
        matches = re.findall(pattern2, name)
    # Filter: must have at least 1 letter and at least 3 chars
    candidates = [m.strip() for m in matches if any(c.isalpha() for c in m) and len(m.strip()) >= 3]
    return candidates[-1] if candidates else ""

def normalize_sku(sku):
    return re.sub(r'[^a-zA-Z0-9]', '', sku).lower() if sku else ""

def extract_total_expected_products(driver):
    try:
        # Find the initial count from the header
        el = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.my-product-title span#product-search-item-count"))
        )
        initial_count = int(el.text.strip())
        
        # Verify count from the amscroll message
        try:
            scroll_msg = driver.find_element(By.CSS_SELECTOR, "section.plpnew_wrap div.amscroll-count-message")
            if scroll_msg:
                # Extract numbers from text like "شاهدت 5 من 10 منتج"
                numbers = [int(num) for num in re.findall(r'\d+', scroll_msg.text)]
                if numbers and len(numbers) >= 2:
                    confirmed_count = max(numbers)
                    if confirmed_count == initial_count:
                        return initial_count
                    else:
                        print(f"⚠️ Count mismatch: header shows {initial_count}, scroll message shows {confirmed_count}")
                        return max(initial_count, confirmed_count)
        except:
            pass
        
        return initial_count
    except Exception as e:
        print(f"⚠️ Could not extract expected count: {str(e)}")
        return None

# === Scrape Each Category ===
print("🚀 Starting Btech Scraper")
for category, url in category_links:
    print(f"\n➡️ Category: {category} | URL: {url}")
    driver.get(url)
    time.sleep(5)  # Initial longer wait for full page load
    
    # Wait for the main product container to be present and visible
    try:
        # Try different possible container selectors
        container_selectors = [
            "div.products.wrapper.grid.products-grid",
            "div.products.list.items.product-items",
            "ol.products.list.items.product-items"
        ]
        
        container_found = False
        for selector in container_selectors:
            try:
                print(f"Trying to find container with selector: {selector}")
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                container_found = True
                print(f"✅ Found container with selector: {selector}")
                break
            except:
                print(f"❌ Container not found with selector: {selector}")
                continue
        
        if not container_found:
            print("⚠️ Could not find any product container, trying to save page source for debugging")
            with open(f"debug_page_{safe_category}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            raise Exception("No product container found")
            
    except Exception as e:
        print(f"⚠️ Page did not load properly: {str(e)}, skipping category")
        continue
        
    # Scroll down a bit to trigger any lazy loading
    driver.execute_script("window.scrollBy(0, 300)")
    time.sleep(2)

    safe_category = re.sub(r"[^\w\s-]", "", category).replace(" ", "_")
    date_str = datetime.now().strftime("%Y-%m-%d")
    output_path = os.path.join(output_dir, f"btech_{safe_category}_{date_str}.xlsx")

    expected_total = extract_total_expected_products(driver)
    max_scrape_limit = expected_total + 2 if expected_total else float("inf")
    print(f"📊 Expected products: {expected_total} | Max scrape: {max_scrape_limit}")

    previous_count = -1
    attempt = 0
    max_attempts = 40

    while attempt < max_attempts:
        time.sleep(2)
        products = driver.find_elements(By.CSS_SELECTOR, "div.plpContentWrapper")
        current_count = len(products)
        print(f"🟨 Products loaded: {current_count}")
        print(f"🟨 Products loaded: {current_count}")

        if expected_total and current_count >= expected_total + 2:
            print("🛑 Expected count reached.")
            break
        if current_count == previous_count:
            # Double check if we really need to break
            try:
                load_more_visible = driver.find_element(By.CSS_SELECTOR, "div.amscroll-load-button").is_displayed()
                if not load_more_visible:
                    print("✅ No more products to load.")
                    break
            except:
                print("✅ No Load More button found.")
                break
        previous_count = current_count

        try:
            # Look for the load more button with its complete attributes
            load_more_btn = wait.until(
                EC.presence_of_element_located((
                    By.CSS_SELECTOR, 
                    "div.amscroll-load-button.btn-outline.primary.medium[amscroll_type='after']"
                ))
            )
            
            # Verify button is visible and has expected Arabic text
            if not load_more_btn.is_displayed():
                print("ℹ️ Load More button not visible.")
                break
            
            button_text = load_more_btn.text.strip()
            if not button_text or not any(text in button_text for text in ["عرض", "المزيد"]):
                print(f"⚠️ Load More button has unexpected text: {button_text}")
                break
                
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", load_more_btn)
            time.sleep(1.5)  # Give more time for any animations
            
            try:
                load_more_btn.click()
                print("🔁 Clicked Load More (expecting +30 products)")
            except ElementClickInterceptedException:
                print("⚠️ Intercepted, retrying with JS")
                driver.execute_script("arguments[0].click();", load_more_btn)
            
            # Wait for new products to load
            time.sleep(2)
        except TimeoutException:
            print("ℹ️ Load More button not found.")
            break
        attempt += 1

    # === Parse Products ===
    data = []
    wrappers = driver.find_elements(By.CSS_SELECTOR, "a.listingWrapperSection")
    print(f"Found {len(wrappers)} products to parse")

    for wrapper in wrappers:
        try:
            title_el = wrapper.find_elements(By.CSS_SELECTOR, "h2.plpTitle")
            if not title_el or not title_el[0].text.strip():
                continue
            title = title_el[0].text.strip()
            
            new_price_el = wrapper.find_elements(By.CSS_SELECTOR, "span.special-price span.price-wrapper")
            old_price_el = wrapper.find_elements(By.CSS_SELECTOR, "span.old-price.was-price span.price-wrapper")

            new_price = normalize_price(new_price_el[0].text) if new_price_el else None
            old_price = normalize_price(old_price_el[0].text) if old_price_el else None

            product_url = wrapper.get_attribute("href")
            
            if not new_price:
                print(f"⚠️ Could not find any price for: {title}")
            
            product_code = extract_sku(title)
            normalized_code = normalize_sku(product_code)

            data.append({
                "Item Name": title,
                "Old Price": old_price,
                "New Price": new_price,
                "Product Code": product_code,
                "Normalized Code": normalized_code,
                "Product URL": product_url
            })
        except Exception as e:
            print("❌ Skipped product:", e)

    # === Save Output ===
    if data:
        # Limit to expected_total + 2 if expected_total is available
        limit = expected_total + 2 if expected_total else None
        if limit and len(data) > limit:
            print(f"⚠️ Scraped {len(data)} products, but expected only {expected_total}. Limiting to {limit}.")
            data = data[:limit]
        df_out = pd.DataFrame(data)
        # Reorder columns so Product URL is last (6th)
        ordered_cols = [
            "Item Name",
            "Old Price",
            "New Price",
            "Product Code",
            "Normalized Code",
            "Product URL"
        ]
        df_out = df_out[ordered_cols]
        df_out.to_excel(output_path, index=False, engine='openpyxl')
        style_excel_file(output_path, category, datetime.now().strftime("%y-%m-%d"))
        print(f"✅ Saved {len(data)} products to {output_path}")
    else:
        print("⚠️ No data extracted.")

# === Done ===
driver.quit()
print("🏁 All categories processed for Btech.")
