from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, ElementClickInterceptedException, StaleElementReferenceException
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
# options.add_argument('--headless=new')  # Uncomment for headless mode
options.add_argument('--disable-gpu')
options.add_argument('--window-size=1920,1080')
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)

# === Input Excel ===
input_excel = "raneen-targets.xlsx"
df = pd.read_excel(input_excel, header=1)
df.columns = df.columns.str.strip()
df = df.dropna(subset=["Category", "URL"])
category_links = list(zip(df["Category"], df["URL"]))
# === Output Directory ===
output_dir = "raneen-outputs"
os.makedirs(output_dir, exist_ok=True)

# === Excel Styling ===
def style_excel_file(path, category_name, date_str):
    wb = load_workbook(path)
    ws = wb.active

    header_fill = PatternFill(start_color="990000", end_color="990000", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    body_font = Font(color="000000")
    center_align = Alignment(horizontal="center", vertical="center")
    border = Border(bottom=Side(border_style="thin", color="000000"))

    # Insert a new row at the top for the merged header
    ws.insert_rows(1)
    # Merge the first 6 columns in the first row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    merged_cell = ws.cell(row=1, column=1)
    merged_cell.value = f"Raneen {category_name} {date_str}"
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
def normalize_price(text):
    if not text:
        return None
    return int(re.sub(r"[^\d]", "", text))

def extract_sku(name):
    """
    Extracts the last (rightmost) SKU-like block from a product name.
    Criteria:
    - At least 2 characters long
    - Contains at least one letter (A-Z, case-insensitive)
    - May contain letters, numbers, spaces, dashes (-), slashes (/), underscores (_), plus signs (+), and parentheses ((, ))
    - Ignores RTL marks and extra whitespace
    """
    if not name:
        return ""
    # Remove RTL marks and normalize whitespace
    cleaned = re.sub(r'[\u200e\u200f\u202a-\u202e]', '', name)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    # Regex: match blocks of allowed characters, at least 2 chars, at least one letter
    pattern = r'([A-Za-z0-9 \-/_+()]{2,})'
    matches = re.findall(pattern, cleaned)
    # Filter: must contain at least one letter
    matches = [m.strip() for m in matches if re.search(r'[A-Za-z]', m)]
    return matches[-1] if matches else ""

def normalize_sku(sku):
    # Remove all separators and lowercase
    return re.sub(r'[\-_/\\\.\(\)\s]', '', sku).lower() if sku else ""

# === Start Scraping ===
print("🚀 Starting Raneen Scraper")
for category, url in category_links:
    print(f"\n➡️ Scraping Category: {category} | URL: {url}")
    driver.get(url)
    time.sleep(2)

    # Load all products using "Load More"
    prev_count = -1
    same_count_repeats = 0
    max_repeats = 3

    while same_count_repeats < max_repeats:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        products = driver.find_elements(By.CSS_SELECTOR, "div.product-item-info")
        current_count = len(products)
        print(f"🔍 Products loaded so far: {current_count}")
        if current_count == prev_count:
            same_count_repeats += 1
        else:
            same_count_repeats = 0
            prev_count = current_count

    product_cards = driver.find_elements(By.CSS_SELECTOR, "div.product-item-info")
    print(f"✅ Total products loaded: {len(product_cards)}")

    data = []
    for card in product_cards:
        try:
            title_el = card.find_element(By.CSS_SELECTOR, "a.product-item-link")
            title = title_el.text.strip()
            product_url = title_el.get_attribute("href")

            # === Final Unified Price Logic (Handles all formats) ===
            try:
                price_box = card.find_element(By.CSS_SELECTOR, ".price-box.price-final_price")

                new_price = old_price = None

                # Case 1: Discounted (special + old)
                special_price_els = price_box.find_elements(By.CSS_SELECTOR, ".special-price .price-wrapper")
                old_price_els = price_box.find_elements(By.CSS_SELECTOR, ".old-price .price-wrapper")

                if special_price_els:
                    new_price = normalize_price(special_price_els[0].text)
                    if old_price_els:
                        old_price = normalize_price(old_price_els[0].text)
                else:
                    # Case 2: Regular price with wrapper
                    regular_price_els = price_box.find_elements(By.CSS_SELECTOR, ".price-container .price-wrapper")
                    if regular_price_els:
                        new_price = normalize_price(regular_price_els[0].text)
                    else:
                        # Case 3: Raw text fallback (no .price-wrapper)
                        current_price_span = price_box.find_elements(By.CSS_SELECTOR, ".current-price")
                        old_price_span = price_box.find_elements(By.CSS_SELECTOR, ".old-price")

                        if current_price_span:
                            new_price = normalize_price(current_price_span[0].text)
                        if old_price_span:
                            old_price = normalize_price(old_price_span[0].text)

            except Exception as e:
                print("⚠️ Failed to extract price:", e)
                new_price = None
                old_price = None

            product_code = extract_sku(title)
            normalized_code = normalize_sku(product_code)

            data.append({
                "Item Name": title,
                "Old Price": old_price,
                "New Price": new_price,
                "Product Code": product_code,
                "Normalized Code": normalized_code,
                "Product URL": product_url  # Product URL is now last
            })

        except Exception as e:
            print("⚠️ Skipping product due to error:", e)

    # Save to Excel
    if data:
        timestamp = datetime.now().strftime("%Y-%m-%d")
        safe_category = re.sub(r"[^\w\s-]", "", category).replace(" ", "_")
        output_file = os.path.join(output_dir, f"raneen_{safe_category}_{timestamp}.xlsx")
        df_out = pd.DataFrame(data)
        df_out.to_excel(output_file, index=False, engine="openpyxl")
        style_excel_file(output_file, category, datetime.now().strftime("%y-%m-%d"))
        print(f"💾 Saved {len(data)} products to {output_file}")
    else:
        print("⚠️ No product data extracted.")

# === Done ===
driver.quit()
print("🏁 All categories processed for Raneen.")
print("✅ Scraping completed successfully!")
print(f"📂 Output files saved in: {output_dir}")