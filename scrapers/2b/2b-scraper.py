from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")

# === Input Excel ===
input_excel = "2b-targets.xlsx"
df = pd.read_excel(input_excel, header=1)
df.columns = df.columns.str.strip()
df = df.dropna(subset=["Category", "URL"])
category_links = list(zip(df["Category"], df["URL"]))

# === Output Directory ===
output_dir = "2b-outputs"
os.makedirs(output_dir, exist_ok=True)

# === Excel Styling ===
def style_excel_file(path, category_name, date_str):
    wb = load_workbook(path)
    ws = wb.active

    header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    body_font = Font(color="000000")
    center_align = Alignment(horizontal="center", vertical="center")
    border = Border(bottom=Side(border_style="thin", color="000000"))

    # Insert a new row at the top for the merged header
    ws.insert_rows(1)
    # Merge the first 6 columns in the first row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    merged_cell = ws.cell(row=1, column=1)
    merged_cell.value = f"2B {category_name} {date_str}"
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
                        horizontal="left",
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
    if not name:
        return ""
    # Remove RTL marks and normalize whitespace
    name = re.sub(r'[\u200e\u200f\u202a-\u202e]', '', name)
    name = ' '.join(name.split())

    # Regex: match blocks of at least 2 allowed chars, must contain at least one letter
    pattern = r'([A-Z0-9 \-/_+()]{2,})'
    matches = re.finditer(pattern, name, re.IGNORECASE)

    candidates = []
    for match in matches:
        candidate = match.group().strip()
        # Must contain at least one letter
        if re.search(r'[A-Z]', candidate, re.IGNORECASE):
            candidates.append(candidate)

    return candidates[-1] if candidates else ""

def normalize_sku(sku):
    return re.sub(r'[^a-zA-Z0-9]', '', sku).lower() if sku else ""

# === Start Browser ===
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)

# === Main Loop ===
print("🚀 Starting 2B Scraper...")
for category, url in category_links:
    print(f"\n➡️ Scraping Category: {category}\n🔗 {url}")
    driver.get(url)
    time.sleep(2)

    # Scroll to load all products
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

    products = driver.find_elements(By.CSS_SELECTOR, "div.product-item-info")
    print(f"✅ Found {len(products)} products.")

    data = []
    for product in products:
        try:
            title_el = product.find_element(By.CSS_SELECTOR, "a.product-item-link")
            title = title_el.text.strip()
            product_url = title_el.get_attribute("href").strip()

            # Extract new price
            try:
                new_price_el = product.find_element(By.CSS_SELECTOR, ".special-price .price")
                new_price = normalize_price(new_price_el.text)
            except:
                try:
                    new_price_el = product.find_element(By.CSS_SELECTOR, ".price-box .price")
                    new_price = normalize_price(new_price_el.text)
                except:
                    new_price = None

            # Extract old price
            try:
                old_price_el = product.find_element(By.CSS_SELECTOR, ".old-price .price")
                old_price = normalize_price(old_price_el.text)
            except:
                old_price = None

            # SKU Extraction
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
            print(f"⚠️ Skipped product due to error: {e}")
            continue

    if data:
        timestamp = datetime.now().strftime("%Y-%m-%d")
        safe_category = re.sub(r"[^\w\s-]", "", category).replace(" ", "_")
        output_file = os.path.join(output_dir, f"2b_{safe_category}_{timestamp}.xlsx")
        df_out = pd.DataFrame(data)
        df_out.to_excel(output_file, index=False, engine='openpyxl')
        # Pass category and date for merged header
        style_excel_file(output_file, category, datetime.now().strftime("%y-%m-%d"))
        print(f"💾 Saved {len(data)} products to {output_file}")
    else:
        print("⚠️ No product data collected.")

# === Done ===
driver.quit()
print("🏁 All categories processed for 2B.")
print("✅ Scraping completed successfully!")
print(f"📂 Output files saved in: {output_dir}")
