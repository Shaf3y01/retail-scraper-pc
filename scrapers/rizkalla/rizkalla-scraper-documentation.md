
---

## 📥 Input: `rizkalla-targets.xlsx`

- **Header row**: Row 2 (first row ignored)
- **Columns**:
  - `Category`: Human-readable name (e.g., `Blenders`)
  - `URL`: Full URL (e.g., `https://riz.shop/collections/blenders` or `https://riz.shop/search?q=خلاط`)

> Supports Arabic characters. UTF-8 encoding required.

---

## 📤 Output: `rizkalla-outputs/`

- **File naming**: `rizkalla_{Category}_{YYYY-MM-DD}.xlsx`
- **Columns (in order)**:
  1. `Item Name`
  2. `Old Price`
  3. `New Price`
  4. `Product Code`
  5. `Normalized Code`
  6. `Product URL`

---

## 🎨 Excel Styling

- **Header (Row 1)**: Merged, black fill, white bold, centered
- **Column Headers (Row 2)**: Black fill, white bold, **centered**
- **Body**: Centered (except URL), black text, bottom border
- **URL Column**: Width = 30, left-aligned, no text wrap
- Auto column width for all others

> Matches Btech scraper style for consistency.

---

## 🔍 Detection & Logic

### Page Type Detection
Automatically detects:
- **Category pages** → 20 products/page
- **Search pages** → 16 products/page

Uses URL (`search`, `q=`) and DOM (`h3.search-results_title`) to decide.

---

## 🧮 Total Product Count

Extracts total count from:
- **Category**: `<div class="products-showing">عرض 1 - 20 من 31 عنصر</div>`
- **Search**: `<h3 class="search-results_title">نتائج البحث: 123 نتيجة بحث عن "خلاط"</h3>`

Finds all numbers, returns the largest.

---

## 📦 Product Grid Selectors

| Page Type | Grid Selector | Card Selector |
|---------|-------------|-------------|
| Category | `div#main-collection-product-grid` | `> product-card` |
| Search | `.search-results_inner > .product-product-grid` | `> product-card` |

---

## 🧩 Data Extraction

| Field | Selector | Notes |
|------|---------|-------|
| **Item Name** | `section > header > div.product-card_vendor-title > h3 > a` | Text only |
| **Product URL** | Same as above | `href` attribute |
| **New Price** | `footer div.product-price div.price-sale` | Discounted price |
| **Old Price** | `footer div.product-price del.price-compare` | Original price |
| **Product Code** | From title | See SKU Logic |
| **Normalized Code** | Cleaned SKU | Alphanumeric only, lowercase |

---

## 🔤 SKU Extraction (`extract_sku`)

Uses **Btech-style logic**:
1. Looks for block at end of title or after dash
2. Falls back to any block of 3+ alphanum chars
3. Must contain at least one letter
4. Returns last valid match

**Example**:  
`"خلاط يدوي HM-120T"` → `"HM-120T"`

---

## 🔤 SKU Normalization (`normalize_sku`)

Removes all separators and converts to lowercase:
```python
re.sub(r'[^a-zA-Z0-9]', '', sku).lower()