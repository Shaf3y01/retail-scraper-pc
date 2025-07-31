# Raya Scraper – Documentation

A Selenium-based scraper for [rayashop.com](https://www.rayashop.com) that extracts product data from category and search pages with infinite scroll. Outputs all categories into a single Excel workbook with one sheet per category.

---

## 📁 File Structure


---

## 📥 Input: `raya-targets.xlsx`

- **Header row**: Row 2 (first row ignored)
- **Columns**:
  - `Category`: Display name (e.g., `Water Dispensers`)
  - `URL`: Full URL (e.g., `https://www.rayashop.com/ar/water-dispensers`)
- UTF-8 encoded (supports Arabic)

---

## 📤 Output: Single Workbook

- **File**: `raya-outputs/raya-all-categories_YYYY-MM-DD.xlsx`
- **One sheet per category**
- **Sheet name**: Sanitized category name (≤31 chars, safe for Excel)
- **Columns (in order)**:
  1. `Item Name`
  2. `Old Price`
  3. `New Price`
  4. `Product Code`
  5. `Normalized Code`
  6. `Product URL`

---

## 🎨 Excel Styling

- **Row 1 (Merged Header)**:
  - Text: `Raya {Category} YY-MM-DD`
  - Black fill, white bold, centered
- **Row 2 (Column Headers)**:
  - Black fill, white bold, centered
- **Body (Rows 3+)**:
  - Centered (except URL)
  - Bottom thin border
- **URL Column**:
  - Width: **30**
  - Alignment: **Left**
  - `wrap_text=False` → no overflow or line breaks
- Other columns: auto-width

---

## 🌐 Infinite Scroll Handling

- Scrolls to bottom repeatedly
- Stops when no new products load
- Breaks early if loaded count ≥ expected

---

## 🔢 Total Product Count (Dual Verification)

Extracts total from **two sources**:
1. `<h3 class="text-secondary-500">82 Product</h3>`
2. `<p class="text-secondary-400">Showing 82 out of 82</p>`

Uses the **largest number** from both for accuracy.

---

## 🧩 Product Extraction

| Field | Selector | Notes |
|------|---------|-------|
| **Item Name** | `p.name.clamp-text` inside `a.flex.flex-col` | Text only |
| **Product URL** | Parent `<a href="/ar/...">` | Prepend base: `https://www.rayashop.com` |
| **New Price** | `span.text-primary-500:not(.line-through)` | E.g., `6,759` |
| **Old Price** | `span.line-through` | E.g., `9,800` |
| **Product Card** | `article.ProductCard` | |

---

## 🔤 SKU Extraction (`extract_sku`)

Uses **Btech-style logic**:
- Looks for block at end or after dash
- Must contain ≥1 letter and ≥3 chars
- Returns last valid match

**Example**:  
`"موزع مياه تورنيدو 3 حنفيات - MAR-2270D"` → `"MAR-2270D"`

---

## 🔤 SKU Normalization (`normalize_sku`)

Removes all non-alphanumeric chars and converts to lowercase:
```python
re.sub(r'[^a-zA-Z0-9]', '', sku).lower()