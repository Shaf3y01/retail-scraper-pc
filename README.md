# Price Intelligence Project

This project automates price monitoring across major Egyptian electronics retailers: **2B, Btech, and Raneen**. Using Selenium-based scrapers, it extracts product data (name, price, SKU, URL) from targeted category pages and saves structured Excel outputs with professional styling.

Each scraper handles dynamic content loading via infinite scroll or "Load More" buttons, robustly parses pricing (new/old), and applies intelligent SKU extraction from product titles using regex. SKUs are normalized for cross-platform matching.

The **comparison tools** (`dynamic-pc-long.py`, `dynamic-pc-short.py`) consolidate data by:
- Matching products via normalized SKUs (exact and fuzzy)
- Calculating confidence scores using `rapidfuzz`
- Identifying the lowest price per product
- Generating categorized outputs: matched, weak-matched, and unmatched

Outputs are exported to Excel with auto-adjusted columns, merged headers, and conditional formatting for clarity.

Designed for scalability and maintainability, this system enables data-driven pricing decisions, competitive analysis, and market intelligence with minimal manual intervention.

**Tech Stack**: Python, Selenium, Pandas, OpenPyXL, RapidFuzz
