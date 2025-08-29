# 📘 Project Best Practices

## 1. Project Purpose
Minifig Profit Tool for BrickLink sellers. It aggregates BrickLink order exports (XML/CSV), maintains inventory-level costs, evaluates how many minifigs/sets can be built from current inventory against "wanted lists", and updates a Google Sheets workbook with summaries, inventory leftovers, and order details. The domain is e-commerce inventory and cost accounting for LEGO parts, sets, and minifigs.

## 2. Project Structure
- Root
  - scripts/
    - main.py — Entrypoint that merges orders, loads data, computes buildable counts, and updates Google Sheets.
    - merge_orders.py — Merges multiple orders XML/CSV into canonical orders.xml/orders.csv (sorted by date, unique by Order ID).
    - orders.py — Loads orders from XML/CSV into an inventory structure and row list for Google Sheets.
    - wanted_lists.py — Parses wanted list XMLs and normalizes items for build logic.
    - build_logic.py — Core algorithm to determine build counts, cost, and updated inventory.
    - sheets.py — All Google Sheets update/formatting logic (Summary, Inventory, Leftover Inventory, Orders).
    - colors.py — BrickLink color ID → name map and get_color_name utility.
    - config.py — Constants, Google auth/client bootstrap, worksheet helpers, and config value retrieval.
    - credentials.json — Google service account credentials (local, not committed).
  - orders/ — BrickLink exports and the merged outputs (orders.xml, orders.csv).
  - wanted_lists/ — Input wanted list XMLs.
  - tests/
    - test_core.py — Unit tests for color mapping and build logic.

Conventions and entry points
- Run the tool from the scripts directory (or add scripts to PYTHONPATH) because intra-module imports use "from module import ..." rather than package-relative imports.
  - Example: cd scripts && python main.py
  - Alternative: From repo root, python -c "import sys; sys.path.insert(0, 'scripts'); import main; main.main()"
- Configuration is stored in the Google Sheet (Config tab). Missing numeric values prompt for input (Shipping Fee in B1, Materials Cost in B2).

## 3. Test Strategy
- Framework: Python unittest.
- Location: tests/ with files named test_*.py.
- Scope:
  - Unit tests cover pure logic first (e.g., build_logic.determine_buildable, colors.get_color_name).
  - Google Sheets interactions are not unit-tested; prefer thin wrappers and mock in future tests.
- Mocking guidelines:
  - For new tests, isolate I/O (filesystem, network, Google APIs) with mocks/stubs.
  - Use small in-memory fixtures (dicts/lists) to validate algorithm behavior.
- Coverage philosophy:
  - Prioritize correctness of core calculations: inventory updates, cost distributions, limiting factors for builds, date sorting and de-duplication.
  - Add regression tests for any discovered data quirks in XML/CSV formats.
- How to run tests:
  - From repo root: python -m unittest -v

## 4. Code Style
- Language: Python 3.
- Naming:
  - snake_case for functions/variables; UPPER_SNAKE_CASE for constants; lower_snake files.
- Typing:
  - Currently untyped. When adding new functions, consider type hints for complex dict structures.
- Docstrings & comments:
  - Public functions include short docstrings. Add examples or parameter/return descriptions for non-trivial logic.
- Error handling:
  - Fail gracefully for formatting/GSuite formatting (sheets.py uses try/except around formatting).
  - When parsing dates, try ISO format then fallback; default to datetime.min on failure (merge_orders.py).
  - For user-config values, prompt when missing (config.get_config_value).
- Data handling idioms:
  - Inventory keys:
    - Parts (Item Type 'P') are keyed by (item_id, color_id).
    - Sets/Minifigs (Item Types 'S'/'M') use key (item_id, None) to ignore color.
  - Monetary values are rounded to 2 decimals when writing to Sheets.

## 5. Common Patterns
- Separation of concerns:
  - Parsing (orders.py, wanted_lists.py), algorithm (build_logic.py), external I/O (sheets.py, merge_orders.py), orchestration (main.py).
- Fee allocation:
  - order-level fees are distributed proportionally to item totals, then merged into inventory cost basis.
- Build determination:
  - Deep copy inventory before mutation; compute limiting factors across required items; consume inventory greedily; compute aggregate cost using unit_cost.
- Google Sheets updates:
  - Write headers then data; preserve existing price column where present; use USER_ENTERED to keep formulas.
  - Post-write formatting wrapped in try/except to avoid hard failures.
- Date sorting & de-duplication:
  - Merge order files by unique Order ID; sort by order date descending.

## 6. Do's and Don'ts
- Do
  - Keep scripts small and single-purpose; place shared logic in scripts/ modules.
  - Preserve inventory key semantics: parts require color_id; sets/minifigs must use (item_id, None).
  - Maintain proportional fee distribution logic when adding new order fields.
  - Use get_color_name to normalize color labels for UI/Sheets.
  - Keep Google Sheets formulas "USER_ENTERED" and avoid overwriting existing user-provided prices in Summary.
  - Round currency values to 2 decimals for display; maintain precise floats internally when possible.
  - Add unit tests for new logic, especially around limiting build counts and cost computation.
- Don't
  - Don’t mutate the source inventory passed into build logic; always work on a copy.
  - Don’t hardcode absolute paths; use os.path and project constants (e.g., ORDERS_DIR, WANTED_LISTS_DIR).
  - Don’t assume date formats; continue to parse with fallbacks.
  - Don’t swallow broad exceptions around core algorithms; only guard I/O or UI formatting.
  - Don’t change Orders/Inventory/Summary column orders without updating dependent code and tests.

## 7. Tools & Dependencies
- Key libraries
  - gspread — Google Sheets API client.
  - google-auth — Service account credentials for gspread.
  - stdlib: csv, xml.etree.ElementTree, datetime, os, copy, collections.
- Setup
  - Python 3.10+ recommended.
  - pip install gspread google-auth
  - Place Google service account JSON at scripts/credentials.json.
  - Ensure the service account has access to the target Google Drive/Sheets, or let the tool create the sheet.
- Running
  - Merge + compute + update Sheets: cd scripts && python main.py
  - Tests: python -m unittest -v

## 8. Other Notes
- Orders sheet columns are explicitly defined and formatted; maintain headers when adding new fields.
- Summary formulas compute Profit, Margin, Markup, and suggested prices (75/100/125/150%); ensure Config!B1 (Shipping Fee) and Config!B2 (Materials Cost) exist.
- wanted_lists items may mark parts as isMinifigPart=True to trigger parts-only build logic.
- colors.get_color_name returns None for unknown numeric IDs, and returns the string itself for non-numeric inputs; respect this in displays and tests.
- Tests currently manipulate sys.path to import from scripts; if you refactor to a package, update imports to relative (from .config import ...) and tests accordingly.
