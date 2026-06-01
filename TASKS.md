# Tasks

## Status

- open
- in progress
- done
- blocked

## Task List

### Task 001

- Status: done
- Title: Excel to WooCommerce export pipeline
- Goal: Convert a local Excel product sheet into WooCommerce-ready import files for digital products.
- What to do: Add a Python converter, document the input/output format, and verify the result on a sample workbook.
- Files: `README.md`, `START_HERE.md`, `scripts/wp_excel_to_woocommerce_csv.py`, `tests/test_wp_excel_to_woocommerce_csv.py`
- How to verify: Run `python -m unittest discover -s .\tests` and a sample conversion command that creates `output/products.csv`.

### Task 002

- Status: done
- Title: Desktop app and product folder assets
- Goal: Let the user run the workflow from a window and use a product folder containing listing images plus a ZIP download.
- What to do: Add a Tkinter app, a double-click launcher, image/ZIP discovery from `product_folder`, CSV `Images` output, and manifest asset types.
- Files: `README.md`, `START_HERE.md`, `Start Product Import App.cmd`, `scripts/product_import_app.py`, `scripts/wp_excel_to_woocommerce_csv.py`, `tests/test_wp_excel_to_woocommerce_csv.py`
- How to verify: Run `python -m unittest discover -s .\tests` and `python -m py_compile .\scripts\product_import_app.py .\scripts\wp_excel_to_woocommerce_csv.py`.

### Task 003

- Status: done
- Title: Source workbook status and draft/publish mode
- Goal: Support the provided Excel structure, choose draft or ready listing mode, and write processing status/date back to the workbook.
- What to do: Map `Photos`, `Digital File (ZIP)`, and `Путь к файлам`; add `draft/publish` selection in the app; update source workbook columns `status`, `status_date`, and `publish_mode`.
- Files: `README.md`, `TASKS.md`, `scripts/product_import_app.py`, `scripts/wp_excel_to_woocommerce_csv.py`, `tests/test_wp_excel_to_woocommerce_csv.py`
- How to verify: Run `python -m unittest discover -s .\tests`.

### Task 004

- Status: done
- Title: Site settings tab
- Goal: Let the user enter and persist WordPress/WooCommerce access settings from the app.
- What to do: Add a separate settings tab, fields for site/API/media settings, local JSON persistence, gitignore for the saved secrets file, and a connection test button.
- Files: `README.md`, `TASKS.md`, `CHANGELOG.md`, `.gitignore`, `scripts/product_import_app.py`, `tests/test_product_import_app_settings.py`
- How to verify: Run `python -m unittest discover -s .\tests`.

### Task 005

- Status: done
- Title: Double-click launcher without console
- Goal: Let the user open the desktop app without using PowerShell or a visible command prompt.
- What to do: Add a `.vbs` launcher, keep the `.cmd` fallback for diagnostics, and document the preferred launcher.
- Files: `Open Product Import App.vbs`, `Start Product Import App.cmd`, `README.md`, `START_HERE.md`, `TASKS.md`, `CHANGELOG.md`
- How to verify: Run `python -m unittest discover -s .\tests` and import the Tkinter app module.

### Task 006

- Status: done
- Title: Locked workbook status fallback
- Goal: Avoid failing the whole export when the source Excel workbook cannot be overwritten.
- What to do: Save status updates to an output copy when the source workbook is locked or permission denied, and report that path in the app log.
- Files: `README.md`, `TASKS.md`, `CHANGELOG.md`, `scripts/product_import_app.py`, `scripts/wp_excel_to_woocommerce_csv.py`, `tests/test_wp_excel_to_woocommerce_csv.py`
- How to verify: Run `python -m unittest discover -s .\tests`.

### Task 007

- Status: open
- Title: Upload files to WordPress automatically
- Goal: Upload listing images and ZIP files to WordPress, then create/update products through WooCommerce API.
- What to do: Add site settings, secure credential handling, Media Library upload, WooCommerce product creation, status `Размещено`, product URL, and a dry-run mode.
- Files: TBD
- How to verify: Publish one test product to a staging or draft WooCommerce product and write status/date/product URL back to Excel.
