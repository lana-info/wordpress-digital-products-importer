# Changelog

## Format

- Date:
- What changed:
- Files touched:
- How verified:
- Risks:

## 2026-05-29

- Date: 2026-05-29
- What changed: Added a working Excel-to-WooCommerce CSV converter for digital products, plus docs, task tracking, and a unit test.
- Files touched: `README.md`, `START_HERE.md`, `TASKS.md`, `CHANGELOG.md`, `scripts/README.md`, `scripts/wp_excel_to_woocommerce_csv.py`, `tests/test_wp_excel_to_woocommerce_csv.py`
- How verified: `python -m unittest discover -s .\tests` passed after the change.
- Risks: WooCommerce download URLs still need a public URL or manual upload step if the source file only exists locally.

## 2026-05-29 - App and product folders

- Date: 2026-05-29
- What changed: Added a desktop Tkinter app, a double-click launcher, product-folder discovery for listing images and ZIP downloads, and CSV `Images` output.
- Files touched: `README.md`, `START_HERE.md`, `TASKS.md`, `CHANGELOG.md`, `Start Product Import App.cmd`, `scripts/README.md`, `scripts/product_import_app.py`, `scripts/wp_excel_to_woocommerce_csv.py`, `tests/test_wp_excel_to_woocommerce_csv.py`
- How verified: `python -m unittest discover -s .\tests` passed; read-only `ast.parse` syntax check passed; `product_import_app` imported successfully.
- Risks: The CSV can reference files only after they have public URLs on WordPress or hosting.

## 2026-05-29 - Workbook status and publish mode

- Date: 2026-05-29
- What changed: Added support for the provided Excel columns, draft/publish selection, source workbook status/date updates, and explicit ZIP priority over folder auto-detection.
- Files touched: `README.md`, `TASKS.md`, `CHANGELOG.md`, `scripts/product_import_app.py`, `scripts/wp_excel_to_woocommerce_csv.py`, `tests/test_wp_excel_to_woocommerce_csv.py`
- How verified: `python -m unittest discover -s .\tests` passed; read-only `ast.parse` syntax check passed.
- Risks: Actual placed status should be written only after WordPress API publishing is added and confirms success.

## 2026-05-29 - Site settings tab

- Date: 2026-05-29
- What changed: Added a separate app tab for WordPress/WooCommerce access settings, local JSON persistence for site credentials, and a connection test button.
- Files touched: `README.md`, `TASKS.md`, `CHANGELOG.md`, `.gitignore`, `scripts/product_import_app.py`, `tests/test_product_import_app_settings.py`
- How verified: `python -m unittest discover -s .\tests` passed; read-only `ast.parse` syntax check passed.
- Risks: Settings are stored locally in `config/site_settings.json`; keep that file private.

## 2026-05-31 - Double-click app launcher

- Date: 2026-05-31
- What changed: Added `Open Product Import App.vbs` for launching the desktop app without a visible command prompt, kept the `.cmd` launcher as a diagnostic fallback, and documented the preferred launcher.
- Files touched: `Open Product Import App.vbs`, `Start Product Import App.cmd`, `README.md`, `START_HERE.md`, `TASKS.md`, `CHANGELOG.md`
- How verified: `python -m unittest discover -s .\tests` passed; read-only `ast.parse` syntax check passed; `product_import_app` imported successfully.
- Risks: The launcher still depends on Python being installed and available as `pythonw.exe` or `python.exe`; a later `.exe` build can remove that dependency.

## 2026-06-01 - Locked Excel fallback

- Date: 2026-06-01
- What changed: When the source workbook cannot be overwritten during status updates, the app now saves an updated copy to the output folder instead of failing the export.
- Files touched: `README.md`, `TASKS.md`, `CHANGELOG.md`, `scripts/product_import_app.py`, `scripts/wp_excel_to_woocommerce_csv.py`, `tests/test_wp_excel_to_woocommerce_csv.py`
- How verified: `python -m unittest discover -s .\tests` passed; read-only `ast.parse` syntax check passed.
- Risks: If the source workbook cannot even be read, the user still needs to close Excel or fix file permissions.
