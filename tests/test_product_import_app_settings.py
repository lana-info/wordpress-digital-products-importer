from __future__ import annotations

import unittest
from pathlib import Path
from sys import path as sys_path

sys_path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from product_import_app import load_site_settings, save_site_settings, test_wordpress_connection


class SiteSettingsTest(unittest.TestCase):
    def test_saves_and_loads_site_settings(self) -> None:
        settings_path = Path(__file__).resolve().parents[1] / "output" / "__test_site_settings.json"
        try:
            if settings_path.exists():
                settings_path.unlink()
        except PermissionError:
            pass

        try:
            save_site_settings(
                {
                    "site_url": "https://example.com",
                    "woocommerce_key": "ck_test",
                    "woocommerce_secret": "cs_test",
                    "wp_username": "admin",
                    "wp_application_password": "app pass",
                    "media_base_url": "https://example.com/wp-content/uploads/products",
                },
                path=settings_path,
            )

            loaded = load_site_settings(settings_path)

            self.assertEqual(loaded["site_url"], "https://example.com")
            self.assertEqual(loaded["woocommerce_key"], "ck_test")
            self.assertEqual(loaded["woocommerce_secret"], "cs_test")
            self.assertEqual(loaded["wp_username"], "admin")
            self.assertEqual(loaded["wp_application_password"], "app pass")
            self.assertEqual(loaded["media_base_url"], "https://example.com/wp-content/uploads/products")
        finally:
            try:
                if settings_path.exists():
                    settings_path.unlink()
            except PermissionError:
                pass

    def test_connection_requires_site_and_woocommerce_keys(self) -> None:
        ok, message = test_wordpress_connection({"site_url": ""}, timeout=1)
        self.assertFalse(ok)
        self.assertIn("URL", message)

        ok, message = test_wordpress_connection({"site_url": "https://example.com"}, timeout=1)
        self.assertFalse(ok)
        self.assertIn("WooCommerce", message)


if __name__ == "__main__":
    unittest.main()
