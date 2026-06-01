from __future__ import annotations

import json
import threading
import tkinter as tk
from base64 import b64encode
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin
from urllib.request import Request, urlopen

from wp_excel_to_woocommerce_csv import convert_workbook


ROOT_DIR = Path(__file__).resolve().parents[1]
SETTINGS_PATH = ROOT_DIR / "config" / "site_settings.json"
SITE_SETTING_KEYS = [
    "site_url",
    "woocommerce_key",
    "woocommerce_secret",
    "wp_username",
    "wp_application_password",
    "media_base_url",
]


def load_site_settings(path: Path = SETTINGS_PATH) -> dict[str, str]:
    if not path.exists():
        return {key: "" for key in SITE_SETTING_KEYS}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {key: "" for key in SITE_SETTING_KEYS}
    return {key: str(raw.get(key, "")) for key in SITE_SETTING_KEYS}


def save_site_settings(settings: dict[str, str], path: Path = SETTINGS_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {key: str(settings.get(key, "")).strip() for key in SITE_SETTING_KEYS}
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def test_wordpress_connection(settings: dict[str, str], timeout: int = 15) -> tuple[bool, str]:
    site_url = settings.get("site_url", "").strip().rstrip("/")
    consumer_key = settings.get("woocommerce_key", "").strip()
    consumer_secret = settings.get("woocommerce_secret", "").strip()
    if not site_url:
        return False, "Укажите URL сайта."
    if not consumer_key or not consumer_secret:
        return False, "Укажите WooCommerce Consumer Key и Consumer Secret."

    try:
        wp_url = urljoin(site_url + "/", "wp-json/")
        with urlopen(Request(wp_url, headers={"User-Agent": "AnastasiaImporter/1.0"}), timeout=timeout) as response:
            if response.status >= 400:
                return False, f"WordPress REST API вернул статус {response.status}."

        wc_url = urljoin(site_url + "/", "wp-json/wc/v3/system_status")
        token = b64encode(f"{consumer_key}:{consumer_secret}".encode("utf-8")).decode("ascii")
        request = Request(
            wc_url,
            headers={
                "Authorization": f"Basic {token}",
                "User-Agent": "AnastasiaImporter/1.0",
            },
        )
        with urlopen(request, timeout=timeout) as response:
            if response.status >= 400:
                return False, f"WooCommerce API вернул статус {response.status}."
        return True, "Подключение к WordPress и WooCommerce работает."
    except HTTPError as error:
        if error.code in {401, 403}:
            return False, "WooCommerce отклонил доступ. Проверьте Consumer Key/Secret и права ключа."
        return False, f"HTTP ошибка при проверке подключения: {error.code}."
    except URLError as error:
        return False, f"Не удалось подключиться к сайту: {error.reason}."
    except Exception as error:
        return False, f"Ошибка проверки подключения: {error}"


class ProductImportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Anastasia - WordPress product import")
        self.geometry("860x600")
        self.minsize(760, 520)

        saved_settings = load_site_settings()
        self.input_path = tk.StringVar(value=str(ROOT_DIR / "input" / "products.xlsx"))
        self.output_dir = tk.StringVar(value=str(ROOT_DIR / "output"))
        self.sheet_name = tk.StringVar(value="")
        self.media_base_url = tk.StringVar(value=saved_settings["media_base_url"])
        self.publication_status = tk.StringVar(value="draft")
        self.update_source_status = tk.BooleanVar(value=True)

        self.site_url = tk.StringVar(value=saved_settings["site_url"])
        self.woocommerce_key = tk.StringVar(value=saved_settings["woocommerce_key"])
        self.woocommerce_secret = tk.StringVar(value=saved_settings["woocommerce_secret"])
        self.wp_username = tk.StringVar(value=saved_settings["wp_username"])
        self.wp_application_password = tk.StringVar(value=saved_settings["wp_application_password"])

        self._build_ui()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        notebook = ttk.Notebook(self)
        notebook.grid(row=0, column=0, sticky="nsew")

        import_tab = ttk.Frame(notebook, padding=18)
        settings_tab = ttk.Frame(notebook, padding=18)
        notebook.add(import_tab, text="Импорт товаров")
        notebook.add(settings_tab, text="Настройки сайта")

        self._build_import_tab(import_tab)
        self._build_settings_tab(settings_tab)

    def _build_import_tab(self, container: ttk.Frame) -> None:
        container.columnconfigure(1, weight=1)
        container.rowconfigure(9, weight=1)

        title = ttk.Label(container, text="Подготовка товаров для WooCommerce", font=("Segoe UI", 15, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 14))

        self._path_row(container, 1, "Excel-файл", self.input_path, self._pick_excel)
        self._path_row(container, 2, "Папка результата", self.output_dir, self._pick_output_dir)

        ttk.Label(container, text="Лист Excel").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=self.sheet_name).grid(row=3, column=1, columnspan=2, sticky="ew", pady=6)

        ttk.Label(container, text="URL папки медиа").grid(row=4, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=self.media_base_url).grid(row=4, column=1, columnspan=2, sticky="ew", pady=6)

        ttk.Label(container, text="Режим листинга").grid(row=5, column=0, sticky="w", pady=6)
        mode_frame = ttk.Frame(container)
        mode_frame.grid(row=5, column=1, columnspan=2, sticky="w", pady=6)
        ttk.Radiobutton(mode_frame, text="Черновик", variable=self.publication_status, value="draft").grid(
            row=0, column=0, sticky="w", padx=(0, 18)
        )
        ttk.Radiobutton(mode_frame, text="Готовый листинг", variable=self.publication_status, value="publish").grid(
            row=0, column=1, sticky="w"
        )

        ttk.Checkbutton(
            container,
            text="Записать статус и дату обратно в Excel",
            variable=self.update_source_status,
        ).grid(row=6, column=1, columnspan=2, sticky="w", pady=6)

        hint = ttk.Label(
            container,
            text=(
                "Если URL пустой, manifest покажет локальные файлы, но CSV не сможет сразу импортировать "
                "картинки и ZIP как публичные ссылки."
            ),
            wraplength=700,
        )
        hint.grid(row=7, column=0, columnspan=3, sticky="w", pady=(4, 12))

        self.run_button = ttk.Button(container, text="Собрать CSV", command=self._run_conversion)
        self.run_button.grid(row=8, column=0, sticky="w", pady=(0, 12))

        self.log = tk.Text(container, height=12, wrap="word")
        self.log.grid(row=9, column=0, columnspan=3, sticky="nsew")

    def _build_settings_tab(self, container: ttk.Frame) -> None:
        container.columnconfigure(1, weight=1)

        title = ttk.Label(container, text="Доступ к WordPress/WooCommerce", font=("Segoe UI", 15, "bold"))
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 14))

        self._settings_row(container, 1, "URL сайта", self.site_url)
        self._settings_row(container, 2, "WooCommerce Consumer Key", self.woocommerce_key, secret=True)
        self._settings_row(container, 3, "WooCommerce Consumer Secret", self.woocommerce_secret, secret=True)
        self._settings_row(container, 4, "WordPress username", self.wp_username)
        self._settings_row(container, 5, "WordPress Application Password", self.wp_application_password, secret=True)
        self._settings_row(container, 6, "URL папки медиа", self.media_base_url)

        hint = ttk.Label(
            container,
            text=(
                "Настройки сохраняются локально на этом компьютере. Поля нужны для следующего шага: "
                "загрузки медиа и создания товаров через API."
            ),
            wraplength=720,
        )
        hint.grid(row=7, column=0, columnspan=2, sticky="w", pady=(10, 14))

        buttons = ttk.Frame(container)
        buttons.grid(row=8, column=0, columnspan=2, sticky="w")
        ttk.Button(buttons, text="Сохранить настройки", command=self._save_settings).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(buttons, text="Загрузить сохранённые", command=self._reload_settings).grid(row=0, column=1, padx=(0, 10))
        self.test_connection_button = ttk.Button(
            buttons,
            text="Тест подключения",
            command=self._run_connection_test,
        )
        self.test_connection_button.grid(row=0, column=2)

        self.settings_status = ttk.Label(container, text="")
        self.settings_status.grid(row=9, column=0, columnspan=2, sticky="w", pady=(12, 0))

    def _settings_row(
        self,
        container: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        secret: bool = False,
    ) -> None:
        ttk.Label(container, text=label).grid(row=row, column=0, sticky="w", pady=6)
        show = "*" if secret else ""
        ttk.Entry(container, textvariable=variable, show=show).grid(row=row, column=1, sticky="ew", pady=6)

    def _path_row(
        self,
        container: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command: Callable[[], None],
    ) -> None:
        ttk.Label(container, text=label).grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=6, padx=(0, 8))
        ttk.Button(container, text="Выбрать", command=command).grid(row=row, column=2, sticky="ew", pady=6)

    def _pick_excel(self) -> None:
        selected = filedialog.askopenfilename(
            title="Выбрать Excel-файл",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if selected:
            self.input_path.set(selected)

    def _pick_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="Выбрать папку результата")
        if selected:
            self.output_dir.set(selected)

    def _collect_settings(self) -> dict[str, str]:
        return {
            "site_url": self.site_url.get(),
            "woocommerce_key": self.woocommerce_key.get(),
            "woocommerce_secret": self.woocommerce_secret.get(),
            "wp_username": self.wp_username.get(),
            "wp_application_password": self.wp_application_password.get(),
            "media_base_url": self.media_base_url.get(),
        }

    def _apply_settings(self, settings: dict[str, str]) -> None:
        self.site_url.set(settings["site_url"])
        self.woocommerce_key.set(settings["woocommerce_key"])
        self.woocommerce_secret.set(settings["woocommerce_secret"])
        self.wp_username.set(settings["wp_username"])
        self.wp_application_password.set(settings["wp_application_password"])
        self.media_base_url.set(settings["media_base_url"])

    def _save_settings(self) -> None:
        save_site_settings(self._collect_settings())
        self.settings_status.configure(text=f"Сохранено: {SETTINGS_PATH}")
        messagebox.showinfo("Сохранено", "Настройки сайта сохранены локально.")

    def _reload_settings(self) -> None:
        self._apply_settings(load_site_settings())
        self.settings_status.configure(text="Сохранённые настройки загружены.")

    def _run_connection_test(self) -> None:
        self.test_connection_button.configure(state="disabled")
        self.settings_status.configure(text="Проверяю подключение...")
        threading.Thread(target=self._test_connection_in_background, daemon=True).start()

    def _show_message(self, ok: bool, title: str, message: str) -> None:
        if ok:
            self.after(0, messagebox.showinfo, title, message)
        else:
            self.after(0, messagebox.showerror, title, message)

    def _finish_connection_test(self, ok: bool, message: str) -> None:
        self.settings_status.configure(text=message)
        self._show_message(ok, "Подключение работает" if ok else "Проверка не прошла", message)
        self.test_connection_button.configure(state="normal")

    def _test_connection_in_background(self) -> None:
        ok, message = test_wordpress_connection(self._collect_settings())
        self.after(0, self._finish_connection_test, ok, message)

    def _append_log(self, text: str) -> None:
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def _run_conversion(self) -> None:
        self.run_button.configure(state="disabled")
        self.log.delete("1.0", "end")
        thread = threading.Thread(target=self._convert_in_background, daemon=True)
        thread.start()

    def _finish_conversion_success(
        self,
        products: int,
        publication_status: str,
        output_path: Path,
        manifest_path: Path,
        update_source_status: bool,
        status_workbook_path: Path | None,
        status_saved_to_source: bool,
        missing: int,
    ) -> None:
        self._append_log(f"Готово. Товаров обработано: {products}")
        self._append_log(f"Режим листинга: {publication_status}")
        self._append_log(f"CSV: {output_path}")
        self._append_log(f"Manifest: {manifest_path}")
        if update_source_status:
            if status_saved_to_source:
                self._append_log("Excel обновлён: status, status_date, publish_mode")
            elif status_workbook_path:
                self._append_log(f"Исходный Excel был недоступен для записи. Обновлённая копия: {status_workbook_path}")
        if missing:
            self._append_log(f"Проблем с файлами: {missing}. Смотри manifest.")
        self._show_message(True, "Готово", "CSV и manifest созданы.")
        self.run_button.configure(state="normal")

    def _finish_conversion_error(self, error: Exception) -> None:
        self._append_log(f"Ошибка: {error}")
        self._show_message(False, "Ошибка", str(error))
        self.run_button.configure(state="normal")

    def _convert_in_background(self) -> None:
        try:
            input_path = Path(self.input_path.get()).expanduser()
            output_dir = Path(self.output_dir.get()).expanduser()
            output_path = output_dir / "products.csv"
            manifest_path = output_dir / "downloads_manifest.csv"
            sheet_name = self.sheet_name.get().strip() or None
            media_base_url = self.media_base_url.get().strip() or None
            publication_status = self.publication_status.get()
            update_source_status = self.update_source_status.get()

            if not input_path.exists():
                raise FileNotFoundError(f"Excel-файл не найден: {input_path}")

            result = convert_workbook(
                input_path=input_path,
                output_path=output_path,
                manifest_path=manifest_path,
                sheet_name=sheet_name,
                download_base_url=media_base_url,
                publication_status=publication_status,
                update_source_status=update_source_status,
            )

            self.after(
                0,
                self._finish_conversion_success,
                result.product_count,
                publication_status,
                output_path,
                manifest_path,
                update_source_status,
                result.status_workbook_path,
                result.status_saved_to_source,
                result.missing_downloads,
            )
        except Exception as error:
            self.after(0, self._finish_conversion_error, error)


if __name__ == "__main__":
    ProductImportApp().mainloop()
