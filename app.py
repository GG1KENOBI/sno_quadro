#!/usr/bin/env python3
"""
SNO Quadro — Веб-приложение для поиска фото товаров из Excel.
Flask + Selenium + Яндекс.Картинки + rembg.
"""

import os
import re
import json
import time
import uuid
import logging
import threading
import zipfile
from pathlib import Path
from io import BytesIO
from urllib.parse import quote_plus, urlparse, unquote

from flask import (
    Flask, render_template, request, jsonify, send_file,
    send_from_directory,
)
from PIL import Image
import requests as http_requests

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

from search_images import extract_products, Product

# ─── Настройка ────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

UPLOAD_DIR = Path("uploads")
TEMP_DIR = Path("temp_images")
OUTPUT_DIR = Path("output")

for d in [UPLOAD_DIR, TEMP_DIR, OUTPUT_DIR]:
    d.mkdir(exist_ok=True)

sessions: dict[str, dict] = {}

# ─── Selenium (singleton) ────────────────────────────────────────────────────

_driver = None
_driver_lock = threading.Lock()


def get_driver():
    global _driver
    with _driver_lock:
        if _driver is not None:
            return _driver
        log.info("Запуск Chrome (headless)...")
        opts = Options()
        opts.add_argument("--headless=new")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument(
            "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        opts.add_argument("--log-level=3")
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        _driver = webdriver.Chrome(options=opts)
        log.info("Chrome запущен.")
        return _driver


def close_driver():
    global _driver
    if _driver is not None:
        try:
            _driver.quit()
        except Exception:
            pass
        _driver = None


# ─── Поиск через Яндекс.Картинки ─────────────────────────────────────────────

def search_yandex_images(query: str, max_results: int = 5) -> list[dict]:
    """
    Ищет изображения через Яндекс.Картинки.
    Возвращает список словарей: [{url, source_url, title}]
    """
    driver = get_driver()
    results = []

    try:
        encoded = quote_plus(query)
        search_url = f"https://yandex.ru/images/search?text={encoded}&isize=large"
        driver.get(search_url)
        time.sleep(3)

        # Собираем превью
        img_elements = driver.find_elements(
            By.CSS_SELECTOR,
            "img.ImagesContentImage-Image, img[class*='ContentImage-Image'], .SerpItem-Thumb img"
        )
        if not img_elements:
            img_elements = [
                el for el in driver.find_elements(By.TAG_NAME, "img")
                if (el.get_attribute("src") or "").startswith("https://avatars.mds.yandex.net")
            ]

        log.info(f"  Превью на странице: {len(img_elements)}")

        for i, img_el in enumerate(img_elements[:max_results + 5]):
            if len(results) >= max_results:
                break

            try:
                driver.execute_script("arguments[0].click();", img_el)
                time.sleep(1.5)

                img_url = None
                source_url = None
                title = ""

                # 1) Кнопка «Открыть» → оригинальный URL картинки
                for sel in [
                    "a.MMViewerButtons-OpenImage",
                    "a[class*='OpenImage']",
                ]:
                    try:
                        links = driver.find_elements(By.CSS_SELECTOR, sel)
                        for link in links:
                            href = link.get_attribute("href") or ""
                            if href.startswith("http") and "yandex" not in href:
                                img_url = href
                                break
                    except Exception:
                        continue
                    if img_url:
                        break

                # 2) Ссылка на сайт-источник
                for sel in [
                    "a.MMViewerButtons-VisitButton",
                    "a[class*='VisitButton']",
                    ".MMSitePanel-Link a",
                    "a[class*='SiteLink']",
                    ".MMViewerButtons a[target='_blank']",
                ]:
                    try:
                        links = driver.find_elements(By.CSS_SELECTOR, sel)
                        for link in links:
                            href = link.get_attribute("href") or ""
                            if href.startswith("http") and "yandex" not in href:
                                source_url = href
                                break
                    except Exception:
                        continue
                    if source_url:
                        break

                # 3) Если нет source_url — берём из img_url домен
                if not source_url and img_url:
                    from urllib.parse import urlparse
                    parsed = urlparse(img_url)
                    source_url = f"{parsed.scheme}://{parsed.netloc}"

                # 4) Название из alt или title
                title = img_el.get_attribute("alt") or ""

                # 5) Fallback для img_url — img в просмотрщике
                if not img_url:
                    for sel in [".MMImage-Origin", "img.MMImage-Origin", ".MMImageContainer img"]:
                        try:
                            viewer_imgs = driver.find_elements(By.CSS_SELECTOR, sel)
                            for vi in viewer_imgs:
                                src = vi.get_attribute("src") or ""
                                if src.startswith("http"):
                                    img_url = src
                                    break
                        except Exception:
                            continue
                        if img_url:
                            break

                # 6) Fallback — превью URL
                if not img_url:
                    src = img_el.get_attribute("src") or ""
                    if src.startswith("http"):
                        img_url = re.sub(r'&n=\d+', '', src)

                if img_url and img_url not in [r["url"] for r in results]:
                    results.append({
                        "url": img_url,
                        "source_url": source_url or "",
                        "title": title[:100],
                    })

                # Закрываем просмотрщик
                closed = False
                for sel in [".MMViewerModal-Close", ".Modal-Close", "[class*='CloseButton']"]:
                    try:
                        btns = driver.find_elements(By.CSS_SELECTOR, sel)
                        for btn in btns:
                            btn.click()
                            closed = True
                            time.sleep(0.3)
                            break
                    except Exception:
                        continue
                    if closed:
                        break
                if not closed:
                    driver.back()
                    time.sleep(0.5)

            except Exception as e:
                log.debug(f"  Ошибка картинки {i}: {e}")
                src = img_el.get_attribute("src") or ""
                if src.startswith("http") and src not in [r["url"] for r in results]:
                    results.append({
                        "url": re.sub(r'&n=\d+', '', src),
                        "source_url": "",
                        "title": img_el.get_attribute("alt") or "",
                    })

    except Exception as e:
        log.error(f"Ошибка поиска: {e}")

    log.info(f"  Найдено URL: {len(results)}")
    return results[:max_results]


# ─── Скачивание / обработка ───────────────────────────────────────────────────

def download_image(url: str, timeout: int = 15) -> bytes | None:
    try:
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        }
        resp = http_requests.get(url, headers=headers, timeout=timeout)
        resp.raise_for_status()
        ct = resp.headers.get("Content-Type", "")
        if "image" not in ct and "octet" not in ct:
            return None
        data = resp.content
        img = Image.open(BytesIO(data))
        img.verify()
        img = Image.open(BytesIO(data))
        w, h = img.size
        if w < 100 or h < 100:
            return None
        return data
    except Exception:
        return None


def remove_background(image_data: bytes) -> bytes:
    from rembg import remove
    result = remove(image_data)
    img = Image.open(BytesIO(result)).convert("RGBA")
    white_bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
    white_bg.paste(img, (0, 0), img)
    output = white_bg.convert("RGB")
    buf = BytesIO()
    output.save(buf, format="JPEG", quality=95)
    return buf.getvalue()


def resize_image(image_data: bytes, width: int, height: int) -> bytes:
    img = Image.open(BytesIO(image_data)).convert("RGB")
    img.thumbnail((width, height), Image.LANCZOS)
    canvas = Image.new("RGB", (width, height), (255, 255, 255))
    x = (width - img.width) // 2
    y = (height - img.height) // 2
    canvas.paste(img, (x, y))
    buf = BytesIO()
    canvas.save(buf, format="JPEG", quality=95)
    return buf.getvalue()


def extract_name_from_url(url: str, brand: str = "", model: str = "") -> str:
    """Извлекает осмысленное имя товара из URL источника."""
    try:
        parsed = urlparse(url)
        path = unquote(parsed.path).lower()
        # Берём последний сегмент пути (без расширения)
        segments = [s for s in path.split("/") if s and not s.startswith(".")]
        if not segments:
            return ""
        last = segments[-1]
        # Убираем расширение файла
        last = re.sub(r'\.(jpe?g|png|webp|gif|bmp|avif|tiff?)$', '', last, flags=re.I)
        # Убираем чисто числовые и хеш-строки
        if re.match(r'^[0-9a-f\-_]{20,}$', last, re.I):
            # Попробуем предпоследний сегмент
            if len(segments) >= 2:
                last = segments[-2]
                last = re.sub(r'\.(jpe?g|png|webp|gif|bmp|avif|tiff?)$', '', last, flags=re.I)
            else:
                return ""
        # Заменяем разделители на пробелы для читаемости
        name = re.sub(r'[-_+]+', ' ', last).strip()
        name = re.sub(r'%[0-9a-fA-F]{2}', ' ', name).strip()
        name = re.sub(r'\s+', ' ', name)
        # Если в имени упоминается бренд/модель — отлично, оставляем
        if name and len(name) > 5:
            return name[:80]
    except Exception:
        pass
    return ""


# ─── Flask routes ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "Файл не выбран"}), 400
    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Файл не выбран"}), 400
    ext = Path(file.filename).suffix.lower()
    if ext not in (".xls", ".xlsx", ".xlsm"):
        return jsonify({"error": "Поддерживаются .xls / .xlsx"}), 400

    session_id = str(uuid.uuid4())[:8]
    filepath = UPLOAD_DIR / f"{session_id}{ext}"
    file.save(filepath)

    try:
        products = extract_products(str(filepath))
    except Exception as e:
        return jsonify({"error": f"Ошибка парсинга: {e}"}), 400
    if not products:
        return jsonify({"error": "Товары не найдены"}), 400

    sessions[session_id] = {
        "products": [
            {"idx": i, "brand": p.brand, "model": p.model,
             "color": p.color, "raw_name": p.raw_name, "display": str(p)}
            for i, p in enumerate(products)
        ],
        "images": {},
    }
    return jsonify({
        "session_id": session_id,
        "products": sessions[session_id]["products"],
        "count": len(products),
    })


@app.route("/search", methods=["POST"])
def search_images_route():
    data = request.json
    session_id = data.get("session_id")
    product_idx = data.get("product_idx")
    if session_id not in sessions:
        return jsonify({"error": "Сессия не найдена"}), 404

    session = sessions[session_id]
    product = session["products"][product_idx]

    queries = [
        f"{product['brand']} {product['model']} {product['color']} очки",
        f"{product['brand']} {product['model']} {product['color']} eyewear white background",
        f"{product['brand']} {product['model']} glasses",
    ]

    all_results = []
    seen_urls = set()
    for q in queries:
        found = search_yandex_images(q, max_results=6)
        for r in found:
            if r["url"] not in seen_urls:
                seen_urls.add(r["url"])
                all_results.append(r)
        if len(all_results) >= 5:
            break
        time.sleep(0.5)

    all_results = all_results[:5]

    images = []
    product_dir = TEMP_DIR / session_id / str(product_idx)
    product_dir.mkdir(parents=True, exist_ok=True)

    for i, res in enumerate(all_results):
        img_data = download_image(res["url"])
        if img_data is None:
            continue

        img_id = f"{product_idx}_{i}"
        img_path = product_dir / f"{img_id}.jpg"
        with open(img_path, "wb") as f:
            f.write(img_data)

        img = Image.open(BytesIO(img_data))
        w, h = img.size

        images.append({
            "id": img_id,
            "url": res["url"],
            "source_url": res["source_url"],
            "title": res["title"],
            "local_path": str(img_path),
            "width": w,
            "height": h,
            "selected": False,
            "bg_removed": False,
        })

    session["images"][str(product_idx)] = images

    return jsonify({
        "product_idx": product_idx,
        "images": [
            {
                "id": img["id"],
                "width": img["width"],
                "height": img["height"],
                "source_url": img["source_url"],
                "title": img["title"],
                "preview_url": f"/preview/{session_id}/{img['id']}",
            }
            for img in images
        ],
    })


@app.route("/preview/<session_id>/<img_id>")
def preview_image(session_id, img_id):
    session = sessions.get(session_id)
    if not session:
        return "Not found", 404
    product_idx = img_id.split("_")[0]
    imgs = session["images"].get(product_idx, [])
    img_info = next((i for i in imgs if i["id"] == img_id), None)
    if not img_info:
        return "Not found", 404
    return send_file(img_info["local_path"], mimetype="image/jpeg")


@app.route("/remove_bg", methods=["POST"])
def remove_bg():
    data = request.json
    session_id = data.get("session_id")
    img_id = data.get("img_id")

    session = sessions.get(session_id)
    if not session:
        return jsonify({"error": "Сессия не найдена"}), 404
    product_idx = img_id.split("_")[0]
    imgs = session["images"].get(product_idx, [])
    img_info = next((i for i in imgs if i["id"] == img_id), None)
    if not img_info:
        return jsonify({"error": "Изображение не найдено"}), 404

    try:
        with open(img_info["local_path"], "rb") as f:
            original = f.read()
        result = remove_background(original)
        new_path = img_info["local_path"].replace(".jpg", "_nobg.jpg")
        with open(new_path, "wb") as f:
            f.write(result)
        img_info["local_path"] = new_path
        img_info["bg_removed"] = True
        img = Image.open(BytesIO(result))
        img_info["width"], img_info["height"] = img.size
        return jsonify({
            "success": True, "img_id": img_id,
            "preview_url": f"/preview/{session_id}/{img_id}?t={int(time.time())}",
            "width": img_info["width"], "height": img_info["height"],
        })
    except Exception as e:
        log.error(f"Ошибка удаления фона: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/save", methods=["POST"])
def save_selected():
    data = request.json
    session_id = data.get("session_id")
    selected = data.get("selected", {})
    width = data.get("width", 0)
    height = data.get("height", 0)

    session = sessions.get(session_id)
    if not session:
        return jsonify({"error": "Сессия не найдена"}), 404

    # Собираем файлы в ZIP
    zip_buf = BytesIO()
    saved = []
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for product_idx_str, img_ids in selected.items():
            product = session["products"][int(product_idx_str)]
            imgs = session["images"].get(product_idx_str, [])
            for img_id in img_ids:
                img_info = next((i for i in imgs if i["id"] == img_id), None)
                if not img_info:
                    continue
                with open(img_info["local_path"], "rb") as f:
                    img_data = f.read()
                if width > 0 and height > 0:
                    img_data = resize_image(img_data, width, height)

                # Пробуем извлечь имя из URL источника
                url_name = extract_name_from_url(
                    img_info.get("source_url") or img_info.get("url", ""),
                    product["brand"], product["model"],
                )
                if url_name:
                    safe_name = re.sub(r'[^\w\-\. ]', '_', url_name)[:60]
                else:
                    safe_name = f"{product['brand']}_{product['model']}_{product['color']}"
                    safe_name = re.sub(r'[^\w\-]', '_', safe_name)

                suffix = f"_{img_id.split('_')[1]}" if len(img_ids) > 1 else ""
                filename = f"{safe_name}{suffix}.jpg"
                zf.writestr(filename, img_data)
                saved.append(filename)

    if not saved:
        return jsonify({"error": "Нечего сохранять"}), 400

    zip_buf.seek(0)
    return send_file(
        zip_buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"sno_quadro_{session_id}.zip",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
