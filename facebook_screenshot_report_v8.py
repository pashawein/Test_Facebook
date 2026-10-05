"""
Facebook Report - Activity Log Screenshot Version
v6.0 - FINAL (Image Hash Matching + Auto Detection + No manual input)

КАК РАБОТАЕТ:
1. Читает Excel от auto_poster (знает: дату, количество постов, хеш картинки)
2. Открывает Activity Log
3. Для КАЖДОГО поста делает скриншот и сравнивает с хешем из Excel
4. Находит ТОЧНО нужные посты БЕЗ ввода текста
5. Создаёт Excel + PDF отчёт
"""

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib import colors
from PIL import Image as PILImage
import hashlib
import imagehash
import time
import os
import logging
from datetime import datetime, timedelta
import glob
import re
from pathlib import Path

# ================= НАСТРОЙКИ =================
REPORTS_FOLDER = "posting_reports"
SCROLL_PAUSE = 1
MAX_SCROLL_ATTEMPTS = 400
SCROLL_WAIT = 2  # пауза после каждого скролла (подгрузка Activity Log)
# Порог схожести хешей (0 = идентичны, 10 = похожи, 20+ = разные)
HASH_SIMILARITY_THRESHOLD = 10

# ================= ЛОГИРОВАНИЕ =================
def setup_logging(report_folder):
    log_file = os.path.join(report_folder, f"report_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
    # Сбрасываем предыдущие handlers
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return log_file


# ================= ФАЙЛОВАЯ СТРУКТУРА =================
def clean_filename(text):
    """Очистка от недопустимых Windows символов"""
    invalid_chars = [':', '/', '\\', '|', '<', '>', '?', '*', '"', ';', ',', '~', '^', '[', ']', '{', '}']
    for char in invalid_chars:
        text = text.replace(char, '-')
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(r'-+', '-', text)
    return text[:100].strip()


def find_existing_reports():
    """Поиск существующих Excel отчётов"""
    if not os.path.exists(REPORTS_FOLDER):
        os.makedirs(REPORTS_FOLDER)
    
    excel_files = glob.glob(os.path.join(REPORTS_FOLDER, "**", "*.xlsx"), recursive=True)
    
    report_list = []
    for file_path in excel_files:
        try:
            df = pd.read_excel(file_path, sheet_name='Summary')
            # Получаем данные из Summary
            info = {'path': file_path, 'name': os.path.basename(file_path)}
            for _, row in df.iterrows():
                param = str(row.get('Параметр', '')).strip()
                value = row.get('Значение', '')
                if param == 'Проект:':
                    info['project'] = str(value)
                elif param == 'Название поста:':
                    info['post_name'] = str(value)
                elif param == '✅ Успешно:':
                    info['total_posts'] = int(value) if pd.notna(value) else 0
                elif param == 'Первый шер:':
                    info['first_post'] = str(value)
                elif param == 'URL оригинала:':
                    info['original_url'] = str(value)
            report_list.append(info)
        except Exception as e:
            continue
    
    return report_list


# ================= ХЕШИ (КЛЮЧЕВАЯ ФУНКЦИЯ) =================
def calculate_perceptual_hash(image_path):
    """
    Вычисляет perceptual hash изображения (phash).
    Perceptual hash похож для визуально похожих картинок!
    
    MD5 - одинаковые байты = одинаковые файлы (слишком строго)
    pHash - визуально похожие картинки = близкие хеши (идеально!)
    """
    try:
        img = PILImage.open(image_path)
        return imagehash.phash(img)
    except Exception as e:
        logging.warning(f"⚠️ Ошибка вычисления хеша: {e}")
        return None


def load_original_hash_from_excel(excel_file):
    """
    Загружает хеш оригинального поста из Excel.
    Пробует листы 'Groups' (v5) и 'Posts' (старые версии).
    """
    try:
        # Пробуем оба варианта названий листа
        df = None
        for sheet_name in ['Groups', 'Posts']:
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                logging.info(f"✅ Читаю лист: {sheet_name}")
                break
            except:
                continue
        
        if df is None:
            logging.warning("⚠️ Листы Groups и Posts не найдены")
            return None, None
        
        # Ищем колонку Image Hash
        if 'Image Hash' not in df.columns:
            logging.warning("⚠️ Колонка Image Hash не найдена (старый отчёт без хеша)")
            return None, None
        
        # Берём первый непустой хеш
        for _, row in df.iterrows():
            hash_str = str(row.get('Image Hash', '')).strip()
            screenshot_path = str(row.get('Screenshot Path', '')).strip()
            
            if hash_str and hash_str != 'nan' and len(hash_str) > 8:
                logging.info(f"✅ Хеш найден: {hash_str[:16]}...")
                return hash_str, screenshot_path
        
        logging.warning("⚠️ Хеш не найден в Excel")
        return None, None
        
    except Exception as e:
        logging.warning(f"⚠️ Ошибка чтения Excel: {e}")
        return None, None


def load_post_text_from_excel(excel_file):
    """Загружает текст оригинального поста из Excel для сравнения"""
    try:
        df = None
        for sheet_name in ['Groups', 'Posts']:
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                break
            except:
                continue
        
        if df is None or 'Post Text' not in df.columns:
            logging.warning("⚠️ Колонка 'Post Text' не найдена в Excel")
            return None
        
        # Берём первый непустой текст
        for _, row in df.iterrows():
            text = str(row.get('Post Text', '')).strip()
            if text and text != 'nan' and len(text) > 10:
                logging.info(f"✅ Текст поста найден: {text[:60]}...")
                return text
        
        logging.warning("⚠️ Текст поста не найден в Excel")
        return None
    except Exception as e:
        logging.warning(f"⚠️ Ошибка чтения текста: {e}")
        return None


def get_post_text_from_activity_log(driver, post_url):
    """Читает текст поста из Activity Log"""
    try:
        # Текст виден прямо в Activity Log без открытия поста
        # Ищем текст рядом с ссылкой поста
        elements = driver.find_elements(By.XPATH,
            "//div[contains(@class,'story') or contains(@class,'feed')] | "
            "//div[@role='article'] | "
            "//div[contains(@data-testid,'story')]"
        )
        for elem in elements:
            try:
                text = elem.text.strip()
                if text and len(text) > 30:
                    return text[:300]
            except:
                continue
        return ""
    except:
        return ""


def texts_are_similar(text1, text2, min_chars=30):
    """
    Сравнивает два текста.
    Возвращает True если они похожи (один содержится в другом).
    """
    try:
        if not text1 or not text2:
            return False
        
        # Берём первые min_chars символов для сравнения
        t1 = text1[:min_chars].lower().strip()
        t2 = text2[:min_chars].lower().strip()
        
        # Проверяем вхождение
        if t1 in t2.lower() or t2 in t1.lower():
            return True
        
        # Проверяем совпадение первых слов
        words1 = set(t1.split()[:5])
        words2 = set(t2.split()[:5])
        common = words1 & words2
        
        return len(common) >= 3
    except:
        return False


def load_report_info_from_excel(excel_file):
    """Загружает всю информацию из Excel для репорта"""
    info = {
        'project': '',
        'post_name': '',
        'total_posts': 0,
        'first_post_date': '',
        'original_url': '',
        'selected_categories': []
    }
    
    try:
        df = pd.read_excel(excel_file, sheet_name='Summary')
        for _, row in df.iterrows():
            param = str(row.get('Параметр', '')).strip()
            value = row.get('Значение', '')
            
            if param == 'Проект:':
                info['project'] = str(value)
            elif param == 'Название поста:':
                info['post_name'] = str(value)
            elif param == '✅ Успешно:':
                info['total_posts'] = int(value) if pd.notna(value) else 0
            elif param == 'Первый шер:':
                info['first_post_date'] = str(value)
            elif param == 'URL оригинала:':
                info['original_url'] = str(value)
        
        logging.info(f"✅ Инфо из Excel: {info['project']} / {info['post_name']} / {info['total_posts']} постов")
        return info
        
    except Exception as e:
        logging.warning(f"⚠️ Ошибка чтения Summary: {e}")
        return info


def images_are_similar(hash1, hash2, threshold=HASH_SIMILARITY_THRESHOLD):
    """
    Сравнивает два хеша.
    Возвращает True если картинки похожи (разница <= threshold).
    """
    try:
        if hash1 is None or hash2 is None:
            return False
        
        # Если хеши - строки, конвертируем
        if isinstance(hash1, str):
            hash1 = imagehash.hex_to_hash(hash1)
        if isinstance(hash2, str):
            hash2 = imagehash.hex_to_hash(hash2)
        
        difference = hash1 - hash2
        return difference <= threshold
        
    except Exception as e:
        logging.debug(f"⚠️ Ошибка сравнения хешей: {e}")
        return False


# ================= ДРАЙВЕР =================
def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--start-maximized")
    options.add_argument("--window-size=1920,1080")
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        logging.info("✅ Chrome запущен")
        return driver
    except Exception as e:
        logging.error(f"❌ Ошибка Chrome: {e}")
        return None


# ================= ПАРСИНГ ДАТ =================
def parse_relative_time(text):
    """Парсинг относительного времени"""
    try:
        text = str(text).lower().strip()
        if 'just now' in text or 'только что' in text:
            now = datetime.now()
            return now.strftime('%d.%m.%Y'), now.strftime('%H:%M')
        
        minutes_match = re.search(r'(\d+)\s*(minutes?|мин)', text)
        if minutes_match:
            mins = int(minutes_match.group(1))
            t = datetime.now() - timedelta(minutes=mins)
            return t.strftime('%d.%m.%Y'), t.strftime('%H:%M')
        
        hours_match = re.search(r'(\d+)\s*(hours?|час)', text)
        if hours_match:
            hours = int(hours_match.group(1))
            t = datetime.now() - timedelta(hours=hours)
            return t.strftime('%d.%m.%Y'), t.strftime('%H:%M')
        
        days_match = re.search(r'(\d+)\s*(days?|дн)', text)
        if days_match:
            days = int(days_match.group(1))
            t = datetime.now() - timedelta(days=days)
            return t.strftime('%d.%m.%Y'), t.strftime('%H:%M')
        
        return None, None
    except:
        return None, None


def parse_absolute_date(text):
    """Парсинг абсолютной даты"""
    try:
        text = str(text).strip()
        patterns = [
            (r'(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})', 'dmy'),
            (r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})', 'mdy'),
            (r'(\d{1,2})\.(\d{1,2})\.(\d{4})', 'dmy_dot'),
            (r'(\d{1,2})/(\d{1,2})/(\d{4})', 'dmy_slash'),
        ]
        
        for pattern, ptype in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                try:
                    if ptype == 'dmy':
                        day, month, year = match.groups()
                        d = datetime.strptime(f"{day} {month[:3]} {year}", "%d %b %Y")
                    elif ptype == 'mdy':
                        month, day, year = match.groups()
                        d = datetime.strptime(f"{day} {month[:3]} {year}", "%d %b %Y")
                    elif ptype == 'dmy_dot':
                        day, month, year = match.groups()
                        d = datetime.strptime(f"{day}.{month}.{year}", "%d.%m.%Y")
                    elif ptype == 'dmy_slash':
                        day, month, year = match.groups()
                        d = datetime.strptime(f"{day}/{month}/{year}", "%d/%m/%Y")
                    return d.strftime('%d.%m.%Y'), "Unknown"
                except:
                    continue
        
        return None, None
    except:
        return None, None


def extract_post_time(driver):
    """Извлечение даты поста с Facebook"""
    try:
        # Способ 1: data-utime timestamp
        try:
            elem = driver.find_element(By.XPATH, "//*[@data-utime]")
            data_utime = elem.get_attribute('data-utime')
            if data_utime:
                dt = datetime.fromtimestamp(int(data_utime))
                return dt.strftime('%d.%m.%Y'), dt.strftime('%H:%M')
        except:
            pass
        
        # Способ 2: Текст с датой в элементах
        elems = driver.find_elements(By.XPATH,
            "//a[@data-utime] | //time | //*[contains(text(), '202') or contains(text(), 'ago') or contains(text(), 'minutes')]"
        )
        for elem in elems:
            try:
                text = elem.text.strip()
                if not text:
                    continue
                
                d, t = parse_absolute_date(text)
                if d:
                    return d, "Unknown"
                
                d, t = parse_relative_time(text)
                if d:
                    return d, t
            except:
                continue
        
        return datetime.now().strftime('%d.%m.%Y'), datetime.now().strftime('%H:%M')
    except:
        return datetime.now().strftime('%d.%m.%Y'), datetime.now().strftime('%H:%M')


# ================= ПОИСК ПОСТОВ =================
MONTHS = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}

# Если текста в записи Activity Log больше этого — считаем, что текст поста виден
# и можно сверять ключевые слова уже на этом этапе. Иначе пост проверится
# после открытия (process_post_with_hash).
TEXT_VISIBLE_MIN_LEN = 30

# JS проходит по DOM В ПОРЯДКЕ ДОКУМЕНТА и возвращает вперемешку заголовки дат
# и записи с кнопкой View. Для каждой записи берётся ТОЛЬКО её собственный
# контейнер (поднимаемся вверх, пока внутри ровно одна кнопка View) — иначе
# внешний div со всеми постами содержит слова из любых постов.
SCAN_LOG_JS = r"""
const MONTH = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b/i;
const isView = (a) => {
  const label = (a.innerText || '').trim();
  const aria = a.getAttribute('aria-label') || '';
  return (label.length < 20 && /view/i.test(label)) || /^view/i.test(aria);
};
const isPostHref = (h) => /facebook\.com\/groups|share\/[rvp]\//.test(h || '');
const out = [];
const nodes = document.querySelectorAll('a[href], h1, h2, h3, h4, h5, span, div');
for (const el of nodes) {
  if (el.tagName === 'A') {
    if (!isView(el)) continue;
    let c = el;
    while (c.parentElement && c.parentElement !== document.body) {
      const p = c.parentElement;
      let n = 0;
      for (const a of p.querySelectorAll('a')) { if (isView(a)) n++; }
      if (n > 1) break;
      if ((p.innerText || '').length > 800) break;
      c = p;
    }
    let ok = isPostHref(el.href);
    if (!ok) { for (const a of c.querySelectorAll('a')) { if (isPostHref(a.href)) { ok = true; break; } } }
    if (!ok) continue;
    out.push({kind: 'view', href: el.href, text: c.innerText || ''});
  } else if (el.childElementCount === 0) {
    const t = (el.textContent || '').trim();
    if (t.length >= 3 && t.length <= 30 &&
        (MONTH.test(t) || /^(today|yesterday|сегодня|вчера)$/i.test(t) ||
         /^\d{1,2}[./]\d{1,2}[./]\d{4}$/.test(t))) {
      out.push({kind: 'date', text: t});
    }
  }
}
return out;
"""


def parse_log_date(text):
    """
    Парсит заголовок даты Activity Log ('October 10, 2026', '10 October',
    'Oct 10', '10.10.2026', 'Today'). Возвращает datetime или None.
    Если год не указан — берём текущий (или прошлый, если дата получилась в будущем).
    """
    try:
        t = str(text).strip().lower().rstrip('.')
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        if t in ('today', 'сегодня'):
            return today
        if t in ('yesterday', 'вчера'):
            return today - timedelta(days=1)

        day = month = year = None
        m = re.fullmatch(r'(\d{1,2})[./](\d{1,2})[./](\d{4})', t)
        if m:
            day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        else:
            m = re.fullmatch(r'(\d{1,2})\s+([a-z]+)\.?(?:,?\s+(\d{4}))?', t)
            if m:
                day, mon, year = int(m.group(1)), m.group(2), m.group(3)
            else:
                m = re.fullmatch(r'([a-z]+)\.?\s+(\d{1,2})(?:,?\s+(\d{4}))?', t)
                if not m:
                    return None
                mon, day, year = m.group(1), int(m.group(2)), m.group(3)
            month = MONTHS.get(mon[:3])
            if not month:
                return None

        if year:
            return datetime(int(year), month, day)
        d = datetime(today.year, month, day)
        if d > today + timedelta(days=1):
            d = datetime(today.year - 1, month, day)
        return d
    except Exception:
        return None


def normalize_text(s):
    """Нижний регистр, без пунктуации, ё→е, одинарные пробелы"""
    s = str(s).lower().replace('ё', 'е')
    s = re.sub(r'[^\w\s]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def keywords_match(text, keywords):
    """
    Проверяет, что в тексте есть ВСЕ ключевые слова
    (если слов больше 3 — допускается пропуск одного).
    keywords — строка или список слов.
    """
    if isinstance(keywords, str):
        keywords = keywords.split()
    words = [w for w in normalize_text(' '.join(keywords)).split() if len(w) > 2]
    if not words:
        return True
    t = normalize_text(text)
    hits = sum(1 for w in words if w in t)
    required = len(words) if len(words) <= 3 else len(words) - 1
    return hits >= required


def scan_activity_log(driver):
    """
    Читает все записи Activity Log, загруженные на странице.
    Возвращает список {'href', 'text', 'date'} (date — datetime или None,
    если над записью не нашли заголовок даты).
    """
    try:
        items = driver.execute_script(SCAN_LOG_JS) or []
    except Exception as e:
        logging.warning(f"⚠️ Ошибка чтения Activity Log: {e}")
        return []

    entries = []
    current_date = None
    for it in items:
        if it.get('kind') == 'date':
            d = parse_log_date(it.get('text', ''))
            if d:
                current_date = d
        else:
            entries.append({
                'href': it.get('href'),
                'text': it.get('text', ''),
                'date': current_date,
            })
    return entries


def extract_group_name_from_text(full_text):
    """Название группы из текста записи ("posted in X")"""
    try:
        if 'posted in' in full_text:
            return full_text.split('posted in')[-1].split('\n')[0].strip()[:60]
        if 'By Pavel' in full_text:
            lines = [l.strip() for l in full_text.split('\n') if l.strip()]
            return lines[0][:60] if lines else ''
    except Exception:
        pass
    return ''


def date_str_to_obj(date_str):
    """Конвертирует строку даты '13.06.2026' в объект datetime"""
    try:
        return datetime.strptime(date_str, '%d.%m.%Y')
    except:
        return None


def collect_all_posts(driver, max_attempts=MAX_SCROLL_ATTEMPTS, stop_date=None,
                      search_words=None, end_date=None):
    """
    Скроллит Activity Log сверху вниз (новые → старые) и для КАЖДОЙ записи:
      1. определяет её дату по заголовку секции,
      2. пропускает записи позже end_date (слишком новые),
      3. останавливает сбор, когда дошли до записей раньше stop_date,
      4. если текст записи виден — сверяет ключевые слова.
    Возвращает список (view_href, group_name, 'ДД.ММ.ГГГГ' или '').
    """
    matched_posts = []
    seen_hrefs = set()
    earliest = date_str_to_obj(stop_date) if stop_date else None
    latest = date_str_to_obj(end_date) if end_date else None
    skipped_late = skipped_kw = undated = 0
    reached_older = False
    bottom_stalls = 0
    last_height = 0

    logging.info("📜 Скроллю в самый верх Activity Log...")
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(2)
    logging.info(f"📜 Скроллинг сверху вниз, беру записи за период "
                 f"{stop_date or '...'} — {end_date or '...'}")

    for attempt in range(max_attempts):
        for e in scan_activity_log(driver):
            href, d = e['href'], e['date']
            if not href or href in seen_hrefs:
                continue
            seen_hrefs.add(href)

            if d is not None:
                if latest and d > latest:
                    skipped_late += 1      # опубликовано/зашерено ПОСЛЕ периода
                    continue
                if earliest and d < earliest:
                    reached_older = True   # раньше периода — игнорируем
                    continue
            else:
                undated += 1

            text = e['text']
            if search_words and len(text.strip()) > TEXT_VISIBLE_MIN_LEN \
                    and not keywords_match(text, search_words):
                skipped_kw += 1            # другой пост (другие слова)
                continue

            group_name = extract_group_name_from_text(text)
            d_str = d.strftime('%d.%m.%Y') if d else ''
            matched_posts.append((href, group_name, d_str))
            logging.info(f"   ✅ {d_str or '??.??.????'} | {group_name[:50] or href[:50]}")

        if reached_older:
            logging.info("   ✅ Дошли до записей старее начала периода — останавливаюсь")
            break

        # Скролл на 80% высоты окна
        try:
            h, y, win_h = driver.execute_script(
                "const h = document.documentElement.scrollHeight;"
                "const y = window.scrollY;"
                "window.scrollBy(0, Math.floor(window.innerHeight * 0.8));"
                "return [h, y, window.innerHeight];"
            )
        except Exception:
            break
        time.sleep(SCROLL_WAIT)

        at_bottom = (y + win_h) >= (h - 5)
        bottom_stalls = bottom_stalls + 1 if (at_bottom and h == last_height) else 0
        last_height = h
        if bottom_stalls >= 3:
            logging.info("   ✅ Дошли до конца Activity Log")
            break

        if attempt % 10 == 9:
            logging.info(f"   ⏳ Шаг {attempt + 1}/{max_attempts}, найдено: {len(matched_posts)}")

    logging.info(f"   ✅ Итого: {len(matched_posts)} | пропущено позже {end_date}: {skipped_late} "
                 f"| пропущено по ключевым словам: {skipped_kw} | без даты: {undated}")
    if undated:
        logging.warning(f"   ⚠️ {undated} записей без заголовка даты — фильтр по дате к ним не применён")
    return matched_posts


# ================= ОБРАБОТКА ПОСТОВ =================
def extract_post_stats(driver):
    """Статистика поста"""
    stats = {'likes': 0, 'comments': 0, 'shares': 0}
    try:
        elems = driver.find_elements(By.XPATH,
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'like')] | "
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'comment')] | "
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'share')]"
        )
        for elem in elems:
            try:
                text = elem.text.strip()
                if not text:
                    continue
                nums = re.findall(r'\d+', text)
                if nums:
                    n = int(nums[0])
                    tl = text.lower()
                    if 'comment' in tl or 'комментар' in tl:
                        stats['comments'] = max(stats['comments'], n)
                    elif 'share' in tl or 'поделился' in tl:
                        stats['shares'] = max(stats['shares'], n)
                    elif 'like' in tl or 'нравится' in tl:
                        stats['likes'] = max(stats['likes'], n)
            except:
                continue
    except:
        pass
    return stats


def get_group_name(driver):
    """Название группы из popup поста"""
    try:
        # Способ 1: Читаем название из заголовка popup
        # Facebook показывает "BARD PRoductions's Post" вверху
        # Но нам нужно название ГРУППЫ где опубликован пост
        
        # Ищем название группы в popup
        group_selectors = [
            # Название группы в заголовке поста
            "//div[@role='dialog']//a[contains(@href, '/groups/')]",
            "//div[@aria-modal='true']//a[contains(@href, '/groups/')]",
            # Заголовок popup (название группы)
            "//div[@role='dialog']//h2",
            "//div[@role='dialog']//strong",
            # Ссылки на группы на странице
            "//a[contains(@href, '/groups/') and not(contains(@href, '/members')) and not(contains(@href, '/posts'))]",
        ]
        
        for selector in group_selectors:
            try:
                elems = driver.find_elements(By.XPATH, selector)
                for elem in elems:
                    text = elem.text.strip()
                    if (text and 
                        2 < len(text) < 150 and 
                        'facebook' not in text.lower() and
                        'groups' not in text.lower() and
                        text not in ['Post', 'Share', 'Like', 'Comment']):
                        return text
            except:
                continue
        
        # Способ 2: Из URL поста извлекаем ID группы
        try:
            current_url = driver.current_url
            group_match = re.search(r'/groups/([^/]+)/', current_url)
            if group_match:
                group_id = group_match.group(1)
                # Ищем название группы по этому ID на странице
                elems = driver.find_elements(By.XPATH,
                    f"//a[contains(@href, '/groups/{group_id}')]"
                )
                for elem in elems:
                    text = elem.text.strip()
                    if text and 2 < len(text) < 150:
                        return text
        except:
            pass
            
    except:
        pass
    return "Unknown Group"


def process_post_with_hash(driver, post_url, screenshots_folder, index, total, original_hash=None, original_text=None, group_name='', log_date=''):
    """
    НОВАЯ ЛОГИКА (июль 2026):
    Текст уже проверен в Activity Log — здесь только открываем пост и делаем скриншот.
    post_url = href кнопки View (уже отфильтрованный совпавший пост)
    """
    result = {
        'success': False,
        # Если есть ключевые слова — решаем по тексту открытого поста (ниже)
        'is_matching': not original_text,
        'match_reason': 'activity_log_text',
        'group_name': group_name or 'Unknown',
        'url': post_url,
        'posted_date': log_date or datetime.now().strftime('%d.%m.%Y'),
        'posted_time': '' if log_date else datetime.now().strftime('%H:%M'),
        'stats': {'likes': 0, 'comments': 0, 'shares': 0},
        'screenshot': None,
        'image_hash': None
    }
    
    try:
        logging.info(f"\n[{index}/{total}] Открываю пост...")
        driver.set_page_load_timeout(20)  # максимум 20 сек на загрузку
        try:
            driver.get(post_url)
        except Exception:
            # Timeout загрузки — продолжаем с тем что есть
            pass
        time.sleep(5)  # увеличено с 3 до 5 сек — ждём загрузки popup
        
        # === ШАГ 1: БЫСТРАЯ ПРОВЕРКА ТЕКСТА (без скриншота) ===
        # Сначала определяем совпадает ли пост — и только потом
        # делаем скриншот и собираем данные (экономит время)

        # Нажимаем "See more" если есть
        try:
            see_more_buttons = driver.find_elements(By.XPATH,
                "//*[text()='See more' or text()='Ещё' or text()='See More']"
            )
            for btn in see_more_buttons:
                try:
                    btn.click()
                    time.sleep(0.5)
                    break
                except:
                    continue
        except:
            pass

        # Читаем полный текст popup
        post_text = ""
        try:
            popup = driver.find_element(By.XPATH,
                "//div[@role='dialog'] | //div[@aria-modal='true']"
            )
            post_text = popup.text
        except:
            pass

        # Если текст пустой — кликаем кнопку "View" чтобы открыть полный пост
        if not post_text.strip():
            try:
                view_btn = driver.find_element(By.XPATH,
                    "//a[contains(text(),'View') or contains(text(),'Посмотреть') or contains(text(),'View post')]"
                    " | //span[contains(text(),'View') or contains(text(),'Посмотреть')]/.."
                )
                view_btn.click()
                time.sleep(3)
                # Читаем текст снова после открытия полного поста
                try:
                    popup = driver.find_element(By.XPATH,
                        "//div[@role='dialog'] | //div[@aria-modal='true']"
                    )
                    post_text = popup.text
                except:
                    post_text = driver.find_element(By.TAG_NAME, 'body').text
            except:
                pass

        # Совсем пусто — читаем всю страницу
        if not post_text.strip():
            try:
                post_text = driver.find_element(By.TAG_NAME, 'body').text
            except:
                pass

        # СПОСОБ 1: Сравнение по ТЕКСТУ (все ключевые слова должны быть в посте)
        if original_text and len(original_text) > 3:
            print(f"   [{index}/{total}] Текст на странице: {post_text[:80].strip()!r}")

            if not post_text.strip():
                result['is_matching'] = True
                result['match_reason'] = 'unverified'
                print("   ⚠️ Текст прочитать не удалось — беру пост без проверки")
            elif keywords_match(post_text, original_text):
                result['is_matching'] = True
                result['match_reason'] = 'text'
                print(f"   ✅ СОВПАДЕНИЕ: '{original_text[:30]}'")
            else:
                print(f"   ⏭️ Другой пост (нет слов '{original_text[:30]}') — пропускаю")

        # Если нет фильтра — берём все посты
        if not original_text and not original_hash:
            result['is_matching'] = True
            result['match_reason'] = 'no_filter'

        # === ШАГ 2: СКРИНШОТ И ДАННЫЕ ТОЛЬКО ДЛЯ СОВПАВШИХ ПОСТОВ ===
        # Несовпавшие пропускаем — экономит ~4 сек на каждый пост
        if result['is_matching']:
            # Скриншот
            try:
                group_name_safe = re.sub(r'[^\w\-]', '_', result.get('group_name', 'unknown'))[:30]
                screenshot_filename = f"{index:03d}_{group_name_safe}.png"
                screenshot_path = os.path.join(screenshots_folder, screenshot_filename)
                driver.save_screenshot(screenshot_path)
                result['screenshot'] = screenshot_path
            except Exception as e:
                logging.warning(f"   ⚠️ Ошибка скриншота: {e}")

            # СПОСОБ 2: Хеш (запасной, только если текст не совпал)
            if result['match_reason'] != 'text' and original_hash:
                current_hash = calculate_perceptual_hash(result.get('screenshot', ''))
                result['image_hash'] = str(current_hash) if current_hash else None
                if current_hash and images_are_similar(original_hash, current_hash):
                    result['is_matching'] = True
                    result['match_reason'] = 'image_hash'
                    logging.info(f"   ✅ СОВПАДЕНИЕ по ХЕШУ!")

            # Данные поста
            page_group = get_group_name(driver)
            if page_group != "Unknown Group" or not result['group_name'] \
                    or result['group_name'] == 'Unknown':
                result['group_name'] = page_group
            if log_date:
                # Дата из Activity Log надёжнее, чем разбор страницы поста
                result['posted_date'], result['posted_time'] = log_date, ''
            else:
                result['posted_date'], result['posted_time'] = extract_post_time(driver)
            result['stats'] = extract_post_stats(driver)
            result['success'] = True

            logging.info(f"   📅 {result['posted_date']} {result['posted_time']}")
            logging.info(f"   👍 {result['stats']['likes']} | 💬 {result['stats']['comments']} | ↗️ {result['stats']['shares']}")
        else:
            result['success'] = True  # пост обработан, просто не совпал
        
    except Exception as e:
        logging.error(f"   ❌ Ошибка: {e}")
    
    return result


# ================= EXCEL И PDF =================
def create_report_excel(report_folder, project_name, post_name, date_range):
    """Создание Excel для отчёта"""
    folder_name = os.path.basename(report_folder)
    excel_file = os.path.join(report_folder, f"{folder_name}_report.xlsx")
    
    summary_data = {
        'Параметр': [
            'Проект:', 'Пост:', 'Дата/Период:', 'Дата создания:',
            '', 'Всего постов:'
        ],
        'Значение': [
            project_name, post_name, date_range, datetime.now().strftime('%d.%m.%Y %H:%M'),
            '', 0
        ]
    }
    
    df_summary = pd.DataFrame(summary_data)
    df_posts = pd.DataFrame(columns=[
        '№', 'Group Name', 'Posted Date', 'URL', 'Screenshot'
    ])
    
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='Summary', index=False)
        df_posts.to_excel(writer, sheet_name='Groups', index=False)
    
    return excel_file


def save_results_to_excel(excel_file, matching_results):
    """Сохранение только СОВПАДАЮЩИХ постов в Excel"""
    try:
        df_posts = pd.DataFrame([
            {
                '№': i,
                'Group Name': r['group_name'],
                'Posted Date': r.get('posted_date', ''),
                'URL': r.get('url', ''),
                'Screenshot': os.path.basename(r['screenshot']) if r['screenshot'] else '—'
            }
            for i, r in enumerate(matching_results, 1)
        ])
        
        df_summary = pd.read_excel(excel_file, sheet_name='Summary')
        total = len(matching_results)
        
        updates = {
            'Всего постов:': total,
        }
        
        for param, value in updates.items():
            mask = df_summary['Параметр'] == param
            if mask.any():
                df_summary.loc[mask, 'Значение'] = value
        
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            df_posts.to_excel(writer, sheet_name='Groups', index=False)
        
        logging.info(f"✅ Excel сохранён: {excel_file}")
        
    except Exception as e:
        logging.error(f"❌ Ошибка Excel: {e}")


def generate_pdf(excel_file):
    """
    Генерация PDF отчёта со скриншотами.

    Структура PDF:
    - Страница 1: Титульная (сводная таблица + итоги)
    - Страницы 2+: Один пост = одна страница (новые → старые)
                   Заголовок группы, дата/время, статистика, скриншот

    Решение DPI проблемы:
    - Chrome сохраняет скриншоты в реальных пикселях (1920×1080)
    - НЕ используем img.thumbnail() — оно ухудшает качество
    - Используем оригинальный PNG без изменения размера
    - Вычисляем размер в PDF через: pixel_size / 96 * 72
      (96 DPI — стандарт экрана, 72 DPI — ReportLab)
    - Ограничиваем размер страницей A4 (595×842 points)
    """
    try:
        pdf_file = excel_file.replace('.xlsx', '.pdf')
        df = pd.read_excel(excel_file, sheet_name='Groups')

        # === РЕГИСТРАЦИЯ ШРИФТА (поддержка кириллицы) ===
        try:
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont

            font_paths = [
                'C:/Windows/Fonts/arial.ttf',
                'C:/Windows/Fonts/Arial.ttf',
                'C:/Windows/Fonts/calibri.ttf',
                'C:/Windows/Fonts/times.ttf',
            ]

            font_name = 'Helvetica'  # fallback
            for font_path in font_paths:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('CyrillicFont', font_path))
                    font_name = 'CyrillicFont'
                    logging.info(f"   ✅ Шрифт загружен: {font_path}")
                    break
            else:
                logging.warning("   ⚠️ Кириллический шрифт не найден, используем Helvetica")
        except Exception as e:
            font_name = 'Helvetica'
            logging.warning(f"   ⚠️ Ошибка шрифта: {e}")

        # === СТИЛИ ===
        # A4: 595 × 842 points, поля 1cm = ~28 points
        # Рабочая область: ~540 × 790 points = ~7.5 × 10.97 inch
        PAGE_W = 7.5 * inch   # рабочая ширина страницы
        PAGE_H = 10.0 * inch  # рабочая высота страницы

        title_style = ParagraphStyle(
            'CyrTitle', fontName=font_name, fontSize=18,
            textColor=colors.HexColor('#1877F2'), spaceAfter=10, spaceBefore=0
        )
        heading_style = ParagraphStyle(
            'CyrHeading', fontName=font_name, fontSize=12,
            textColor=colors.HexColor('#333333'), spaceAfter=4, spaceBefore=0
        )
        normal_style = ParagraphStyle(
            'CyrNormal', fontName=font_name, fontSize=9, spaceAfter=3
        )
        small_style = ParagraphStyle(
            'CyrSmall', fontName=font_name, fontSize=8,
            textColor=colors.HexColor('#666666'), spaceAfter=2
        )

        def safe_text(val, max_len=None):
            try:
                s = str(val) if val is not None else ''
                return s[:max_len] if max_len else s
            except:
                return ''

        def safe_int(val):
            try:
                return int(float(val))
            except:
                return 0

        # === СОЗДАНИЕ PDF ===
        doc = SimpleDocTemplate(
            pdf_file,
            pagesize=A4,
            leftMargin=1*cm,
            rightMargin=1*cm,
            topMargin=1*cm,
            bottomMargin=1*cm
        )
        story = []

        # ============================================================
        # СТРАНИЦА 1: Сводная таблица
        # ============================================================
        story.append(Paragraph("Facebook Posting Report", title_style))
        story.append(Paragraph(
            f"Generated: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
            normal_style
        ))
        story.append(Spacer(1, 0.15*inch))

        if len(df) > 0:
            # Заголовочная статистика
            stats_text = f"Всего постов: {len(df)}"
            story.append(Paragraph(stats_text, normal_style))
            story.append(Spacer(1, 0.1*inch))

            # Сводная таблица постов
            table_data = [[
                Paragraph('№', normal_style),
                Paragraph('Group', normal_style),
            ]]
            for idx, (_, row) in enumerate(df.iterrows(), 1):
                table_data.append([
                    Paragraph(str(idx), normal_style),
                    Paragraph(safe_text(row.get('Group Name', ''), 60), normal_style),
                ])

            col_widths = [
                0.35*inch,  # №
                6.65*inch,  # Group (расширили на всю ширину)
            ]
            tbl = Table(table_data, colWidths=col_widths)
            tbl.setStyle(TableStyle([
                ('BACKGROUND',    (0, 0), (-1, 0), colors.HexColor('#1877F2')),
                ('TEXTCOLOR',     (0, 0), (-1, 0), colors.white),
                ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME',      (0, 0), (-1, -1), font_name),
                ('FONTSIZE',      (0, 0), (-1, -1), 8),
                ('GRID',          (0, 0), (-1, -1), 0.4, colors.grey),
                ('ROWBACKGROUNDS',(0, 1), (-1, -1), [colors.white, colors.HexColor('#F0F2F5')]),
                ('TOPPADDING',    (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]))
            story.append(tbl)

        else:
            story.append(Paragraph("No posts found", normal_style))

        # ============================================================
        # СТРАНИЦЫ 2+: Один пост = одна страница (от новых к старым)
        # ============================================================
        if len(df) > 0:
            # df уже в порядке обнаружения (от новых к старым — Activity Log идёт так)
            for idx, (_, row) in enumerate(df.iterrows(), 1):
                story.append(PageBreak())

                group_name   = safe_text(row.get('Group Name', 'Unknown'), 60)
                post_url     = safe_text(row.get('URL', ''))
                posted_date  = safe_text(row.get('Posted Date', ''))
                posted_time  = safe_text(row.get('Posted Time', ''))
                screenshot   = safe_text(row.get('Screenshot', ''))

                # Заголовок
                story.append(Paragraph(
                    f"#{idx}  {group_name}",
                    heading_style
                ))

                # URL (мелко)
                if post_url:
                    story.append(Paragraph(post_url[:80], small_style))

                story.append(Spacer(1, 0.08*inch))

                # === СКРИНШОТ ===
                # Ищем файл скриншота рядом с Excel
                report_folder = os.path.dirname(excel_file)
                screenshot_path = None

                if screenshot:
                    # Сначала пробуем как абсолютный путь
                    if os.path.exists(screenshot):
                        screenshot_path = screenshot
                    else:
                        # Пробуем как имя файла в папке screenshots/
                        candidate = os.path.join(report_folder, 'screenshots', screenshot)
                        if os.path.exists(candidate):
                            screenshot_path = candidate
                        else:
                            # Пробуем прямо в папке отчёта
                            candidate2 = os.path.join(report_folder, screenshot)
                            if os.path.exists(candidate2):
                                screenshot_path = candidate2

                if screenshot_path and os.path.exists(screenshot_path):
                    try:
                        from PIL import Image as PILImage

                        img = PILImage.open(screenshot_path)
                        img_w_px, img_h_px = img.size

                        # === КЛЮЧЕВАЯ ЛОГИКА DPI ===
                        # Chrome делает скриншоты в экранных пикселях.
                        # На Windows с 100% масштабом: 1 CSS пиксель = 1 реальный пиксель.
                        # ReportLab измеряет в points (72 points = 1 inch).
                        # Экранный DPI обычно 96 (для Windows).
                        # Формула: size_in_points = px / 96 * 72 = px * 0.75
                        SCREEN_DPI = 96
                        REPORTLAB_DPI = 72
                        scale = REPORTLAB_DPI / SCREEN_DPI  # = 0.75

                        img_w_pt = img_w_px * scale  # в points
                        img_h_pt = img_h_px * scale  # в points

                        # Доступная область на странице A4
                        # (высота - заголовок ~80pt - поля)
                        avail_w = PAGE_W   # ~540 points
                        avail_h = PAGE_H - 80  # резерв на заголовок

                        # Масштабируем пропорционально если не вмещается
                        scale_w = avail_w / img_w_pt if img_w_pt > avail_w else 1.0
                        scale_h = avail_h / img_h_pt if img_h_pt > avail_h else 1.0
                        final_scale = min(scale_w, scale_h)

                        final_w = img_w_pt * final_scale
                        final_h = img_h_pt * final_scale

                        # Используем оригинальный файл — НЕ изменяем размер через PIL
                        # (это и решает проблему DPI — нет пересэмплинга)
                        rl_img = RLImage(screenshot_path, width=final_w, height=final_h)
                        story.append(rl_img)

                        logging.info(
                            f"   📸 Скриншот #{idx}: {img_w_px}×{img_h_px}px → "
                            f"{final_w:.0f}×{final_h:.0f}pt в PDF"
                        )

                    except Exception as img_err:
                        logging.warning(f"   ⚠️ Ошибка скриншота #{idx}: {img_err}")
                        story.append(Paragraph(f"[Screenshot error: {img_err}]", small_style))
                else:
                    story.append(Paragraph("[No screenshot]", small_style))
                    if screenshot:
                        logging.warning(f"   ⚠️ Файл не найден: {screenshot}")

        doc.build(story)
        logging.info(f"✅ PDF создан: {pdf_file}")
        return pdf_file

    except Exception as e:
        logging.warning(f"⚠️ Ошибка PDF: {e}")
        import traceback
        logging.warning(traceback.format_exc())
        return None


# ================= ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ НОВОГО MAIN =================

def extract_keywords_from_post(driver, post_url):
    """
    Открывает оригинальный пост и извлекает:
    - Первые 3-4 слова текста (ключевые слова для поиска)
    - Количество шеров (справочно)

    Возвращает: (keywords_str, shares_count)
    """
    keywords = None
    shares = 0

    try:
        logging.info(f"Открываю оригинальный пост: {post_url}")
        driver.set_page_load_timeout(20)
        try:
            driver.get(post_url)
        except Exception:
            pass
        time.sleep(4)

        # Читаем текст поста
        post_text = ""
        try:
            # Пробуем popup/dialog
            popup = driver.find_element(By.XPATH,
                "//div[@role='dialog'] | //div[@aria-modal='true']"
            )
            post_text = popup.text
        except:
            pass

        if not post_text:
            try:
                post_text = driver.find_element(By.TAG_NAME, 'body').text
            except:
                pass

        # Извлекаем первые 3-4 слова (без эмодзи, без коротких)
        if post_text:
            # Убираем эмодзи и спецсимволы, оставляем буквы/цифры/пробелы
            clean = re.sub(r'[^\w\s]', ' ', post_text, flags=re.UNICODE)
            words = [w for w in clean.split() if len(w) > 2][:4]
            if words:
                keywords = ' '.join(words)
                logging.info(f"   🔑 Ключевые слова: {keywords}")

        # Читаем количество шеров
        try:
            body_text = driver.find_element(By.TAG_NAME, 'body').text
            # Форматы: "75 shares", "75 Share", "Поделились: 75"
            m = re.search(r'(\d+)\s*share', body_text, re.IGNORECASE)
            if not m:
                m = re.search(r'поделил[^\d]*(\d+)', body_text, re.IGNORECASE)
            if m:
                shares = int(m.group(1))
                logging.info(f"   🔁 Шеров: {shares}")
        except:
            pass

    except Exception as e:
        logging.warning(f"⚠️ Ошибка чтения поста: {e}")

    return keywords, shares


def input_date(prompt, default=None):
    """
    Запрашивает дату в формате ДД.ММ.ГГГГ.
    Возвращает строку даты или default если Enter без ввода.
    """
    while True:
        val = input(prompt).strip()
        if not val and default:
            return default
        try:
            datetime.strptime(val, '%d.%m.%Y')
            return val
        except ValueError:
            print("   ❌ Формат даты: ДД.ММ.ГГГГ (например 11.06.2026)")


# ================= MAIN =================
def main():
    print("\n" + "="*70)
    print("   FACEBOOK SCREENSHOT REPORT v8.0")
    print("   Ручной ввод | Диапазон дат | Скриншоты совпавших постов")
    print("="*70 + "\n")

    # ШАГ 1: КЛЮЧЕВЫЕ СЛОВА
    print("─"*70)
    print("🔑 КЛЮЧЕВЫЕ СЛОВА ИЗ ПОСТА")
    print("─"*70)
    print("   Открой оригинальный пост и введи первые 3-4 слова текста.")
    print("   Пример: ДЖАЗ НА ПИРСЕ МАШКИТ")
    print()
    keywords = input("👉 Ключевые слова: ").strip()
    if not keywords or len(keywords) < 3:
        print("⚠️ Ключевые слова не введены — буду собирать все посты за период.")
        keywords = None

    # ШАГ 2: КОЛИЧЕСТВО ШЕРОВ (справочно)
    shares = 0
    shares_str = input("👉 Количество шеров под постом (Enter чтобы пропустить): ").strip()
    if shares_str.isdigit():
        shares = int(shares_str)
        print(f"   🔁 Шеров: {shares} (справочно)")

    # Убираем знаки препинания из ключевых слов
    import re as _re
    if keywords:
        keywords_clean = _re.sub(r'[^\w\s]', ' ', keywords).strip()
        keywords_clean = ' '.join(keywords_clean.split())
        print(f"   🔍 Поиск по: '{keywords_clean}'")
    else:
        keywords_clean = None

    original_text = keywords_clean

    # ШАГ 3: ДИАПАЗОН ДАТ
    print("\n" + "─"*70)
    print("📅 ДИАПАЗОН ДАТ ПОСТИНГА")
    print("─"*70)
    print("   (формат: ДД.ММ.ГГГГ, например 11.06.2026)")
    today = datetime.now().strftime('%d.%m.%Y')
    earliest_date = input_date("👉 Дата С (самая ранняя): ")
    latest_date = input_date(f"👉 Дата ПО (самая поздняя) [{today}]: ", default=today)

    # Проверяем что earliest <= latest
    earliest_date_obj = date_str_to_obj(earliest_date)
    latest_date_obj = date_str_to_obj(latest_date)
    if earliest_date_obj and latest_date_obj and earliest_date_obj > latest_date_obj:
        print("⚠️ Дата С больше чем По — меняю местами")
        earliest_date, latest_date = latest_date, earliest_date
        earliest_date_obj, latest_date_obj = latest_date_obj, earliest_date_obj

    print(f"\n✅ Период: {earliest_date} — {latest_date}")
    if keywords:
        print(f"✅ Ищем: \"{keywords}\"")

    # ШАГ 4: ВЫБОР СТРАНИЦЫ И СОЗДАЁМ ПАПКУ ДЛЯ ОТЧЁТА
    PAGES = {"1": "Bard", "2": "Liberman"}
    print("\n📂 Выберите страницу:")
    for k, v in PAGES.items():
        print(f"   [{k}] {v}")
    page_choice = input("👉 Номер страницы [1]: ").strip() or "1"
    page_name = PAGES.get(page_choice, "Bard")
    print(f"📂 Страница: {page_name}")

    post_name_clean = clean_filename(keywords[:20] if keywords else 'Report')
    report_folder = os.path.join(
        REPORTS_FOLDER,
        f"{page_name}_{post_name_clean}_REPORT_{datetime.now().strftime('%d%b%Y_%H%M')}"
    )
    os.makedirs(report_folder, exist_ok=True)
    screenshots_folder = os.path.join(report_folder, "screenshots")
    os.makedirs(screenshots_folder, exist_ok=True)

    date_range = f"{earliest_date} — {latest_date}" if earliest_date != latest_date else earliest_date
    new_excel = create_report_excel(report_folder, page_name, post_name_clean, date_range)

    # ШАГ 5: ЗАПУСК БРАУЗЕРА И ЛОГИН
    print("\n🚀 Запускаю Chrome...")
    driver = setup_driver()
    if not driver:
        return

    log_file = setup_logging(report_folder)
    logging.info(f"Проект: {page_name} | Слова: {keywords} | Период: {date_range} | Шеров: {shares}")

    driver.get("https://www.facebook.com")

    print("\n" + "!"*70)
    print("🔐 ЗАЛОГИНИТЕСЬ В FACEBOOK")
    print("!"*70)
    print("1. Залогинитесь в Facebook в открывшемся браузере")
    print(f"2. Перейдите на страницу {page_name}")
    print("!"*70)
    input("\n👉 ENTER когда залогинились...")

    # ШАГ 6: ACTIVITY LOG
    print("\n" + "!"*70)
    print("📜 ОТКРОЙТЕ ACTIVITY LOG")
    print("!"*70)
    print(f"1. Перейдите на страницу {page_name}")
    print("2. Откройте Activity Log → Group posts and comments")
    print("3. Скролить вручную НЕ нужно — скрипт сам пролистает с самого верха")
    print(f"   и возьмёт только записи за {earliest_date} — {latest_date}")
    print("4. Нажмите ENTER")
    print("!"*70)
    if earliest_date != latest_date:
        print(f"\n⚠️  Посты за НЕСКОЛЬКО дней: {earliest_date} — {latest_date}")

    input("\n👉 ENTER когда Activity Log открыт...")

    # ШАГ 7: СБОР ПОСТОВ
    print("\n🔍 Собираю посты из Activity Log...")
    print(f"   📅 Период: {earliest_date} — {latest_date} (записи вне периода пропускаются)")

    # Разбиваем ключевые слова для поиска в Activity Log
    search_words = original_text.split() if original_text else []
    all_post_urls = collect_all_posts(driver, stop_date=earliest_date, end_date=latest_date,
                                     search_words=search_words)

    if not all_post_urls:
        print("❌ Посты не найдены!")
        driver.quit()
        return

    print(f"\n✅ Найдено совпадений в Activity Log: {len(all_post_urls)}")
    by_date = {}
    for _, _, d in all_post_urls:
        by_date[d or 'без даты'] = by_date.get(d or 'без даты', 0) + 1
    for d in sorted(by_date, key=lambda x: (date_str_to_obj(x) or datetime.max)):
        print(f"   {d}: {by_date[d]}")
    if shares:
        print(f"🔁 Шеров на оригинальном посте: {shares} (справочно)")

    if input(f"\n👉 Начать скриншоты {len(all_post_urls)} совпавших постов? (y/n): ").strip().lower() != 'y':
        driver.quit()
        return

    # ШАГ 8: ОБРАБОТКА ПОСТОВ
    print("\n📸 Обрабатываю посты...\n")
    all_results = []
    matching_results = []

    def is_browser_alive(driver):
        """Проверяет жив ли браузер"""
        try:
            _ = driver.current_url
            return True
        except:
            return False

    def restart_browser(driver, post_url):
        """Перезапускает браузер и открывает нужный URL"""
        print("\n⚠️ Браузер упал! Перезапускаю...")
        try:
            driver.quit()
        except:
            pass
        new_driver = setup_driver()
        if new_driver:
            new_driver.get("https://www.facebook.com")
            print("🔐 Залогинитесь заново в Facebook!")
            input("👉 ENTER когда залогинились...")
            print("✅ Продолжаю обработку...")
        return new_driver

    for i, (url, group_name, log_date) in enumerate(all_post_urls, 1):
        # Проверяем жив ли браузер перед каждым постом
        if not is_browser_alive(driver):
            driver = restart_browser(driver, url)
            if not driver:
                print("❌ Не удалось перезапустить браузер!")
                break

        # Открываем пост и делаем скриншот (текст уже проверен в Activity Log)
        result = process_post_with_hash(
            driver, url, screenshots_folder, i, len(all_post_urls),
            original_text=original_text, group_name=group_name, log_date=log_date
        )
        all_results.append(result)

        if result['is_matching']:
            matching_results.append(result)
            reason = result.get('match_reason', '')
            reason_text = '(текст)' if reason == 'text' else '(не проверен)' if reason == 'unverified' else ''
            print(f"   ✅ #{len(matching_results)} {reason_text}: {result['group_name'][:40]}")

        if i < len(all_post_urls):
            time.sleep(2)

    # ШАГ 9: СОХРАНЕНИЕ
    print(f"\n💾 Сохраняю результаты...")
    save_results_to_excel(new_excel, matching_results)
    pdf_file = generate_pdf(new_excel)

    # ИТОГИ
    print("\n" + "="*70)
    print("🏁 ГОТОВО!")
    print("="*70)
    print(f"📊 Проверено постов: {len(all_results)}")
    print(f"✅ Найдено совпадений: {len(matching_results)}")
    if shares:
        print(f"🔁 Шеров на оригинале: {shares}")
        if matching_results:
            pct = round(len(matching_results) / shares * 100)
            print(f"   ({pct}% от шеров видны в Activity Log)")
    print()
    print(f"📄 Excel: {new_excel}")
    if pdf_file:
        print(f"📊 PDF:   {pdf_file}")
    print(f"📸 Скриншоты: {screenshots_folder}")
    print("="*70)

    input("\n👉 ENTER чтобы закрыть браузер...")
    driver.quit()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n⚠️ Прервано пользователем")
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
        import traceback
        traceback.print_exc()
