from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import time, random, re


ITEM_SELECTOR = 'div[data-marker="item"]'
PRICE_META_SELECTOR = 'meta[itemprop="price"]'
PRICE_DISPLAY_SELECTOR = 'span[data-marker="item-price-value"]'
ITEM_ID_ATTR = 'data-item-id'
LINK_SELECTOR = 'a[itemprop="url"]'
TITLE_SELECTOR = '[data-marker="item-title"]'
CONDITION_SELECTOR = '[data-marker="item-specific-params"]'
SELLER_SELECTOR = '[data-marker="seller-info/summary"]'

#функция воссздания названия excel файла
def make_filename(query, condition, city, price_filter, ext='xlsx'):
    def clean(s):
        if not s: return 'all'
        s = re.sub(r'[^\wа-яё\-]+', '_', str(s).lower().strip(), flags=re.I)
        s = re.sub(r'_+', '_', s).strip('_')
        return s or 'all'
    parts = [clean(query), clean(condition), clean(city), clean(price_filter)]
    return '_'.join(parts) + '.' + ext

#функция сохранения в excel файл
def save_to_excel(data, filename):
    if not filename.lower().endswith('.xlsx'):
        filename = filename.rsplit('.', 1)[0] + '.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"

    ws.append(["URL", "Title", "Price", "City", "Description"])
    for item in data:
        ws.append([
            item.get('url', ''),
            item.get('title', ''),
            item.get('price', ''),
            item.get('city', ''),
            item.get('description', '')
        ])
    wb.save(filename)
    return filename

def parse_price_from_meta(card):
    try:
        meta = card.find_element(By.CSS_SELECTOR, PRICE_META_SELECTOR)
        price_str = meta.get_attribute('content')
        return float(price_str) if price_str else None
    except:
        return None

#изъятие данных с карточки авито
def get_item_data_from_card(card, driver):
    data = {}
    try:
        try:
            link_el = card.find_element(By.CSS_SELECTOR, LINK_SELECTOR)
            data['url'] = link_el.get_attribute('href')
            data['title'] = link_el.get_attribute('title') or link_el.text.strip()
        except:
            data['url'] = ''
            data['title'] = ''
        
        price_num = parse_price_from_meta(card)
        data['price_numeric'] = price_num
    except:
        pass
    return data

import re

#функция фильтрации предмета
def filter_item(item_data, city, price_range, condition=None, price_tol=0.1):
    price_num = item_data.get('price_numeric')
    if price_range is not None and price_num is not None:
        if isinstance(price_range, (tuple, list)) and len(price_range) == 2:
            if not (price_range[0] <= price_num <= price_range[1]):
                return False
        elif isinstance(price_range, (int, float)):
            if abs(price_num - price_range) > price_range * price_tol:
                return False

    if city and str(city).lower().strip() not in ['вся россия', 'россия', 'all', '']:
        ad_city = str(item_data.get('city', '')).lower().strip()
        target = str(city).lower().strip()
        
        if not ad_city:
            return False  # если город не указан в объявлении - пропуск

        def clean(c):
            c = re.sub(r'\b(г\.|город|п\.|пос\.|с\.|д\.|ст\.|мкр\.|р-н|район|обл\.|область|край|ао|округ)\b\.?\s*', '', c, flags=re.I)
            return re.sub(r'[^\wа-яё\s\-]', '', c).strip()

        norm_target = clean(target)
        norm_ad = clean(ad_city)

        match = norm_target in norm_ad or norm_ad in norm_target

        if not match:
            t_parts = set(norm_target.replace('-', ' ').split())
            a_parts = set(norm_ad.replace('-', ' ').split())
            common = t_parts & a_parts
            if common and len(common) >= min(len(t_parts), len(a_parts)) * 0.5:
                match = True

        if not match:
            return False

    return True

#функция парсинга с карточки (только урла , название товара , цена, город, описание)
def parse_ad_page(driver):
    data = {}
    
    data['url'] = driver.current_url
    
    try:
        title = driver.find_element(By.CSS_SELECTOR, 'h1[data-marker="item-view/title-info"]').text.strip()
        data['title'] = title
    except:
        data['title'] = driver.title.split('—')[0].strip() if '—' in driver.title else driver.title
    
    try:
        price_meta = driver.find_element(By.CSS_SELECTOR, 'meta[itemprop="price"]')
        price_val = price_meta.get_attribute('content')
        data['price'] = f"{int(float(price_val)):,} ₽" if price_val else 'Н/Д'
        data['price_numeric'] = float(price_val) if price_val else None
    except:
        try:
            price_text = driver.find_element(By.CSS_SELECTOR, '[data-marker="item-view/item-price"]').text.strip()
            data['price'] = price_text
            data['price_numeric'] = parse_price(price_text)
        except:
            data['price'] = 'Н/Д'
            data['price_numeric'] = None
    
    try:
        addr = driver.find_element(By.CSS_SELECTOR, '[itemtype="http://schema.org/PostalAddress"]')
        spans = addr.find_elements(By.TAG_NAME, 'span')
        city_val = ''
        for span in spans:
            text = span.text.strip()
            if text and len(text) < 50 and ',' not in text:
                city_val = text
                break
        data['city'] = city_val if city_val else (spans[0].text.strip() if spans else 'Н/Д')
    except:
        data['city'] = 'Н/Д'
    
    try:
        desc = driver.find_element(By.CSS_SELECTOR, '[data-marker="item-view/item-description"]').text.strip()
        data['description'] = desc
    except:
        data['description'] = ''
    
    return data

def set_city(driver, city):
    if not city or city.lower() in ['вся россия','россия','all','']: return True
    try:
        time.sleep(0.02)
        try:
            bn = driver.find_element(By.CSS_SELECTOR, 'button[class*="close"], [data-marker*="close"], button[aria-label*="Закрыть"]')
            if bn.is_displayed(): driver.execute_script("arguments[0].click();", bn); time.sleep(1)
        except: pass
        loc = None
        try:
            wrap = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-marker="search-form/change-location"]')))
            loc = wrap.find_element(By.CSS_SELECTOR, 'a, button, [role="button"]')
        except: pass
        if not loc:
            try: loc = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-marker="header-region"]')))
            except: pass
        if not loc:
            try:
                btns = driver.find_elements(By.CSS_SELECTOR, 'a, button, [role="button"]')
                for b in btns:
                    try:
                        t = b.text.strip()
                        if t and ',' in t and len(t)<60 and '@' not in t and not t.replace(',','').replace('-','').replace(' ','').isdigit():
                            loc = b; break
                    except: continue
            except: pass
        if not loc:
            try: loc = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@data-marker="search-form/change-location"]//a | //header//a[contains(text(), ",")]')))
            except: pass
        if not loc:
            driver.save_screenshot("debug_no_loc.png"); return False
        driver.execute_script("arguments[0].scrollIntoView({block:'center',behavior:'instant'});", loc)
        time.sleep(random.uniform(1,2)); driver.execute_script("arguments[0].click();", loc); time.sleep(random.uniform(2,4))
        inp = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-marker="popup-location/region/search-input"], input[placeholder*="Город"], input[placeholder*="регион"]')))
        time.sleep(0.3)
        inp.send_keys(Keys.CONTROL + "a")
        time.sleep(0.121)
        inp.send_keys(Keys.DELETE)
        time.sleep(0.2)
        for ch in city: inp.send_keys(ch); time.sleep(random.uniform(0.05,0.15))
        time.sleep(random.uniform(0.2,0.3))
        suggs = driver.find_elements(By.CSS_SELECTOR, 'button[data-marker*="custom-option"], div[data-marker*="suggest"], [class*="suggest"] button, [role="option"]')
        for s in suggs:
            if city.lower() in s.text.lower():
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", s)
                time.sleep(random.uniform(0.05,0.15)); driver.execute_script("arguments[0].click();", s); time.sleep(random.uniform(1,2)); break
        try:
            save = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-marker="popup-location/save-button"]')))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", save)
            time.sleep(random.uniform(0.5,1.5)); driver.execute_script("arguments[0].click();", save); time.sleep(random.uniform(3,5))
        except: driver.find_element(By.TAG_NAME,'body').send_keys(Keys.ESCAPE); time.sleep(2)
        return True
    except Exception as e:
        print(f"City err: {e}"); driver.save_screenshot("debug_city_err.png"); return False

def collect_price(q='iPhone', cond='б/у', city='Ростов-на-Дону', price_range=None, pages=3, output=None, price_tol=0.1):
    if output is None:
        pf_str = f"{price_range[0]}-{price_range[1]}" if isinstance(price_range, (tuple,list)) else (str(price_range) if price_range else 'any')
        output = make_filename(q, cond, city, pf_str)
    opts = Options()
    opts.add_argument('--disable-blink-features=AutomationControlled')
    opts.add_experimental_option('excludeSwitches', ['enable-automation'])
    opts.add_argument('--start-maximized')
    opts.add_argument('--disable-popup-blocking')
    driver = webdriver.Chrome(options=opts)
    res = []
    try:
        driver.get('https://www.avito.ru')
        time.sleep(0.03)
        if city and city.lower() not in ['вся россия', 'россия', 'all', '']:
            set_city(driver, city)
        sb = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-marker="search-form/suggest/input"]')))
        sb.clear()
        sb.send_keys(q)
        time.sleep(0.01)
        sb.send_keys(Keys.ENTER)
        time.sleep(0.03)
        for pg in range(pages):
            lh = driver.execute_script("return document.body.scrollHeight")
            for _ in range(3):
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(0.01)
                nh = driver.execute_script("return document.body.scrollHeight")
                if nh == lh: break
                lh = nh
            cards = driver.find_elements(By.CSS_SELECTOR, ITEM_SELECTOR)
            for card in cards:
                try:
                    item_data = get_item_data_from_card(card, driver)
                    if not item_data.get('url') or item_data['url'] in [r['url'] for r in res]:
                        continue
                    
                    price_num = item_data.get('price_numeric')
                    if price_range and price_num:
                        if isinstance(price_range, (tuple, list)) and len(price_range) == 2:
                            if not (price_range[0] <= price_num <= price_range[1]):
                                continue
                        elif isinstance(price_range, (int, float)):
                            if abs(price_num - price_range) > price_range * price_tol:
                                continue
                    
                    driver.execute_script("window.open(arguments[0], '_blank');", item_data['url'])
                    driver.switch_to.window(driver.window_handles[-1])
                    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    time.sleep(1.5)

                    ad_data = parse_ad_page(driver)

                    ad_data = parse_ad_page(driver)
                    item_data.update(ad_data)
                    target_data = item_data

                    if filter_item(target_data, city, price_range, cond, price_tol):
                        res.append(target_data)
                        print(f"Added: {target_data.get('title')}")
                    else:
                        print(f"Filtered: {target_data.get('title')} | city={target_data.get('city')} cond={target_data.get('condition')}")
                    
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    time.sleep(0.01)
                    
                except Exception as e:
                    if len(driver.window_handles) > 1:
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                    continue
            if pg < pages - 1:
                try:
                    nb = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[data-marker="pagination-button/nextPage"]')))
                    driver.execute_script("arguments[0].click();", nb)
                    time.sleep(0.05)
                except: break
        save_to_excel(res, output)
        print(f"Готово! Найдено: {len(res)} | Файл: {output}")
    except Exception as e:
        print(f"Ошибка: {e}")
    finally:
        driver.quit()
        return res

def parse_price(s):
    if not s: return None
    s = re.sub(r'[^\d,.\-]', '', str(s))
    s = s.replace(',', '.')
    try: return float(s)
    except: return None

if __name__ == '__main__':
    collect_price(q='iphone14 pro max' , cond='б/у' , city="Ростов-на-Дону" , price_range=(32_000,35_000) , pages=6 , output=None , price_tol=0.9)

#аргумент функции q - название товара, cond-состояние товара(в нынешней версии полностью игнорируется, чуть позже доработаю) , city-город с которого парсить данные, price_range-диапазон цен, pages - кол-во страниц для парсинга , output-название файла , состояние None , означает что файл будет назваться так как хадаст алгоритм описанный в функции make_filename,price_total-максимальное отклонение в процентах(0.1 = 10%)