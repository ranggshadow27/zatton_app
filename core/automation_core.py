import time
import base64
from io import BytesIO
from PIL import Image
import pytesseract
import os
import pandas as pd
import pyperclip
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchWindowException, ElementNotInteractableException
from tkinter import messagebox
from datetime import datetime

def initialize_driver():
    options = Options()
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-notifications')
    options.add_argument('--disable-infobars')
    options.add_experimental_option('prefs', {
        'credentials_enable_service': False,
        'profile.password_manager_enabled': False,
        'profile.password_manager_leak_detection': False
    })
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def is_driver_alive(driver):
    """Cek apakah webdriver masih hidup"""
    if driver is None:
        return False
    try:
        _ = driver.current_url
        return True
    except:
        return False

def handle_login(driver, add_log):
    try:
        add_log("> Check Login Page")
        username_field = driver.find_element(By.XPATH, '//*[@id="input-18"]')
        username_field.send_keys("mahaga_pratama")
        password_field = driver.find_element(By.XPATH, '//*[@id="input-21"]')
        password_field.send_keys("mahaga_pratama_pass")
        login_button = driver.find_element(By.XPATH, '//*[@id="app"]/div/main/div/div/div/div/div/div/div/div/div/div/form/div/button')

        # === REQUEST LU: Messagebox captcha sebelum klik login ===
        messagebox.showinfo("Captcha Verification", "Silahkan selesaikan verifikasi captcha")
        # ========================================================

        login_button.click()
        add_log("> Login selesai, menunggu URL ditemukan..")
        WebDriverWait(driver, 30).until(lambda d: "host-finder" in d.current_url)
        add_log("> URL ditemukan refresh ke URL yang diinginkan...")
    except Exception as e:
        raise Exception(f"Login gagal: {e}")

def set_date_input(driver, xpath, date_value, add_log):
    try:
        date_input = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        add_log(f"> Input ditemukan, mencoba mengisi tanggal...")

        driver.execute_script("arguments[0].removeAttribute('readonly')", date_input)

        pyperclip.copy(date_value)
        date_input.click()
        time.sleep(.1)
        date_input.send_keys(Keys.CONTROL + 'a')
        time.sleep(.1)
        date_input.send_keys(Keys.CONTROL + 'v')
        time.sleep(.25)
        date_input.send_keys(Keys.ESCAPE)

        value = driver.execute_script("return arguments[0].value", date_input)
        if value != date_value:
            add_log(f"> Gagal mengisi value dengan {date_value}, mencoba dengan send_keys...")
            date_input.clear()
            date_input.send_keys(date_value)

        add_log(f"> Berhasil mengisi tanggal {date_value} pada date input")
    except Exception as e:
        raise Exception(f"Error mengisi date input: {e}")

def check_canvas_text(driver, canvas, add_log):
    try:
        canvas_base64 = driver.execute_script("return arguments[0].toDataURL('image/png').substring(21);", canvas)
        canvas_data = base64.b64decode(canvas_base64)
        image = Image.open(BytesIO(canvas_data))

        text = pytesseract.image_to_string(image)
        required_texts = ["Volume total", "Volume in", "Volume out", "Speed in", "Speed out", "Speed total"]
        missing_texts = [t for t in required_texts if t not in text]

        if not missing_texts:
            add_log(f"> Semua teks ditemukan di canvas: {required_texts} ✅")
            return True, image
        else:
            add_log(f"> Teks berikut belum ditemukan di canvas: {missing_texts} ❌")
            return False, None
    except Exception as e:
        raise Exception(f"Error saat OCR canvas: {e}")

def process_site_tree(driver, url, site_id, save_folder, start_date, end_date, add_log):
    status = "Done Capture"
    try:
        add_log(f"> Current working directory: {os.getcwd()}")

        try:
            driver.get(url)
        except NoSuchWindowException:
            add_log(f"> No such window for site {site_id}, reinitializing driver...")
            driver = initialize_driver()
            driver.get(url)
        time.sleep(1)

        current_url = driver.current_url
        if "auth/login" in current_url:
            add_log("> Halaman login terdeteksi, melakukan login...")
            handle_login(driver, add_log)
            driver.get(url)
            time.sleep(1)
            current_url = driver.current_url

        if "site-tree" in current_url:
            add_log("> Sudah berada di halaman site-tree, mencari tombol...")

            try:
                WebDriverWait(driver, 45).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="devices"]/div[1]/button'))
                )
                add_log("> Tombol Capture sudah enabled...")

                button = WebDriverWait(driver, 45).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="devices"]/div[2]/div/div/div[2]/div[2]/div/div/div[2]/div/div[4]/div/div[1]/button'))
                )
                add_log("> Tombol Traffic Router ditemukan! Klik tombol...")
                button.click()

                time.sleep(.2)

                set_date_input(driver, '/html/body/div/div/div/div[3]/div/div/div[2]/div/div/div[3]/div[1]/div/div/div[1]/div/div[1]/div[2]/input', start_date, add_log)
                set_date_input(driver, '/html/body/div/div/div/div[3]/div/div/div[2]/div/div/div[3]/div[3]/div/div/div[1]/div/div[1]/div[2]/input', end_date, add_log)

                add_log("> Start date dan end date berhasil diisi, mencari tombol submit...")
                submit_button = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[3]/div/div/div[2]/div/div/div[3]/button'))
                )
                add_log("> Tombol submit ditemukan! Klik tombol...")
                submit_button.click()

                time.sleep(2)

                # dropdown handling (sama persis kode lama)
                add_log("> Mencari elemen dropdown...")
                dropdown = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div/div[3]/div/div/div[2]/div/div[2]/div[3]/div/div/div[1]/div/div/div[1]/div[1]'))
                )
                dropdown.click()
                time.sleep(0.5)

                try:
                    input_element = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div/div[3]/div/div/div[2]/div/div[2]/div[3]/div/div/div[1]/div/div/div[1]/div[1]/input'))
                    )
                    input_element.send_keys(Keys.ARROW_DOWN)
                    time.sleep(0.1)
                    input_element.send_keys(Keys.ARROW_DOWN)
                    time.sleep(0.1)
                    input_element.send_keys(Keys.ENTER)
                    time.sleep(0.2)
                    add_log("> Dropdown berhasil dipilih menggunakan input element!")
                except TimeoutException:
                    active_element = driver.switch_to.active_element
                    active_element.send_keys(Keys.ARROW_DOWN)
                    time.sleep(0.1)
                    active_element.send_keys(Keys.ARROW_DOWN)
                    time.sleep(0.1)
                    active_element.send_keys(Keys.ENTER)
                    time.sleep(0.2)
                    add_log("> Dropdown berhasil dipilih menggunakan active element!")

                # retry table check (sama persis)
                max_retries = 3
                for attempt in range(max_retries):
                    add_log(f"> Attempt {attempt + 1} to find 'Volume Total' in table...")
                    try:
                        table = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div[3]/div/div/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div/div/table'))
                        )
                        rows = table.find_elements(By.TAG_NAME, "tr")
                        volume_total_found = any(
                            cells and cells[0].text.strip() == "Volume Total"
                            for row in rows
                            for cells in [row.find_elements(By.TAG_NAME, "td")]
                        )
                        if volume_total_found:
                            add_log("> Elemen table ditemukan dengan value 'Volume Total'!")
                            break
                        else:
                            time.sleep(2)
                    except Exception as e:
                        add_log(f"> Error checking table on attempt {attempt + 1}: {e}")
                        time.sleep(2)
                else:
                    status = "Table value mismatch: 'Volume Total' not found after retries"
                    add_log(status)
                    return False, status

                # canvas
                canvas = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div[3]/div/div/div[2]/div/div[2]/div[3]/div/div/div/div/div/canvas'))
                )

                timeout = 30
                start_time = time.time()
                while time.time() - start_time < timeout:
                    text_found, image = check_canvas_text(driver, canvas, add_log)
                    if text_found:
                        save_path = os.path.join(save_folder, f"{site_id}.png")
                        image.save(save_path)
                        add_log(f"> Traffic Capture disimpan di: {save_path} ✅")
                        break
                    time.sleep(1)
                else:
                    status = "Timeout: Canvas text not found"
                    return False, status

                return True, status

            except Exception as e:
                status = f"Error saat mencari tombol, input tanggal, table, dropdown, atau canvas: {e}"
                add_log(status)
                return False, status
        else:
            status = f"Unknown URL: {current_url}"
            add_log(status)
            return False, status

    except Exception as e:
        status = f"Error saat proses site {site_id}: {e}"
        add_log(status)
        return False, status

def main(excel_file, save_folder, start_date, end_date, add_log, update_progress, driver):
    try:
        df = pd.read_excel(excel_file)
        if 'site_id' not in df.columns:
            raise ValueError("Kolom 'site_id' tidak ditemukan di Excel")
        add_log(f"> Total site di Excel: {len(df)} Site, Proses Capture")
    except Exception as e:
        add_log(f"> Error membaca Excel: {e}")
        raise

    if 'capture_status' not in df.columns:
        df['capture_status'] = ""

    df_to_process = df[df['capture_status'].isna() | (df['capture_status'] == "")]

    if df_to_process.empty:
        add_log("> Tidak ada site ID dengan capture_status kosong atau null untuk diproses.")
        return

    url_template = "https://manager.zabbix-bakti.io/host/site-tree?site_uniq_id={}"
    loop_count = 0
    display_loop_count = 0
    total_sites = len(df_to_process)

    for index, row in df_to_process.iterrows():
        site_id = row['site_id']
        display_loop_count += 1
        add_log(f"> Memproses site ID: {site_id} (Loop ke-{display_loop_count})")
        update_progress(display_loop_count, total_sites, site_id)
        url = url_template.format(site_id)

        try:
            success, status = process_site_tree(driver, url, site_id, save_folder, start_date, end_date, add_log)
        except NoSuchWindowException:
            add_log(f"> No such window for site {site_id}, reinitializing driver...")
            driver = initialize_driver()
            success, status = process_site_tree(driver, url, site_id, save_folder, start_date, end_date, add_log)

        df.at[index, 'capture_status'] = status

        try:
            df.to_excel(excel_file, index=False)
            add_log(f"> Status untuk {site_id} disimpan ke Excel: {status}")
        except Exception as e:
            add_log(f"> Error menyimpan Excel untuk {site_id}: {e}")

        loop_count += 1
        if loop_count >= 50:
            add_log("> 50 progress completed, reinitializing driver...")
            try:
                driver.quit()
            except:
                pass
            driver = initialize_driver()
            loop_count = 0

    return driver