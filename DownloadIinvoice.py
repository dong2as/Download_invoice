import os
import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

DOWNLOAD_DIR = os.getcwd()

def mo_trinh_duyet():
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)

    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    })

    driver = webdriver.Chrome(options=chrome_options)
    driver.get("https://www.meinvoice.vn/tra-cuu")
    return driver

def nhap_ma_tra_cuu(driver, ma_tra_cuu):
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "txtCode")))
    o_nhap_ma = driver.find_element(By.ID, "txtCode")
    o_nhap_ma.clear()
    o_nhap_ma.send_keys(ma_tra_cuu)

def bam_nut_tra_cuu(driver):
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnSearchInvoice")))
    driver.find_element(By.ID, "btnSearchInvoice").click()
    time.sleep(3)

def tai_file_pdf(driver):
    try:
        nut_tai = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "download-invoice"))
        )
        nut_tai.click()
        print(" Đã bấm tải hóa đơn PDF...")
        time.sleep(5)
    except Exception as e:
        print(" Không tìm thấy nút tải hóa đơn:", e)

def tra_cuu_hoa_don(ma_tra_cuu):
    print(f" Đang tra cứu mã: {ma_tra_cuu}")
    driver = mo_trinh_duyet()
    try:
        nhap_ma_tra_cuu(driver, ma_tra_cuu)
        bam_nut_tra_cuu(driver)
        tai_file_pdf(driver)
        print(f" Đã xử lý mã: {ma_tra_cuu}")
    except Exception as e:
        print(f" Lỗi với mã {ma_tra_cuu}: {e}")
    finally:
        driver.quit()

def doc_ma_tu_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    return [str(row[0].value).strip() for row in sheet.iter_rows(min_row=2) if row[0].value]

if __name__ == "__main__":
    danh_sach_ma = doc_ma_tu_excel("ma_tra_cuu.xlsx")
    for ma in danh_sach_ma:
        tra_cuu_hoa_don(ma)
