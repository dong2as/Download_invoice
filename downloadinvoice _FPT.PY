import os
import time
import openpyxl
import xml.etree.ElementTree as ET

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
URL = "https://tracuuhoadon.fpt.com/"

def tao_trinh_duyet():
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)

    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    })
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-gpu")

    # CHỈNH LẠI đường dẫn chromedriver phù hợp máy bạn
    service = Service(executable_path=r"C:\Users\Admin\Documents\Visual Studio 2022\DownloadInvoice\chromedriver.exe")
    return webdriver.Chrome(service=service, options=chrome_options)

def lay_file_xml_moi_nhat():
    files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith(".xml")]
    if not files:
        return None
    files.sort(key=lambda x: os.path.getmtime(os.path.join(DOWNLOAD_DIR, x)), reverse=True)
    return os.path.join(DOWNLOAD_DIR, files[0])

def tra_cuu_hoa_don(driver, url, ma_tra_cuu, ma_so_thue):
    try:
        driver.get(url)
        wait = WebDriverWait(driver, 10)

        # Nhập MST
        input_mst = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//label[contains(text(), 'Mã số thuế')]/following-sibling::input")
        ))
        input_mst.clear()
        input_mst.send_keys(ma_so_thue)

        # Nhập mã tra cứu
        input_tra_cuu = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//label[contains(text(), 'Mã tra cứu')]/following-sibling::input")
        ))
        input_tra_cuu.clear()
        input_tra_cuu.send_keys(ma_tra_cuu)

        # Click nút "Tra cứu"
        btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(text(), 'Tra cứu')]")
        ))
        btn.click()

        # Đợi xử lý và file XML được tải về
        time.sleep(5)

        file_path = lay_file_xml_moi_nhat()
        if file_path:
            return "success", file_path
        else:
            return "fail", None

    except Exception as e:
        print(f"Lỗi tra cứu: {e}")
        return "fail", None

def doc_file_excel(input_path):
    wb = openpyxl.load_workbook(input_path)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        
        mst = row[1]
        ma_tra_cuu = row[2]
        if mst and ma_tra_cuu:
            data.append((str(mst).strip(), str(ma_tra_cuu).strip()))
    return data

def doc_thong_tin_tu_xml(xml_path):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        ns = {'ns': root.tag.split('}')[0].strip('{')}

        thong_tin = {
            'SoHoaDon': root.findtext('.//ns:SoHD', namespaces=ns),
            'DonViBanHang': root.findtext('.//ns:Ten', namespaces=ns),
            'MSTBan': root.findtext('.//ns:MaSoThue', namespaces=ns),
            'DiaChiBan': root.findtext('.//ns:DiaChi', namespaces=ns),
            'SoTaiKhoan': root.findtext('.//ns:TaiKhoan', namespaces=ns),
            'NguoiMua': root.findtext('.//ns:TenNguoiMua', namespaces=ns),
            'DiaChiMua': root.findtext('.//ns:DiaChiMua', namespaces=ns),
            'MSTMua': root.findtext('.//ns:MSTNguoiMua', namespaces=ns)
        }
        return thong_tin
    except Exception as e:
        print(f"Lỗi đọc XML: {e}")
        return None

def ghi_ket_qua_excel(data_list, file_path):
    wb = openpyxl.Workbook()
    sheet = wb.active
    headers = [
        "Mã số thuế", "Mã tra cứu", "Trạng thái",
        "Số HĐ", "Đơn vị bán", "MST bán", "Địa chỉ bán", "STK bán",
        "Người mua", "Địa chỉ mua", "MST mua"
    ]
    sheet.append(headers)
    for row in data_list:
        sheet.append(row)
    wb.save(file_path)

if __name__ == "__main__":
    driver = tao_trinh_duyet()
    danh_sach = doc_file_excel("input.xlsx")
    ket_qua = []

    for mst, ma_tra_cuu in danh_sach:
        print(f" Đang tra cứu: {mst} - {ma_tra_cuu}")
        status, xml_path = tra_cuu_hoa_don(driver, URL, ma_tra_cuu, mst)

        if status == "fail" or not xml_path:
            ket_qua.append([mst, ma_tra_cuu, "fail"] + [""] * 8)
        else:
            thong_tin = doc_thong_tin_tu_xml(xml_path)
            if thong_tin:
                ket_qua.append([
                    mst, ma_tra_cuu, "success",
                    thong_tin['SoHoaDon'], thong_tin['DonViBanHang'], thong_tin['MSTBan'],
                    thong_tin['DiaChiBan'], thong_tin['SoTaiKhoan'], thong_tin['NguoiMua'],
                    thong_tin['DiaChiMua'], thong_tin['MSTMua']
                ])
            else:
                ket_qua.append([mst, ma_tra_cuu, "fail"] + [""] * 8)

    driver.quit()
    ghi_ket_qua_excel(ket_qua, "output.xlsx")
    print("Đã hoàn tất. Kết quả lưu ở: output.xlsx")
