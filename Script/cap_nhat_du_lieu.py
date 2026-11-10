# -*- coding: utf-8 -*-
"""
Chép nội dung "Tồn kho.xlsx" vào data.js để website mở trực tiếp (bấm đúp index.html) vẫn thấy dữ liệu.

  python cap_nhat_du_lieu.py           -> tạo data.js một lần
  pyw    cap_nhat_du_lieu.py --watch   -> chạy ngầm: mỗi khi Excel được lưu thì tự tạo lại data.js
"""
import base64
import datetime
import io
import json
import os
import sys
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
FILE_EXCEL = os.path.join(HERE, "..", "Tồn kho.xlsx")
FILE_DATA = os.path.join(HERE, "data.js")
KIEM_TRA_GIAY = 3


def doc_excel():
    """Đọc file Excel; bỏ qua nếu Excel đang lưu dở (file chưa hoàn chỉnh)."""
    for _ in range(5):
        try:
            with open(FILE_EXCEL, "rb") as f:
                data = f.read()
            zipfile.ZipFile(io.BytesIO(data)).testzip()
            return data
        except (OSError, zipfile.BadZipFile):
            time.sleep(1)
    return None


def tao_data_js():
    data = doc_excel()
    if data is None:
        return False
    mtime = os.path.getmtime(FILE_EXCEL)
    noi_dung = {
        "modified": datetime.datetime.fromtimestamp(mtime).isoformat(timespec="seconds"),
        "b64": base64.b64encode(data).decode("ascii"),
    }
    tam = FILE_DATA + ".tmp"
    with open(tam, "w", encoding="utf-8") as f:
        f.write("/* Tạo tự động từ Tồn kho.xlsx bởi cap_nhat_du_lieu.py – không sửa tay */\n")
        f.write("window.TON_KHO_XLSX = " + json.dumps(noi_dung) + ";\n")
    os.replace(tam, FILE_DATA)
    return True


def theo_doi():
    lan_truoc = None
    while True:
        try:
            m = os.path.getmtime(FILE_EXCEL)
            if m != lan_truoc and tao_data_js():
                lan_truoc = m
                print(time.strftime("%H:%M:%S"), "– đã cập nhật data.js theo Tồn kho.xlsx", flush=True)
        except OSError:
            pass
        time.sleep(KIEM_TRA_GIAY)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    if "--watch" in sys.argv:
        print("Đang theo dõi Tồn kho.xlsx – mỗi lần lưu Excel, website tự cập nhật.")
        print("Giữ cửa sổ này mở (có thể thu nhỏ). Đóng cửa sổ = tắt tự cập nhật.\n", flush=True)
        theo_doi()
    elif tao_data_js():
        print("Đã cập nhật data.js từ Tồn kho.xlsx")
    else:
        print("Không đọc được Tồn kho.xlsx")
        sys.exit(1)
