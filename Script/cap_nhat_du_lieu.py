# -*- coding: utf-8 -*-
"""
Chép nội dung "Tồn kho.xlsx" vào data.js để website mở trực tiếp (bấm đúp index.html) vẫn thấy dữ liệu.

  python cap_nhat_du_lieu.py           -> tạo data.js một lần
  python cap_nhat_du_lieu.py --watch   -> theo dõi: mỗi khi Excel được lưu thì tự tạo lại data.js,
                                          rồi commit + đẩy Tồn kho.xlsx và data.js lên GitHub (nếu đã nối GitHub)
"""
import base64
import datetime
import io
import json
import os
import subprocess
import sys
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, ".."))
FILE_EXCEL = os.path.join(ROOT, "Tồn kho.xlsx")
FILE_DATA = os.path.join(HERE, "data.js")
KIEM_TRA_GIAY = 3
CHO_TRUOC_KHI_DAY = 10        # giây chờ sau lần lưu cuối rồi mới đẩy lên GitHub


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


def git(*args):
    """Chạy lệnh git trong thư mục gốc (chứa Tồn kho.xlsx); trả về (mã thoát, nội dung in ra)."""
    env = dict(os.environ, GIT_TERMINAL_PROMPT="0")
    try:
        r = subprocess.run(["git", *args], cwd=ROOT, env=env, capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=180)
        return r.returncode, (r.stdout + r.stderr).strip()
    except (OSError, subprocess.TimeoutExpired) as e:
        return 1, str(e)


def co_github():
    ma, ra = git("remote")
    return ma == 0 and "origin" in ra.split()


def day_len_github():
    """Commit Tồn kho.xlsx + data.js rồi đẩy lên GitHub. Trả về True nếu xong (hoặc không có gì để đẩy)."""
    git("add", "--", "Tồn kho.xlsx", "Script/data.js")
    ma, _ = git("diff", "--cached", "--quiet")
    if ma != 0:                                        # có thay đổi -> tạo commit
        ma, ra = git("commit", "-m", "Cập nhật dữ liệu " + time.strftime("%m/%d/%Y %H:%M"))
        if ma != 0:
            print("   ! Không tạo được commit:", ra.splitlines()[-1] if ra else "", flush=True)
            return False
    ma, ra = git("rev-list", "--count", "@{u}..HEAD")
    if ma == 0 and ra.strip() == "0":
        return True                                    # GitHub đã có bản mới nhất
    ma, ra = git("push")
    if ma != 0:                                        # GitHub có thay đổi khác -> lấy về rồi đẩy lại
        git("pull", "--rebase", "--autostash")
        ma, ra = git("push")
    if ma == 0:
        print(time.strftime("%H:%M:%S"), "– đã đẩy lên GitHub, website sẽ cập nhật sau 1–2 phút", flush=True)
        return True
    print("   ! Chưa đẩy được lên GitHub (sẽ thử lại sau 1 phút):", ra.splitlines()[-1] if ra else "", flush=True)
    return False


def theo_doi():
    lan_truoc = None
    github = co_github()
    cho_day = None                                     # thời điểm sẽ đẩy lên GitHub
    if not github:
        print("(Thư mục chưa nối với GitHub – chỉ cập nhật data.js trên máy.)\n", flush=True)
    while True:
        try:
            m = os.path.getmtime(FILE_EXCEL)
            if m != lan_truoc and tao_data_js():
                lan_truoc = m
                print(time.strftime("%H:%M:%S"), "– đã cập nhật data.js theo Tồn kho.xlsx", flush=True)
                if github:
                    cho_day = time.time() + CHO_TRUOC_KHI_DAY   # đợi Excel lưu xong hẳn rồi mới đẩy
        except OSError:
            pass
        if cho_day is not None and time.time() >= cho_day:
            cho_day = None if day_len_github() else time.time() + 60
        time.sleep(KIEM_TRA_GIAY)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    if "--watch" in sys.argv:
        print("Đang theo dõi Tồn kho.xlsx – mỗi lần lưu Excel, website (trên máy và trên GitHub) tự cập nhật.")
        print("Giữ cửa sổ này mở (có thể thu nhỏ). Đóng cửa sổ = tắt tự cập nhật.\n", flush=True)
        theo_doi()
    elif tao_data_js():
        print("Đã cập nhật data.js từ Tồn kho.xlsx")
    else:
        print("Không đọc được Tồn kho.xlsx")
        sys.exit(1)
