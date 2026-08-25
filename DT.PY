INPUT_FOLDER = r"C:\Users\Lenovo\Desktop\KLTN\DATA\DATA TỔNG HỢP\BCTN"
import os
import re
import math
import glob
import pandas as pd
import pymupdf as fitz
OUTPUT_FILE = "digi_score_output.xlsx"
KEYWORD_GROUPS = {
    "Cong_nghe_nen_tang": [
        "trí tuệ nhân tạo", "học máy", "dữ liệu lớn", "điện toán đám mây",
        "blockchain", "internet vạn vật", "iot", "công nghệ 4.0",
        "cách mạng công nghiệp 4.0", "big data", "cloud computing",
    ],
    "Ung_dung_so_hoa": [
        "chuyển đổi số", "số hóa", "ngân hàng số", "ngân hàng điện tử",
        "thanh toán điện tử", "thanh toán không dùng tiền mặt",
        "ứng dụng di động", "app di động", "hóa đơn điện tử",
        "định danh điện tử", "chữ ký số",
    ],
    "Mo_hinh_kinh_doanh_so": [
        "thương mại điện tử", "nền tảng trực tuyến", "kinh doanh trực tuyến",
        "hệ sinh thái số", "kênh số", "giao dịch trực tuyến",
    ],
}
 
# Gộp tất cả từ khóa thành 1 danh sách phẳng để đếm
ALL_KEYWORDS = [kw for group in KEYWORD_GROUPS.values() for kw in group]
def extract_text_from_pdf(pdf_path: str) -> str:
    """Trích xuất toàn bộ text từ 1 file PDF (dạng text, không phải scan ảnh)."""
    text = ""
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            text += page.get_text()
        doc.close()
    except Exception as e:
        print(f"  [LỖI] Không đọc được {pdf_path}: {e}")
    return text
 
 
def normalize_text(text: str) -> str:
    """Chuẩn hóa văn bản: lowercase, gộp khoảng trắng thừa."""
    text = text.lower()
    text = re.sub(r"\s+", " ", text)
    return text
 
 
def count_keywords(text: str, keywords: list) -> dict:
    """Đếm số lần xuất hiện của từng từ khóa trong văn bản (đếm chồng lấp cho phép)."""
    counts = {}
    for kw in keywords:
        pattern = re.escape(kw.lower())
        counts[kw] = len(re.findall(pattern, text))
    return counts
 
 
def count_total_words(text: str) -> int:
    """Đếm tổng số từ trong văn bản (ước lượng, tách theo khoảng trắng)."""
    return len(text.split())
 
 
def parse_filename(filename: str):
    """
    Tách Mã CK và Năm từ tên file, giả định định dạng: MACK_NAM.pdf
    VD: FPT_2020.pdf -> ("FPT", 2020)
    Nếu tên file không đúng định dạng, trả về (tên file, None) và bạn
    cần bổ sung thủ công trong file Excel kết quả.
    """
    base = os.path.splitext(os.path.basename(filename))[0]
    match = re.match(r"([A-Za-z0-9]+)[_\-](\d{4})", base)
    if match:
        return match.group(1).upper(), int(match.group(2))
    return base, None
 
 
def is_pdf_likely_scanned(text: str) -> bool:
    """Cảnh báo nếu PDF có thể là bản scan ảnh (text trích ra quá ít)."""
    return len(text.strip()) < 200
def main():
    if not os.path.isdir(INPUT_FOLDER):
        print(f"[LỖI] Không tìm thấy thư mục '{INPUT_FOLDER}'.")
        print(f"      Hãy tạo thư mục '{INPUT_FOLDER}' và đặt các file PDF BCTN vào đó.")
        return
 
    pdf_files = glob.glob(os.path.join(INPUT_FOLDER, "*.pdf"))
    if not pdf_files:
        print(f"[LỖI] Không có file .pdf nào trong thư mục '{INPUT_FOLDER}'.")
        return
 
    print(f"Tìm thấy {len(pdf_files)} file PDF. Bắt đầu xử lý...\n")
 
    detail_rows = []   # chi tiết từng từ khóa
    summary_rows = []  # tổng hợp theo công ty-năm
 
    for pdf_path in sorted(pdf_files):
        filename = os.path.basename(pdf_path)
        ma_ck, nam = parse_filename(filename)
        print(f"-> Đang xử lý: {filename}  (Mã CK: {ma_ck}, Năm: {nam})")
 
        raw_text = extract_text_from_pdf(pdf_path)
 
        if is_pdf_likely_scanned(raw_text):
            print(f"   [CẢNH BÁO] File có vẻ là bản scan ảnh, text trích ra rất ít.")
            print(f"              Cần OCR bằng pytesseract để đọc được nội dung.")
 
        norm_text = normalize_text(raw_text)
        total_words = count_total_words(norm_text)
        kw_counts = count_keywords(norm_text, ALL_KEYWORDS)
        total_kw_count = sum(kw_counts.values())
 
        # Ghi chi tiết từng từ khóa
        for group_name, kws in KEYWORD_GROUPS.items():
            for kw in kws:
                detail_rows.append({
                    "Ma_CK": ma_ck,
                    "Nam": nam,
                    "File": filename,
                    "Nhom_tu_khoa": group_name,
                    "Tu_khoa": kw,
                    "So_lan_xuat_hien": kw_counts[kw],
                })
 
        # Tính biến DIGI theo các cách phổ biến trong nghiên cứu
        digi_ln = math.log(total_kw_count + 1)  # ln(tổng tần suất + 1)
        digi_ratio = (total_kw_count / total_words) if total_words > 0 else 0
        digi_binary = 1 if total_kw_count > 0 else 0
 
        summary_rows.append({
            "Ma_CK": ma_ck,
            "Nam": nam,
            "File": filename,
            "Tong_so_tu_trong_BCTN": total_words,
            "Tong_tan_suat_tu_khoa": total_kw_count,
            "DIGI_ln": round(digi_ln, 4),
            "DIGI_ty_le(%)": round(digi_ratio * 100, 4),
            "DIGI_nhi_phan": digi_binary,
        })
 
        print(f"   Tổng từ khóa tìm thấy: {total_kw_count} | Tổng số từ: {total_words} "
              f"| DIGI_ln = {round(digi_ln, 4)}\n")
 
    # Xuất ra Excel
    df_detail = pd.DataFrame(detail_rows)
    df_summary = pd.DataFrame(summary_rows)
 
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="Tong_hop", index=False)
        df_detail.to_excel(writer, sheet_name="Chi_tiet", index=False)
 
    print(f"\n✅ HOÀN TẤT! Kết quả đã lưu tại: {OUTPUT_FILE}")
    print(f"   - Sheet 'Tong_hop': dùng để đưa vào Stata/R/Python chạy hồi quy")
    print(f"   - Sheet 'Chi_tiet': dùng để kiểm tra lại từng từ khóa nếu cần")
 
 
if __name__ == "__main__":
    main()
print("Đường dẫn đang kiểm tra:", INPUT_FOLDER)
print("Tồn tại không:", os.path.isdir(INPUT_FOLDER))
