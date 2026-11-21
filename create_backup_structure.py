import os
import shutil

ROOT_DIR = r"c:\Users\Admin\Desktop\tester"
BACKUP_ROOT = os.path.join(ROOT_DIR, "dan_chung_bao_cao")

def safe_copy(src, dst):
    if not os.path.exists(src):
        print(f"Warning: source does not exist: {src}")
        return False
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)
    print(f"Copied: {os.path.basename(src)} -> {dst}")
    return True

def safe_copy_tree(src, dst):
    if not os.path.exists(src):
        print(f"Warning: source dir does not exist: {src}")
        return False
    os.makedirs(dst, exist_ok=True)
    for root, dirs, files in os.walk(src):
        rel_path = os.path.relpath(root, src)
        dest_dir = os.path.join(dst, rel_path) if rel_path != '.' else dst
        os.makedirs(dest_dir, exist_ok=True)
        for f in files:
            s_file = os.path.join(root, f)
            d_file = os.path.join(dest_dir, f)
            shutil.copy2(s_file, d_file)
    print(f"Copied directory: {src} -> {dst}")
    return True

def run_backup():
    print("=== STARTING BACKUP TO:", BACKUP_ROOT, "===")
    os.makedirs(BACKUP_ROOT, exist_ok=True)

    # 1. Mục 1
    dir_m1 = os.path.join(BACKUP_ROOT, "01_Muc_1_UAT_Inxify_Va_Bao_Ve_Du_Lieu")
    safe_copy(os.path.join(ROOT_DIR, "taikhoan.md"), os.path.join(dir_m1, "M1_Danh_sach_09_tai_khoan_UAT_Inxify.md"))

    # 2. Mục 2
    dir_m2 = os.path.join(BACKUP_ROOT, "02_Muc_2_Doi_Chieu_Truong_2_Man_Hinh_Inxify")
    safe_copy(os.path.join(ROOT_DIR, "Mau_Excel_QC_Inxify_DoiChieu_va_TheoDoiPhieu_Hoan_Chinh.xlsx"), 
              os.path.join(dir_m2, "M2_Bang_doi_chieu_20_truong_DatPhong_va_18_truong_KhachHang.xlsx"))

    # 3. Mục 3 & 4
    dir_m3_4 = os.path.join(BACKUP_ROOT, "03_Muc_3_Va_4_Kiem_Soat_Va_Theo_Doi_Phieu_Yeu_Cau")
    safe_copy(os.path.join(ROOT_DIR, "Mau_Excel_QC_Inxify_DoiChieu_va_TheoDoiPhieu_Hoan_Chinh.xlsx"), 
              os.path.join(dir_m3_4, "M3_M4_So_theo_doi_04_phieu_thuc_te_va_Bao_cao_tuan.xlsx"))

    # 4. Mục 5
    dir_m5 = os.path.join(BACKUP_ROOT, "04_Muc_5_Kich_Ban_Quet_QR_CCCD")
    safe_copy(os.path.join(ROOT_DIR, "Tai_lieu_Kich_ban_Kiem_thu_Nghiem_thu_Quet_QR_CCCD_Inxify_Muc_5.docx"), 
              os.path.join(dir_m5, "M5_Tai_lieu_Kich_ban_nghiem_thu_Quet_QR_CCCD_Inxify.docx"))

    # 5. Mục 6
    dir_m6 = os.path.join(BACKUP_ROOT, "05_Muc_6_Canh_Bao_Kiem_Tra_Luu_Du_Lieu")
    safe_copy(os.path.join(ROOT_DIR, "Bao_cao_Kiem_thu_Canh_bao_Luu_Du_lieu_Inxify_Muc_6_Co_Anh_Dan_Chung.docx"), 
              os.path.join(dir_m6, "M6_Bao_cao_kiem_thu_Canh_bao_luu_du_lieu_Co_anh_dan_chung.docx"))
    safe_copy_tree(os.path.join(ROOT_DIR, "imgtest_muc6"), 
                   os.path.join(dir_m6, "Anh_dan_chung_UAT_Muc_6"))

    # 6. Mục 7
    dir_m7 = os.path.join(BACKUP_ROOT, "06_Muc_7_Nhat_Ky_Thao_Tac_Audit_Log_Uu_Tien_1")
    safe_copy(os.path.join(ROOT_DIR, "Bao_cao_Trinh_Ban_Giam_Doc_Ket_Qua_Nghiem_Thu_Audit_Log_Inxify_Muc_7.docx"), 
              os.path.join(dir_m7, "M7_Bao_cao_Trinh_Ban_Giam_Doc_Ket_qua_nghiem_thu_Audit_Log.docx"))
    safe_copy(os.path.join(ROOT_DIR, "Template_TestCase_Audit_Log_Inxify_Thang_11_2026_Hoan_chinh.xlsx"), 
              os.path.join(dir_m7, "M7_Ma_tran_35_TestCase_Audit_Log_Inxify_Hoan_chinh.xlsx"))
    safe_copy(os.path.join(ROOT_DIR, "ĐẶC TẢ YÊU CẦU CHỨC NĂNG.docx"), 
              os.path.join(dir_m7, "M7_Dac_ta_yeu_cau_chuc_nang_Audit_Log_BA_Chau_Tuan.docx"))
    safe_copy(os.path.join(ROOT_DIR, "Bao_cao_Kiem_thu_Audit_Log_Inxify_Thang_11_2026_Co_Anh_Dan_Chung.docx"), 
              os.path.join(dir_m7, "M7_Bao_cao_Kiem_thu_Audit_Log_Inxify_Chi_tiet.docx"))
    safe_copy_tree(os.path.join(ROOT_DIR, "imgtest"), 
                   os.path.join(dir_m7, "Anh_dan_chung_UAT_Audit_Log"))
    safe_copy_tree(os.path.join(ROOT_DIR, "json_devtools_uat"), 
                   os.path.join(dir_m7, "Du_lieu_DevTools_Payload_JSON"))

    # 7. Mục 8 & 9
    dir_m8_9 = os.path.join(BACKUP_ROOT, "07_Muc_8_Va_9_Nguon_Kenh_OTA_Va_Phong_Chua_San_Sang")
    safe_copy(os.path.join(ROOT_DIR, "Van_ban_Tong_hop_Yeu_cau_Ky_thuat_Nghiem_thu_Inxify_Muc_8_va_Muc_9_Co_Anh_Dan_Chung.docx"), 
              os.path.join(dir_m8_9, "M8_M9_Van_ban_Tong_hop_Yeu_cau_ky_thuat_nghiem_thu_Inxify.docx"))
    safe_copy_tree(os.path.join(ROOT_DIR, "imgtest_muc8_9"), 
                   os.path.join(dir_m8_9, "Anh_dan_chung_UAT_Muc_8_va_9"))

    # 8. Báo cáo PCNTT & Hồ sơ thử việc
    dir_pcntt = os.path.join(BACKUP_ROOT, "08_Bao_Cao_Tong_Hop_PCNTT_Va_Ho_So_Thu_Viec")
    safe_copy(os.path.join(ROOT_DIR, "PCNTT_MauBC.docx"), 
              os.path.join(dir_pcntt, "PCNTT_MauBC_Da_dien_chuan.docx"))
    safe_copy(os.path.join(ROOT_DIR, "Bao_cao_PCNTT_Tong_hop_Tien_do_Xu_ly_Ton_dong_va_Phuong_an_Giai_quyet.docx"), 
              os.path.join(dir_pcntt, "Bao_cao_PCNTT_Tong_hop_Tien_do_Xu_ly_Ton_dong_va_Phuong_an.docx"))
    safe_copy(os.path.join(ROOT_DIR, "Tong_Hop_Tien_Do_Xu_Ly_Ton_Dong_Va_Phuong_An_Giai_Quyet_Tester.md"), 
              os.path.join(dir_pcntt, "Tong_Hop_Tien_Do_Xu_Ly_Ton_Dong_Va_Phuong_An_Giai_Quyet_Tester.md"))
    safe_copy(os.path.join(ROOT_DIR, "Bao_cao_so_bo_sau_7_ngay_thu_viec_A25_Hoang_Minh_Quan_Bo_sung_v5.docx"), 
              os.path.join(dir_pcntt, "Bao_cao_so_bo_sau_7_ngay_thu_viec_A25_Hoang_Minh_Quan_v5.docx"))
    safe_copy(os.path.join(ROOT_DIR, "hronline.md"), 
              os.path.join(dir_pcntt, "Chi_dao_09_dau_muc_cong_viec_hronline.md"))

    # 9. Chuẩn hóa M365
    dir_m365 = os.path.join(BACKUP_ROOT, "09_Chuan_Hoa_He_Thong_So_Va_Dao_Tao_M365")
    safe_copy(os.path.join(ROOT_DIR, "Bao_cao_Phuong_an_Chuan_hoa_Cong_cu_Microsoft_365_A25_Hotel.docx"), 
              os.path.join(dir_m365, "Bao_cao_Phuong_an_Chuan_hoa_Cong_cu_Microsoft_365_A25_Hotel.docx"))
    safe_copy(os.path.join(ROOT_DIR, "Video_Huong_Dan_Chuan_Hoa_Microsoft_365_A25_Hotel.mp4"), 
              os.path.join(dir_m365, "Video_Huong_Dan_Chuan_Hoa_Microsoft_365_A25_Hotel.mp4"))
    safe_copy(os.path.join(ROOT_DIR, "Trung_Tam_Video_Huong_Dan_M365_A25_Hotel.html"), 
              os.path.join(dir_m365, "Trung_Tam_Video_Huong_Dan_M365_A25_Hotel.html"))
    safe_copy(os.path.join(ROOT_DIR, "ý kiến nhân viên microsoftteam.md"), 
              os.path.join(dir_m365, "Y_kien_nhan_vien_microsoftteam.md"))
    safe_copy_tree(os.path.join(ROOT_DIR, "img_m365_guide"), 
                   os.path.join(dir_m365, "Infographics_M365"))
    safe_copy_tree(os.path.join(ROOT_DIR, "img_real_m365"), 
                   os.path.join(dir_m365, "Anh_xac_thuc_Microsoft_HD"))

    # Create README_MUC_LUC_DAN_CHUNG.md
    readme_path = os.path.join(BACKUP_ROOT, "README_MUC_LUC_DAN_CHUNG.md")
    readme_content = """# MỤC LỤC HỆ THỐNG DẪN CHỨNG & BÁO CÁO DỰ ÁN INXIFY PMS & HỆ THỐNG SỐ A25 HOTEL
**Vị trí thực hiện:** Hoàng Minh Quân – Chuyên viên kiểm thử phần mềm (Tester)  
**Đơn vị:** Ban Phát triển và Quản lý phần mềm – Phòng Công nghệ thông tin  
**Thời điểm lưu trữ:** Tháng 11/2026  

---

Thư mục này được tổ chức thành các nhánh thư mục con độc lập theo đúng 09 đầu mục công việc tại văn bản chỉ đạo `hronline.md` và các nhiệm vụ trọng tâm:

### 📂 DANH MỤC THƯ MỤC CON & ĐẦU MỤC CÔNG VIỆC:

1. **`01_Muc_1_UAT_Inxify_Va_Bao_Ve_Du_Lieu/`**:
   - `M1_Danh_sach_09_tai_khoan_UAT_Inxify.md`: Danh sách 09 tài khoản phân quyền nghiệp vụ độc lập được bàn giao ngày 19/11/2026.

2. **`02_Muc_2_Doi_Chieu_Truong_2_Man_Hinh_Inxify/`**:
   - `M2_Bang_doi_chieu_20_truong_DatPhong_va_18_truong_KhachHang.xlsx`: Bảng đối chiếu chi tiết 20 trường Đặt phòng PMS và 18 trường Khách hàng CRM qua khảo sát web/API.

3. **`03_Muc_3_Va_4_Kiem_Soat_Va_Theo_Doi_Phieu_Yeu_Cau/`**:
   - `M3_M4_So_theo_doi_04_phieu_thuc_te_va_Bao_cao_tuan.xlsx`: Sổ theo dõi 04 phiếu thực tế (THEO_DOI_PHIEU) và Báo cáo tuần tự động (BAO_CAO_TUAN) kèm biên bản trả lại phiếu.

4. **`04_Muc_5_Kich_Ban_Quet_QR_CCCD/`**:
   - `M5_Tai_lieu_Kich_ban_nghiem_thu_Quet_QR_CCCD_Inxify.docx`: 06 Kịch bản kiểm thử nghiệm thu Quét mã QR CCCD và bộ Mock Data 7 trường chuẩn BCA; phương án mượn máy quét 2D lễ tân.

5. **`05_Muc_6_Canh_Bao_Kiem_Tra_Luu_Du_Lieu/`**:
   - `M6_Bao_cao_kiem_thu_Canh_bao_luu_du_lieu_Co_anh_dan_chung.docx`: Báo cáo kiểm thử 10 kịch bản, phát hiện 03 lỗi nghiệp vụ.
   - `Anh_dan_chung_UAT_Muc_6/`: Thư mục chứa 12 hình ảnh chụp dẫn chứng thực tế từ UAT (M6_01 đến M6_06).

6. **`06_Muc_7_Nhat_Ky_Thao_Tac_Audit_Log_Uu_Tien_1/`**:
   - `M7_Bao_cao_Trinh_Ban_Giam_Doc_Ket_qua_nghiem_thu_Audit_Log.docx`: Báo cáo chính thức trình Ban Giám đốc.
   - `M7_Ma_tran_35_TestCase_Audit_Log_Inxify_Hoan_chinh.xlsx`: Ma trận 35 Test Cases cho 05 nghiệp vụ, phát hiện 08 lỗi ghi thiếu trường.
   - `M7_Dac_ta_yeu_cau_chuc_nang_Audit_Log_BA_Chau_Tuan.docx`: Bản Đặc tả BA v0.3.
   - `M7_Bao_cao_Kiem_thu_Audit_Log_Inxify_Chi_tiet.docx`: Báo cáo kỹ thuật chi tiết.
   - `Anh_dan_chung_UAT_Audit_Log/`: Thư mục chứa 13 hình ảnh giao diện UAT.
   - `Du_lieu_DevTools_Payload_JSON/`: Thư mục chứa 07 file JSON chụp payload DevTools thực tế.

7. **`07_Muc_8_Va_9_Nguon_Kenh_OTA_Va_Phong_Chua_San_Sang/`**:
   - `M8_M9_Van_ban_Tong_hop_Yeu_cau_ky_thuat_nghiem_thu_Inxify.docx`: Văn bản kỹ thuật 05 tiêu chuẩn (REQ-M8-01 -> REQ-M9-02) và 08 Test Cases.
   - `Anh_dan_chung_UAT_Muc_8_va_9/`: Thư mục chứa 08 hình ảnh dẫn chứng UAT (M8_01 -> M8_04, M9_01 -> M9_04).

8. **`08_Bao_Cao_Tong_Hop_PCNTT_Va_Ho_So_Thu_Viec/`**:
   - `PCNTT_MauBC_Da_dien_chuan.docx`: Báo cáo điền theo đúng mẫu quy chuẩn PCNTT.
   - `Bao_cao_PCNTT_Tong_hop_Tien_do_Xu_ly_Ton_dong_va_Phuong_an.docx`: Bản Word báo cáo độc lập hoàn chỉnh.
   - `Tong_Hop_Tien_Do_Xu_Ly_Ton_Dong_Va_Phuong_An_Giai_Quyet_Tester.md`: Bản Markdown tổng hợp toàn diện.
   - `Bao_cao_so_bo_sau_7_ngay_thu_viec_A25_Hoang_Minh_Quan_v5.docx`: Báo cáo sơ bộ 07 ngày.
   - `Chi_dao_09_dau_muc_cong_viec_hronline.md`: Văn bản gốc chỉ đạo 09 đầu mục.

9. **`09_Chuan_Hoa_He_Thong_So_Va_Dao_Tao_M365/`**:
   - `Bao_cao_Phuong_an_Chuan_hoa_Cong_cu_Microsoft_365_A25_Hotel.docx`: Tờ trình & Báo cáo M365 (4.03 MB).
   - `Video_Huong_Dan_Chuan_Hoa_Microsoft_365_A25_Hotel.mp4`: Video thực hành HD 720p (02:05) giọng đọc Neural Microsoft.
   - `Trung_Tam_Video_Huong_Dan_M365_A25_Hotel.html`: Trình duyệt video HTML5 tương tác.
   - `Infographics_M365/` & `Anh_xac_thuc_Microsoft_HD/`: Toàn bộ ảnh đồ họa và hình ảnh giao diện xác thực Microsoft.

---
*Ghi chú: Toàn bộ file gốc trong thư mục cha `c:\\Users\\Admin\\Desktop\\tester` vẫn được bảo toàn nguyên vẹn 100%.*
"""
    with open(readme_path, "w", encoding="utf-8") as fp:
        fp.write(readme_content)
    print("SUCCESS: Created index file at:", readme_path)
    print("=== BACKUP COMPLETED 100% ===")

if __name__ == "__main__":
    run_backup()
