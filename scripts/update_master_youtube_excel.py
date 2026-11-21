# -*- coding: utf-8 -*-
import os
import openpyxl

wb_path = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\Ke_Hoach_Phat_Song_YouTube_Series.xlsx"
tmp_path = wb_path.replace(".xlsx", ".tmp.xlsx")

wb = openpyxl.load_workbook(wb_path)

# 1. Update CM2_DICH_LUAT_TIENG_NGUOI
ws2 = wb["CM2_DICH_LUAT_TIENG_NGUOI"]

# CM2-CD10 (Row 8)
ws2.cell(row=8, column=12, value="🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
ws2.cell(row=8, column=13, value='=HYPERLINK("C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\CM2-CD10_Van_don_Switch_B_L_trong.md", "👉 Xem Kịch Bản Chi Tiết")')

# CM2-CD14 (Row 12)
ws2.cell(row=12, column=3, value="Áp Dụng Thông Tư 33 Mới Nhất Về Xuất Xứ Hàng Hóa: 6 Trường Hợp Bị Bác C/O & Bí Quyết Gỡ Rối Tránh Bị Truy Thu Thuế Tiền Tỷ!")
ws2.cell(row=12, column=4, value="Thông tư 33/2023/TT-BTC; Nghị định 08/2015/NĐ-CP; Công văn 22560/CHQ-GSQL")
ws2.cell(row=12, column=7, value="Thông tư 33/2023/TT-BTC; Bác C/O; Xuất xứ hàng hóa; Truy thu thuế; C/O Form E")
ws2.cell(row=12, column=12, value="🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
ws2.cell(row=12, column=13, value='=HYPERLINK("C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\CM2-CD14_Thong_tu_33_Nhung_truong_hop_bi_bac_CO.md", "👉 Xem Kịch Bản Chi Tiết")')

# CM2-CD15 (Row 13)
ws2.cell(row=13, column=12, value="🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
ws2.cell(row=13, column=13, value='=HYPERLINK("C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\CM2-CD15_Phan_biet_rach_roi_Hang_phi.md", "👉 Xem Kịch Bản Chi Tiết")')

# CM2-CD19 (Row 17)
ws2.cell(row=17, column=1, value=13)
ws2.cell(row=17, column=2, value="CM2-CD19")
ws2.cell(row=17, column=3, value="Công Văn 22560 Về Nhập Khẩu Ủy Thác Từ Trung Quốc: 3 Nguyên Tắc Cốt Tử Về Người Khai Hải Quan & Tính Hợp Lệ Của C/O Mẫu E Để Hưởng Thuế ACFTA!")
ws2.cell(row=17, column=4, value="Công văn 22560/CHQ-GSQL; Nghị định 08/2015/NĐ-CP; Thông tư 38/2015/TT-BTC; Thông tư 12/2019/TT-BCT")
ws2.cell(row=17, column=5, value="Dịch Luật Ra Tiếng Người (CM2)")
ws2.cell(row=17, column=6, value="YouTube Dài (8-10p)")
ws2.cell(row=17, column=7, value="Nhập khẩu ủy thác; C/O Mẫu E; Công văn 22560; Thuế ACFTA; Ô số 2 C/O Form E")
ws2.cell(row=17, column=8, value="Chị Thúy")
ws2.cell(row=17, column=9, value="03/11/2026")
ws2.cell(row=17, column=10, value="Chị Thúy chỉ đạo nghiên cứu Công văn 22560/CHQ-GSQL")
ws2.cell(row=17, column=11, value="Chị Thúy đã duyệt")
ws2.cell(row=17, column=12, value="🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
ws2.cell(row=17, column=13, value='=HYPERLINK("C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\CM2-CD19_Cong_van_22560_Nhap_khau_uy_thac_va_CO_Form_E.md", "👉 Xem Kịch Bản Chi Tiết")')

# 2. Update CM4_TAM_TINH_THUC_CHIEN
ws4 = wb["CM4_TAM_TINH_THUC_CHIEN"]
ws4.cell(row=10, column=12, value="🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
ws4.cell(row=10, column=13, value='=HYPERLINK("C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\CM4-CD10_Tien_coc_di_lien_khuc_ruot_XUONG_NGOAI.md", "👉 Xem Kịch Bản Chi Tiết")')

ws4.cell(row=11, column=12, value="🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
ws4.cell(row=11, column=13, value='=HYPERLINK("C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\CM4-CD11_Quan_tri_cong_no_khach_hang.md", "👉 Xem Kịch Bản Chi Tiết")')

ws4.cell(row=13, column=12, value="🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
ws4.cell(row=13, column=13, value='=HYPERLINK("C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\CM4-CD13_Khi_nhan_su_kinh_doanh_nong.md", "👉 Xem Kịch Bản Chi Tiết")')

ws4.cell(row=14, column=12, value="🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
ws4.cell(row=14, column=13, value='=HYPERLINK("C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\CM4-CD14_Dam_phan_Local_Charges_va_Demurrage.md", "👉 Xem Kịch Bản Chi Tiết")')

# 3. Update CM3_LICH_SU_THUONG_MAI
ws3 = wb["CM3_LICH_SU_THUONG_MAI"]
cm3_map = {
    14: "ThuyTalk_CM3-CD10_Ky_An_Kenh_Dao_Panama_1914_Studio.xlsx",
    15: "ThuyTalk_CM3-CD11_Cuoc_Khung_Hoang_Kenh_Dao_Suez_Studio.xlsx",
    16: "ThuyTalk_CM3-CD12_Tuyen_Duong_Sat_Xuyen_Luc_Dia_Studio.xlsx",
    17: "ThuyTalk_CM3-CD13_De_Che_FedEx_Cu_Dat_Cuoc_Studio.xlsx",
    18: "ThuyTalk_CM3-CD14_Thuong_Cang_Hoi_An_The_Ky_Studio.xlsx",
    19: "ThuyTalk_CM3-CD15_Sieu_Cang_Tu_Dong_Hoa_Maasvlakte_Studio.xlsx",
    20: "ThuyTalk_CM3-CD16_Ky_An_Chuan_Hoa_Khoa_Twistlock_Studio.xlsx",
    21: "ThuyTalk_CM3-CD17_Ky_An_Tuyen_Duong_Tra_Ma_Studio.xlsx",
    22: "ThuyTalk_CM3-CD18_Cuoc_Dai_Di_Cu_Chuoi_Cung_Studio.xlsx",
}

for row_idx, fname in cm3_map.items():
    ws3.cell(row=row_idx, column=12, value="🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
    full_path = "C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\" + fname
    ws3.cell(row=row_idx, column=13, value=f'=HYPERLINK("{full_path}", "🎬 Xem Studio & Phân Cảnh Flow")')

# Hyperlinks cho CM3-CD01 đến CM3-CD09
cm3_prev = {
    5: "CM3-CD01_E01_Chuyen_Chiec_Thung_Container_Tap_1.md",
    6: "CM3-CD02_Tie_Tra_Boston_1773.md",
    7: "ThuyTalk_CM3-CD03_Tap_1_Thuong_diem_Phuong_Tay_Studio.xlsx",
    8: "CM3-CD04_Phien_Da_Babylon_Bo_Luat_Hammurabi.md",
    9: "CM3-CD05_De_Che_Dong_An_Ha_Lan_VOC.md",
    10: "ThuyTalk_CM3-CD06_Con_duong_To_lua_co_dai_Studio.xlsx",
    11: "ThuyTalk_CM3-CD07_Chien_tranh_Thuoc_phien_1839_1842_Studio.xlsx",
    12: "ThuyTalk_CM3-CD08_Kenh_dao_Suez_1869_Vet_cat_Studio.xlsx",
    13: "ThuyTalk_CM3-CD09_Vu_trom_buc_hoa_Mona_Lisa_Studio.xlsx",
}
for row_idx, fname in cm3_prev.items():
    full_path = "C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\" + fname
    label = "🎬 Xem Studio & Phân Cảnh Flow" if fname.endswith(".xlsx") else "👉 Xem Kịch Bản Chi Tiết"
    ws3.cell(row=row_idx, column=13, value=f'=HYPERLINK("{full_path}", "{label}")')

# Atomic save (RULE 10 - Zero backup)
wb.save(tmp_path)
wb.close()
os.replace(tmp_path, wb_path)
print("✅ Đã cập nhật thành công và an toàn tuyệt đối (Atomic Save) trên Ke_Hoach_Phat_Song_YouTube_Series.xlsx!")
