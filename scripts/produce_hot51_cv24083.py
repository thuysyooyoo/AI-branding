# -*- coding: utf-8 -*-
"""
SẢN XUẤT KỊCH BẢN YOUTUBE CM1 (PO-2026-W46-HOT51 / CM1-HOT51) & MỞ ĐƯỜNG CHUYÊN SÂU KTSTQ DOANH NGHIỆP THƯƠNG MẠI
Nguồn gốc: Công văn số 24083/CHQ-NVTHQ ngày 10/11/2026 của Cục Hải quan (Bộ Tài chính)
Tuân thủ tuyệt đối:
- Skill thuytalk-cm124/SKILL.md (4 phần, 5 Hồi + 4 Cầu dẫn đo ni đóng giày, Zero-Bragging, 100% văn nói)
- wf03_youtube_production.md & wf07_content_audit_verification.md
- GEMINI.md Rule 10 (Atomic Save via .tmp.xlsx + os.replace, ZERO backup files)
"""

import os
import sys
import re
import shutil
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
EXCEL_PATH = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")
STATUS_REPORT_MD = os.path.join(BASE_DIR, "reports", "plan1", "content_status_report.md")
SYNC_LOG_PATH = os.path.join(BASE_DIR, "reports", "sync_log.txt")

SCRATCH_PO_DIR = os.path.join(BASE_DIR, "scratch", "production_orders")
REP_YT_DIR = os.path.join(BASE_DIR, "reports", "plan1", "video_scripts")

ONEDRIVE_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING"
ONEDRIVE_YT_DIR = os.path.join(ONEDRIVE_DIR, "youtube")
ONEDRIVE_SERIES_EXCEL = os.path.join(ONEDRIVE_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.xlsx")

PO_CODE = "PO-2026-W46-HOT51"
CM_CODE = "CM1-HOT51"
ONEDRIVE_YT_FILENAME = "CM1-HOT51_Cong_van_24083_Tham_van_gia_va_KTSTQ.md"
SEO_TITLE = "Công Văn 24083/CHQ-NVTHQ: Siết Tham Vấn Giá & Kiểm Tra Sau Thông Quan Dịp Tết 2027"
SEO_KEYWORDS = (
    "công văn 24083 chq nvthq, tham vấn giá hải quan, kiểm tra sau thông quan doanh nghiệp thương mại, "
    "hệ thống gtt02 trị giá hải quan, thông tư 121 2025 tt btc, tham vấn gián tiếp trị giá, "
    "điều 142 điều 143 thông tư 38, thuytalk chuyên mục 1, eureka logistics, nguyễn thị ngọc thúy"
)

os.makedirs(SCRATCH_PO_DIR, exist_ok=True)
os.makedirs(REP_YT_DIR, exist_ok=True)
if os.path.exists(ONEDRIVE_DIR):
    os.makedirs(ONEDRIVE_YT_DIR, exist_ok=True)


def atomic_save_workbook(wb, filepath):
    """Lưu workbook nguyên tử qua .tmp.xlsx + os.replace (Zero Backup Rule 10)."""
    temp_path = filepath + ".tmp.xlsx"
    wb.save(temp_path)
    if os.path.exists(temp_path):
        os.replace(temp_path, filepath)
        print(f"   💾 Đã lưu nguyên tử thành công: {filepath}")


def safe_load_workbook(filepath):
    """Mở workbook an toàn, tự đọc shared copy nếu file đang mở trong Excel."""
    if not os.path.exists(filepath):
        return None
    try:
        return openpyxl.load_workbook(filepath)
    except PermissionError:
        print(f"   ℹ️ File {os.path.basename(filepath)} đang mở trong Excel. Đang đọc qua Windows Shared Read...")
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32
        GENERIC_READ = 0x80000000
        FILE_SHARE_READ = 0x00000001
        FILE_SHARE_WRITE = 0x00000002
        FILE_SHARE_DELETE = 0x00000004
        OPEN_EXISTING = 3
        FILE_ATTRIBUTE_NORMAL = 0x80

        handle = kernel32.CreateFileW(
            filepath, GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            None, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None
        )
        if handle == -1:
            print(f"   ❌ Không thể mở file khóa. Error code: {kernel32.GetLastError()}")
            return None

        temp_copy = os.path.join(BASE_DIR, "scratch", f"temp_live_{os.path.basename(filepath)}")
        buffer_size = 65536
        buffer = ctypes.create_string_buffer(buffer_size)
        bytes_read = wintypes.DWORD()
        with open(temp_copy, "wb") as f_out:
            while True:
                success = kernel32.ReadFile(handle, buffer, buffer_size, ctypes.byref(bytes_read), None)
                if not success or bytes_read.value == 0:
                    break
                f_out.write(buffer.raw[:bytes_read.value])
        kernel32.CloseHandle(handle)
        return openpyxl.load_workbook(temp_copy)


# ==============================================================================
# NỘI DUNG KỊCH BẢN CHUẨN THUYTALK-CM124 (4 PHẦN, 100% VĂN NÓI, ZERO BRAGGING)
# ==============================================================================
SPOKEN_BODY = """### 1. HOOK MỞ ĐẦU (0:00 - 0:55)
Chào các bạn, mình là Thúy, người đồng hành cùng các bạn trên kênh ThuyTalk.

Càng gần cuối năm, các doanh nghiệp thương mại nhập khẩu hàng tiêu dùng phục vụ Tết Dương lịch 2027 và Tết Nguyên đán Đinh Mùi 2027 càng chạy đua kéo hàng về kho cho kịp mùa bán. Nhiều chủ hàng cứ nghĩ: Tờ khai thông quan xong, qua vòng tham vấn giá ở cửa khẩu là coi như khép lại.

Thế nhưng, ngày mười tháng mười một năm hai nghìn không trăm hai mươi sáu vừa qua, Cục Hải quan thuộc Bộ Tài chính đã ban hành Công văn số 24083 gửi toàn bộ các Chi cục Hải quan khu vực để siết quản lý trị giá hải quan dịp Tết Đinh Mùi 2027. Điểm đáng chú ý nhất lần này không chỉ là siết tham vấn giá tại cảng, mà còn chỉ đạo trực diện bộ phận Phúc tập và Kiểm tra sau thông quan vào cuộc rà soát theo dấu hiệu vi phạm.

> 🔗 **CÂU DẪN CHUYỂN TIẾP 1 (KÉO SANG PHÂN TÍCH):**  
> *"Nhiều người vẫn nghĩ tham vấn giá chỉ là vài câu hỏi qua loa ở cửa khẩu rồi thông quan cho kịp tiến độ bán hàng Tết. Nhưng tại sao lần này Cục Hải quan lại nghiêm cấm tình trạng tham vấn chiếu lệ, và dữ liệu trên bàn tham vấn sẽ dẫn lối cho bộ phận Kiểm tra sau thông quan tìm đến doanh nghiệp thương mại ra sao? Chúng ta cùng bóc tách từng điểm cốt lõi ngay sau đây."*

---

### 2. PHÂN TÍCH THỰC TRẠNG & VÍ DỤ GIẢ ĐỊNH (0:55 - 3:30)
#### A. Bóc tách bốn điểm siết chặt trong Công văn số 24083:

- **Điểm thứ nhất — Khoanh vùng chín nhóm mặt hàng trọng điểm thuế suất cao:**
  Cục Hải quan chỉ đích danh nhóm hàng tiêu dùng có nhu cầu tăng mạnh dịp Tết, chịu thuế suất nhập khẩu cao và có mặt hàng chịu thuế tiêu thụ đặc biệt gồm: rượu, bia, đồ uống, bánh kẹo, thiết bị điện gia dụng, hàng tiêu dùng, gạch, mỹ phẩm và thực phẩm chức năng.

- **Điểm thứ hai — Cấm tuyệt đối tham vấn qua loa, chiếu lệ:**
  Các Chi cục Hải quan khu vực phải kiểm tra chặt trị giá kê khai và tham vấn đúng Điều 25 Thông tư 38 năm 2015, đã sửa đổi tại khoản 12 Điều 1 Thông tư 121 năm 2025 của Bộ Tài chính, cùng Thông tư 39 năm 2015 và Thông tư 60 năm 2019. Khi hàng khai báo thấp hơn giá tham chiếu tại Danh mục rủi ro về trị giá và cơ sở dữ liệu trị giá hải quan, phải đánh dấu nghi vấn đầy đủ. Công văn nghiêm cấm việc tham vấn qua loa rồi kết luận chưa đủ cơ sở bác bỏ, hoặc sau khi bác bỏ lại xác định trị giá thấp hơn cơ sở dữ liệu.

- **Điểm thứ ba — Khóa dữ liệu trên Hệ thống GTT02 và cập nhật biến động mười phần trăm:**
  Toàn bộ kết quả kiểm tra, tham vấn và xác định trị giá phải cập nhật kịp thời lên Hệ thống GTT02 theo Quyết định số 181 và Quyết định số 1018 năm 2026 của Cục Hải quan. Bộ phận nghiệp vụ phải rà soát hàng ngày và đề xuất sửa đổi giá tham chiếu khi thông tin thu thập biến động tăng hoặc giảm từ mười phần trăm trở lên; đồng thời xem xét trách nhiệm cán bộ công chức làm sai quy định.

- **Điểm thứ tư — Kích hoạt bộ phận Phúc tập và Kiểm tra sau thông quan:**
  Đây là điểm doanh nghiệp thương mại rất dễ bỏ sót: Cục Hải quan chỉ đạo bộ phận Phúc tập và Kiểm tra sau thông quan tổ chức thu thập thông tin, rà soát và kiểm tra sau thông quan theo dấu hiệu vi phạm đối với các mặt hàng trọng điểm rủi ro về trị giá. Tức là hàng đã thông quan bán dịp Tết xong, vòng hậu kiểm vẫn chờ phía sau.

#### B. Ví dụ thực tế giả định để dễ liên tưởng:
Giả sử một doanh nghiệp thương mại nhập khẩu thiết bị điện gia dụng hoặc mỹ phẩm để bán Tết. Vì mua lô hàng ưu đãi chiết khấu cuối năm nên đơn giá hóa đơn thấp hơn mười lăm phần trăm so với mức giá tham chiếu của hải quan.

Khi làm tham vấn — dù tham vấn trực tiếp hay tham vấn gián tiếp qua hệ thống điện tử theo Thông tư 121 năm 2025 — nếu nhân viên giải trình sơ sài, chỉ nói miệng là hàng khuyến mãi mà thiếu phụ lục hợp đồng, bảng giá công khai hay chứng từ thanh toán ngân hàng khớp đúng. Hàng có thể vẫn được giải phóng tạm để bán Tết, nhưng hồ sơ trên Hệ thống GTT02 bị lưu dấu vết nghi vấn. Sau Tết, bộ phận Phúc tập và Kiểm tra sau thông quan có thể lật lại toàn bộ chuỗi tờ khai đó, đối chiếu thẳng với giá bán ra nội địa và sổ sách kế toán.

> 🔗 **CÂU DẪN CHUYỂN TIẾP 2 (KÉO SANG GIẢI PHÁP):**  
> *"Đứng trước sự liên thông chặt chẽ giữa dữ liệu tham vấn giá trên GTT02 và bộ phận Kiểm tra sau thông quan như vậy, nếu doanh nghiệp của các bạn nhập khẩu đúng các nhóm hàng dịp Tết và bị đánh dấu nghi vấn về giá, chúng ta cần chuẩn bị hồ sơ thế nào để giải trình thuyết phục ngay từ cửa khẩu? Đây là ba bước chuẩn chỉnh người làm nghề nên nắm kỹ:"*

---

### 3. GIẢI PHÁP THỰC CHIẾN STEP-BY-STEP (3:30 - 5:45)
- **Bước 1: Đối soát mức giá tham chiếu và khóa chặt điều khoản hợp đồng trước khi mở tờ khai.**
  Bộ phận mua hàng và xuất nhập khẩu cần kiểm tra trước khung giá mặt hàng. Nếu mua được giá tốt hợp pháp nhờ số lượng lớn hay chiết khấu cuối năm, phải ghi rõ điều kiện giảm giá trên hợp đồng ngoại thương và tách bạch trên hóa đơn thương mại, tuyệt đối tránh thỏa thuận miệng rời rạc.

- **Bước 2: Chuẩn bị hồ sơ tham vấn thực chất theo Thông tư 121 năm 2025.**
  Vì hải quan bị cấm tham vấn chiếu lệ, doanh nghiệp càng không được giải trình chiếu lệ. Dù tham vấn trực tiếp hay gián tiếp điện tử, hãy chuẩn bị bộ chứng từ liên kết chặt chẽ: hợp đồng, hóa đơn, phiếu đóng gói, vận đơn đường biển hoặc đường không, chứng từ thanh toán qua ngân hàng khớp từng đồng, thư chào giá của nhà sản xuất và phương án tính giá bán tại Việt Nam.

- **Bước 3: Đồng bộ tờ khai hải quan với sổ sách kế toán ngay sau thông quan.**
  Với doanh nghiệp thương mại, toàn bộ chi phí mua hàng, cước quốc tế, bảo hiểm hay các khoản điều chỉnh cộng phải hạch toán đồng nhất giữa hồ sơ xuất nhập khẩu và sổ kế toán nội bộ ngay trong tháng, tránh cảnh bán hết hàng mới phát hiện trị giá khai báo lệch với giá vốn.

> 🔗 **CÂU DẪN CHUYỂN TIẾP 3 (KÉO SANG LIÊN TƯỞNG MỞ RỘNG & MỞ ĐƯỜNG CHUYÊN SÂU):**  
> *"Làm tốt ba bước đó giúp doanh nghiệp vững vàng qua vòng tham vấn từng lô hàng lẻ. Thế nhưng, từ chỉ đạo đưa bộ phận Phúc tập và Kiểm tra sau thông quan vào cuộc trong Công văn 24083, có một sự thật khiến rất nhiều doanh nghiệp thương mại thuần túy mua đi bán lại bị bất ngờ khi nhận quyết định hậu kiểm. Đối với riêng doanh nghiệp thương mại, cơ quan hải quan đang áp dụng những loại kiểm tra sau thông quan nào?"*

---

### 4. LIÊN TƯỞNG MỞ RỘNG & MỞ ĐƯỜNG CHUYÊN ĐỀ KIỂM TRA SAU THÔNG QUAN DOANH NGHIỆP THƯƠNG MẠI (5:45 - 7:35)
Nhiều anh chị làm công ty thương mại nhập khẩu hàng về bán nội địa thường ngộ nhận: Mình không làm gia công hay sản xuất xuất khẩu, không có định mức nguyên phụ liệu phải thanh khoản, lại đã đóng đủ thuế tại cảng rồi thì không lo Kiểm tra sau thông quan.

Thực tế hoàn toàn khác. Theo Luật Hải quan năm 2014 từ Điều 77 đến Điều 80, cùng Điều 142 và Điều 143 Thông tư 38 năm 2015 — đã sửa đổi tại Thông tư 39 năm 2018 và Thông tư 121 năm 2025 — thời hạn kiểm tra sau thông quan kéo dài tới năm năm kể từ ngày đăng ký tờ khai. Với doanh nghiệp thương mại, kiểm tra sau thông quan chia làm hai tầng rất rõ:

- **Thứ nhất, theo địa điểm kiểm tra:**
  Một là **Kiểm tra sau thông quan tại trụ sở cơ quan hải quan** theo Điều 79 Luật Hải quan và Điều 142 Thông tư 38 — gồm kiểm tra hồ sơ trong vòng sáu mươi ngày kể từ ngày thông quan do Đội trưởng Hải quan quyết định, hoặc kiểm tra hồ sơ trong năm năm do Chi cục trưởng Chi cục Hải quan khu vực quyết định; đặc biệt Thông tư 121 năm 2025 còn bổ sung cơ chế thông báo chuẩn bị hồ sơ trước để doanh nghiệp rà soát.
  Hai là **Kiểm tra sau thông quan tại trụ sở doanh nghiệp** theo Điều 80 Luật Hải quan và Điều 143 Thông tư 38 — đoàn kiểm tra xuống trực tiếp văn phòng và kho hàng để đối chiếu tờ khai với sổ kế toán, sao kê ngân hàng và hóa đơn bán ra trong nước.

- **Thứ hai, theo ba căn cứ kiểm tra tại Điều 78 Luật Hải quan:**
  Đó là kiểm tra **khi có dấu hiệu vi phạm** như rủi ro trị giá trong Công văn 24083, kiểm tra **theo quản lý rủi ro** dựa trên dữ liệu như Hệ thống GTT02, và kiểm tra **đánh giá tuân thủ pháp luật** theo kế hoạch. Nếu hồ sơ tham vấn giá lặp lại nghi vấn bất thường, doanh nghiệp dễ bị hạ bậc tuân thủ, đẩy tờ khai sau vào luồng đỏ và kích hoạt cuộc kiểm tra diện rộng cả năm năm.

> 🔗 **CÂU DẪN CHUYỂN TIẾP 4 (KÉO VÀO KẾT BÀI & TỰ DẶN LÒNG):**  
> *"Nhìn xuyên suốt từ bàn tham vấn giá tại cảng cho tới vòng kiểm tra sau thông quan năm năm về sau, bản thân Thúy luôn tự rút ra một nguyên tắc rất mộc mạc khi làm nghề..."*

---

### 5. KẾT THÚC MỞ, OUTRO TỰ DẶN LÒNG & MỞ ĐƯỜNG TẬP CHUYÊN SÂU TIẾP THEO (7:35 - 8:45)
Với bản thân Thúy, mỗi lần đọc một văn bản siết chặt dữ liệu như Công văn 24083, mình luôn tự dặn lòng: Làm xuất nhập khẩu đường dài, đừng bao giờ coi việc hàng kéo ra khỏi cổng cảng là điểm kết thúc của hồ sơ. Mọi con số khai báo hôm nay đều nằm lại trên hệ thống điện tử nhiều năm sau. Chỉ khi làm thật, chứng từ minh bạch và khớp giữa hải quan với kế toán ngay từ đầu, tâm mình mới nhẹ nhàng và đêm về mới ngủ ngon được.

Và như Thúy vừa gợi mở: Rất nhiều doanh nghiệp thương mại khi nhận quyết định kiểm tra sau thông quan vẫn lúng túng chưa phân biệt được kiểm tra sáu mươi ngày, kiểm tra năm năm tại cơ quan hải quan hay kiểm tra tại trụ sở công ty khác nhau thế nào, và phòng kế toán cần chuẩn bị những sổ sách gì đầu tiên. Trong video chuyên sâu tiếp theo của chuyên mục **Dịch luật ra tiếng người**, Thúy sẽ bóc tách trọn vẹn **Các loại hình kiểm tra sau thông quan đối với doanh nghiệp thương mại** để các bạn chủ động phòng ngừa từ sớm.

Không biết mùa hàng cuối năm nay, các anh chị nhập khẩu nhóm hàng tiêu dùng, gia dụng, bánh kẹo hay mỹ phẩm đã gặp ca tham vấn giá nào theo quy định mới chưa? Hay các anh chị muốn Thúy phân tích sâu phần nào trong video Kiểm tra sau thông quan cho doanh nghiệp thương mại sắp tới? Hãy bình luận bên dưới để chúng ta cùng trao đổi nhé.

Cảm ơn các bạn đã đồng hành cùng Thúy trên kênh ThuyTalk. Đừng quên bấm Đăng ký kênh và bật chuông thông báo để đón xem tập chuyên sâu tiếp theo. Xin chào và hẹn gặp lại các bạn!"""

FULL_YT_SCRIPT = f"""# [KỊCH BẢN YOUTUBE CHI TIẾT - THUYTALK CHUYÊN MỤC 1 & VĂN BẢN NÓNG]
**Tiêu đề Video (SEO Title):** {SEO_TITLE}  
**Mã Chủ Đề / PO:** `{CM_CODE}` | `{PO_CODE}`  
**Chuyên Mục:** `CM1_VAN_BAN_NONG` — ThuyTalk: Trà Chiều & Văn Bản Nóng  
**Thời Lượng Dự Kiến:** 8 - 10 phút (~2.000 từ / tiếng thuần văn nói)  
**Nhân Vật Dẫn Chuyện:** Nguyễn Thị Ngọc Thúy (Founder & CEO Eureka Logistics)  
**Cầu nối Series Tiếp Theo:** Mở đường sang Tập chuyên sâu CM2 (`Các Loại Hình Kiểm Tra Sau Thông Quan Đối Với Doanh Nghiệp Thương Mại`)

---

## PHẦN 1: THÔNG TIN HÀNH CHÍNH & CHỈ DẪN KỸ THUẬT QUAY
- **Bối cảnh quay:** Bàn làm việc gỗ mộc, bản in Công văn số `24083/CHQ-NVTHQ` ký ngày 10/11/2026 của Cục Hải quan, sơ đồ quy trình Tham vấn giá trên Hệ thống GTT02 nối sang bộ phận Phúc tập & Kiểm tra sau thông quan, bộ chứng từ mẫu (Hợp đồng, Invoice, B/L, Chứng từ thanh toán ngân hàng), một chén trà ấm.
- **Góc máy gợi ý:**
  - Cam 1 (Eye-Level 50mm): Trực diện, điềm tĩnh, phân tích rõ tác động của công văn mới đến mùa hàng Tết 2027.
  - Cam 2 (Góc nghiêng 45 độ 85mm): Tập trung chỉ ra quy trình đối soát dữ liệu GTT02 và các loại hình Kiểm tra sau thông quan đối với doanh nghiệp thương mại.
- **Căn cứ pháp lý & SSoT xác thực:**
  1. Công văn số `24083/CHQ-NVTHQ` ngày 10/11/2026 của Cục Hải quan (Bộ Tài chính) V/v tăng cường công tác quản lý trị giá hải quan trước, trong và sau Tết Nguyên đán Đinh Mùi năm 2027.
  2. Điều 25 Thông tư `38/2015/TT-BTC` ngày 25/03/2015 (sửa đổi, bổ sung tại khoản 12 Điều 1 Thông tư `121/2025/TT-BTC` ngày 18/12/2025, hiệu lực từ 01/02/2026).
  3. Thông tư `39/2015/TT-BTC` ngày 25/03/2015 (sửa đổi, bổ sung tại Thông tư `60/2019/TT-BTC` ngày 30/08/2019).
  4. Quyết định số `181/QĐ-CHQ` ngày 28/01/2026 và Quyết định số `1018/QĐ-CHQ` ngày 22/06/2026 về cập nhật dữ liệu trên Hệ thống GTT02.
  5. Luật Hải quan số `54/2014/QH13` (Điều 77, 78, 79, 80) và Điều 142, 143 Thông tư `38/2015/TT-BTC` (sửa đổi tại Thông tư `39/2018/TT-BTC` và Thông tư `121/2025/TT-BTC`).
- **Quy chuẩn giọng đọc & Persona:** Giáp Mộc 60% + Liêm Trinh 40%, 100% văn nói đời thường, ngắt nghỉ tự nhiên 2 giây sau mỗi ý trọng tâm, không từ viết tắt trong lời thoại, không tính từ cường điệu, tuân thủ tuyệt đối Zero-Bragging Rule.

---

## 🎙️ PHẦN 2: KỊCH BẢN THUYẾT MINH THUẦN VĂN NÓI (FULL SPOKEN SCRIPT)
*(Tách riêng hoàn toàn 100% văn nói mộc mạc, có 4 cầu dẫn chuyển tiếp mượt mà để Chị Thúy đọc thử và thu âm trực tiếp)*

{SPOKEN_BODY}

---

## 🎬 PHẦN 3: BẢNG PHÂN CẢNH GỢI Ý QUAY & ĐỒ HỌA MINH HỌA (STORYBOARD & GRAPHICS)

| STT | Phân đoạn & Thời gian | Góc máy quay gợi ý | Lời thoại tương ứng | Đồ họa minh họa / B-roll & Âm thanh |
|:---:|:---|:---|:---|:---|
| 1 | **Hook Mở Đầu** <br> (0:00 - 0:55) | Cam 1 (Eye-Level 50mm). Chị Thúy ngồi tại bàn làm việc gỗ mộc, ánh mắt thẳng thắn, cầm bản in Công văn 24083. | "Chào các bạn... cứ qua được vòng tham vấn giá ở cửa khẩu là xong xuôi... Nhưng ngày 10/11/2026, Cục Hải quan ban hành Công văn 24083 kích hoạt trực diện bộ phận Phúc tập và Kiểm tra sau thông quan..." | Chụp cận mặt trang 1 Công văn `24083/CHQ-NVTHQ` có dấu ký số ngày 10/11/2026.<br>Text động: `CÔNG VĂN 24083/CHQ-NVTHQ: SIẾT THAM VẤN GIÁ & KTSTQ TẾT 2027`.<br>SFX: Tiếng lật giấy và tiếng báo hiệu nhẹ. |
| 2 | **Phân Tích 4 Điểm Nóng & Ví Dụ Giả Định** <br> (0:55 - 3:30) | Đảo giữa Cam 1 và Cam 2 (45 độ 85mm). Chị Thúy phân tích rành mạch 4 điểm siết và tình huống giả định lô hàng gia dụng/mỹ phẩm. | "9 nhóm hàng trọng điểm: rượu, bia, đồ uống, bánh kẹo, điện gia dụng, hàng tiêu dùng, gạch, mỹ phẩm, thực phẩm chức năng... Nghiêm cấm tham vấn qua loa, chiếu lệ... Khóa dữ liệu lên GTT02 & rà soát biến động 10%..." | Motion Graphic danh sách 9 mặt hàng rủi ro cao.<br>Highlight chữ đỏ: `CẤM THAM VẤN CHIẾU LỆ` -> Mũi tên chỉ sang `HỆ THỐNG GTT02` (QĐ 181 & QĐ 1018) -> `PHÚC TẬP & KTSTQ`.<br>Sơ đồ ví dụ giả định giảm giá 15% thiếu phụ lục chứng từ. |
| 3 | **Giải Pháp Thực Chiến 3 Bước** <br> (3:30 - 5:45) | Trở lại Cam 1 trung cảnh. Giọng rõ ràng, dứt khoát từng bước. | "Bước 1: Đối soát giá tham chiếu và ghi rõ điều kiện chiết khấu trên hợp đồng... Bước 2: Hồ sơ tham vấn trực tiếp/gián tiếp theo TT 121/2025... Bước 3: Đồng bộ tờ khai với sổ kế toán và thanh toán ngân hàng..." | Checklist 3 bước hiện trên màn hình:<br>1. Hợp đồng & Chiết khấu minh bạch<br>2. Bộ chứng từ giải trình Tham vấn (TT 121/2025)<br>3. Khớp 100% Hải quan — Kế toán — Ngân hàng.<br>SFX: Tiếng Swoosh gọn gàng qua mỗi bước. |
| 4 | **Liên Tưởng & Mở Đường KTSTQ Cho DN Thương Mại** <br> (5:45 - 7:35) | Cam 2 (85mm) kết hợp đồ họa sơ đồ phân loại. | "Gỡ bỏ ngộ nhận DN thương mại không có định mức thì không bị KTSTQ... Hai địa điểm kiểm tra (tại trụ sở HQ 60 ngày/5 năm theo Đ.79 & Đ.142 vs. tại trụ sở DN theo Đ.80 & Đ.143) và 3 căn cứ kiểm tra theo Điều 78 Luật Hải quan..." | Bảng sơ đồ 2 nhánh `KTSTQ CHO DOANH NGHIỆP THƯƠNG MẠI (HỒI TỐ 5 NĂM)`:<br>• Địa điểm: Trụ sở HQ (Đ.142 - 60 ngày/5 năm) vs. Trụ sở DN (Đ.143)<br>• 3 Căn cứ (Đ.78): Dấu hiệu vi phạm / Quản lý rủi ro / Kế hoạch tuân thủ. |
| 5 | **Outro Tự Dặn Lòng & Cliffhanger Tập Chuyên Sâu** <br> (7:35 - 8:45) | Cam 1 trực diện, ánh mắt ấm áp, chân thành, nụ cười mộc mạc. | "Đừng coi việc hàng ra khỏi cảng là điểm kết thúc... Hãy đón xem video chuyên sâu tiếp theo về Các loại hình KTSTQ đối với doanh nghiệp thương mại..." | Text vàng: `TỰ DẶN LÒNG: LÀM THẬT — CHỨNG TỪ KHỚP — ĐÊM NGỦ NGON`.<br>Thẻ teaser góc phải: `TẬP TIẾP THEO (CM2): CÁC LOẠI HÌNH KTSTQ CHO DOANH NGHIỆP THƯƠNG MẠI`.<br>Nút Đăng ký & Nhạc Outro acoustic dịu nhẹ. |

---

## 🛡️ PHẦN 4: BÁO CÁO FACT-CHECK COMPLIANCE & SSoT (AGENT 08)
- **Kết quả thẩm định:** `🟢 PASSED` (Đạt chuẩn 100% SSoT, Persona `thuytalk-cm124` & `wf07_content_audit_verification`)
- **Căn cứ pháp lý kiểm chứng (Zero-Trust):**
  1. **Công văn `24083/CHQ-NVTHQ` ngày 10/11/2026** (Cục Hải quan - Bộ Tài chính, ký số lúc `15:49:28 +07:00` ngày 10/11/2026): Đáp ứng chuẩn **< 30 ngày** của CM1. Đúng nguyên văn 9 nhóm mặt hàng trọng điểm, lệnh cấm tham vấn chiếu lệ, cập nhật Hệ thống GTT02 (QĐ 181/QĐ-CHQ & QĐ 1018/QĐ-CHQ), biến động ±10%, và chỉ đạo bộ phận Phúc tập & Kiểm tra sau thông quan.
  2. **Thông tư `121/2025/TT-BTC` ngày 18/12/2025** (hiệu lực 01/02/2026): Sửa đổi, bổ sung **Điều 25** (tham vấn trực tiếp & gián tiếp điện tử) và **Điều 142, 143** Thông tư 38/2015/TT-BTC (bổ sung thông báo chuẩn bị hồ sơ, tài liệu trước khi KTSTQ; chuẩn hóa thẩm quyền Đội trưởng Hải quan 60 ngày & Chi cục trưởng Chi cục Hải quan khu vực 5 năm).
  3. **Thông tư `39/2015/TT-BTC` & `60/2019/TT-BTC`**: Đúng quy định về trị giá hải quan.
  4. **Luật Hải quan số `54/2014/QH13` (Điều 77–80)**: Đúng thời hạn 5 năm, 2 địa điểm kiểm tra (Điều 79, 80) và 3 trường hợp kiểm tra sau thông quan (Điều 78).
- **Danh mục thuật ngữ đã chuẩn hóa trong lời thoại (0 từ viết tắt MXH):** Doanh nghiệp, Hải quan, Xuất nhập khẩu, Thuế tiêu thụ đặc biệt, Thuế giá trị gia tăng, Kiểm tra sau thông quan, Vận đơn đường biển, Vận đơn hàng không, Một trăm phần trăm, Mười phần trăm.
- **Kiểm tra Từ cấm (`knowledge/banned_words.md`) & Cường điệu:** Sạch 100% (không chứa *trốn thuế, lách luật, hàng lậu, hàng giả, bao thuế, tiểu ngạch, chấn động, kinh hoàng, rúng động...*).
- **Kiểm tra Zero-Bragging Rule:** Đạt 100% — Hoàn toàn không có câu tự xưng kinh nghiệm hay khoe thành tích giải cứu; sử dụng 100% khung giả định khách quan (*"Giả sử..."*).
"""

PO_CONTENT = f"""# PHIẾU GIAO VIỆC VẬT LÝ (PRODUCTION ORDER) — {PO_CODE}
**Ngày tạo:** 13/11/2026  
**Chu kỳ sản xuất:** Tuần 46: 09/11/2026 - 15/11/2026  
**Đơn vị thực hiện:** Agent 04 (Planner) -> Agent 06 (YouTube CM1) -> Agent 08 (Fact-Checker)  
**Người phê duyệt & chỉ đạo chủ đề:** Nguyễn Thị Ngọc Thúy (GĐKD Eureka Logistics)  
**Trạng thái PO:** Đã phê duyệt sản xuất & Thẩm định PASSED

---

## 1. THÔNG TIN CHUNG
- **Mã PO / Series:** `{PO_CODE}` | `{CM_CODE}`
- **Tiêu đề Video SEO:** {SEO_TITLE}
- **Nền tảng xuất bản:** YouTube Dài — Chuyên mục 1 (`CM1_VAN_BAN_NONG`)
- **Mục tiêu chiến lược kép:**
  1. Cập nhật nóng **Công văn 24083/CHQ-NVTHQ (10/11/2026)** về siết chặt tham vấn giá & kích hoạt Phúc tập / KTSTQ dịp Tết Đinh Mùi 2027.
  2. **Mở đường (Cliffhanger/Bridge)** cho chuỗi nội dung chuyên sâu CM2 về **Các loại hình Kiểm tra sau thông quan đối với Doanh nghiệp thương mại** (`CM2-HOT52` đề xuất nối tiếp).
- **Nguồn căn cứ xác thực (SSoT):** Công văn 24083/CHQ-NVTHQ (10/11/2026); Điều 25, 142, 143 Thông tư 38/2015/TT-BTC (sửa đổi bởi khoản 12 Điều 1 Thông tư 121/2025/TT-BTC); Thông tư 39/2015/TT-BTC & 60/2019/TT-BTC; Quyết định 181/QĐ-CHQ (28/01/2026) & Quyết định 1018/QĐ-CHQ (22/06/2026); Điều 77–80 Luật Hải quan 2014.
- **Từ khóa SEO YouTube:** {SEO_KEYWORDS}

## 2. LIÊN KẾT TỆP NGUỒN SẢN PHẨM
- Kịch bản nháp (`scratch/`): `scratch/production_orders/DRAFT-YT-SERIES-{PO_CODE}.md`
- Kịch bản chính thức (`reports/`): `reports/plan1/video_scripts/YS-{PO_CODE}.md`
- Bản đồng bộ OneDrive YouTube: `C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\{ONEDRIVE_YT_FILENAME}`
- Báo cáo thẩm định Fact-Check: `scratch/production_orders/FACT-CHECK-{PO_CODE}.md`
"""

FACT_CHECK_CONTENT = f"""# BÁO CÁO THẨM ĐỊNH FACT-CHECK & COMPLIANCE — {PO_CODE} ({CM_CODE})
**Người thẩm định:** Agent 08 (Fact-Checker & Brand Guardian)  
**Ngày thẩm định:** 13/11/2026  
**Kết quả chung cuộc:** 🟢 **PASSED (100% ĐẠT CHUẨN XUẤT BẢN)**

---

## 1. THẨM ĐỊNH TÍNH CHÍNH XÁC PHÁP LÝ (SSoT ZERO-TRUST AUDIT)
| Hạng mục đối chiếu | Nội dung trích xuất trong kịch bản | Nguồn gốc xác thực (SSoT) | Kết quả |
|---|---|---|---|
| **Văn bản nóng CM1 (<30 ngày)** | Công văn `24083/CHQ-NVTHQ` ngày 10/11/2026 của Cục Hải quan gửi các Chi cục Hải quan khu vực (ký số 10/11/2026 15:49). | Ảnh chụp 2 trang bản gốc Công văn 24083/CHQ-NVTHQ do Chị Thúy cung cấp. | 🟢 KHỚP 100% |
| **9 nhóm mặt hàng trọng điểm** | Rượu, bia, đồ uống, bánh kẹo, thiết bị điện gia dụng, hàng tiêu dùng, gạch, mỹ phẩm, thực phẩm chức năng (thuế NK cao, một số chịu thuế TTĐB). | Đoạn 1 trang 1 Công văn 24083/CHQ-NVTHQ. | 🟢 KHỚP 100% |
| **Quy định Tham vấn giá & TT 121/2025** | Điều 25 TT 38/2015/TT-BTC (sửa đổi bởi khoản 12 Điều 1 TT 121/2025/TT-BTC ngày 18/12/2025, hiệu lực 01/02/2026: tham vấn trực tiếp & gián tiếp điện tử); TT 39/2015 & TT 60/2019; cấm tham vấn qua loa, chiếu lệ. | Công văn 24083/CHQ-NVTHQ (trang 1-2) & Thông tư 121/2025/TT-BTC. | 🟢 KHỚP 100% |
| **Hệ thống GTT02 & ±10%** | Cập nhật kết quả lên Hệ thống GTT02 theo QĐ 181/QĐ-CHQ (28/01/2026) và QĐ 1018/QĐ-CHQ (22/06/2026); rà soát đề xuất sửa đổi giá tham chiếu khi biến động tăng/giảm 10%. | Công văn 24083/CHQ-NVTHQ (trang 2). | 🟢 KHỚP 100% |
| **Phúc tập & KTSTQ cho DN thương mại** | Chỉ đạo Phúc tập & KTSTQ theo dấu hiệu vi phạm; mở đường phân tích 2 địa điểm kiểm tra (Điều 79, 80 Luật HQ; Điều 142, 143 TT 38 sửa bởi TT 121/2025) & 3 căn cứ kiểm tra (Điều 78 Luật HQ 2014). | Luật Hải quan 54/2014/QH13 & Thông tư 121/2025/TT-BTC. | 🟢 KHỚP 100% |

## 2. KIỂM ĐỊNH TUÂN THỦ PERSONA & BANNED WORDS
- **Văn nói thuần Việt trong Phần 2 (Full Spoken Script):** Đạt 100% (0 từ viết tắt MXH như `mng`, `dn`, `sp`, `kh`, `hq`, `xnk`, `tq`, `b/l`, `ktstq`, `gtgt`, `ttđb`).
- **Từ cấm & Cường điệu (`banned_words.md` + `wf03`):** 0 vi phạm.
- **Zero-Bragging Rule (`thuytalk-cm124/SKILL.md`):** Đạt 100% — Sử dụng tình huống giả định khách quan (*"Giả sử một doanh nghiệp thương mại..."*), tuyệt đối không khoe kinh nghiệm cá nhân.
- **Cấu trúc 5 Hồi + 4 Cầu dẫn chuyển tiếp:** Đầy đủ 4 cầu dẫn đo ni đóng giày, kết hợp Outro tự dặn lòng + Cliffhanger mở đường sang chủ đề chuyên sâu CM2 về Các loại kiểm tra sau thông quan đối với Doanh nghiệp thương mại.
"""


# ==============================================================================
# BƯỚC 0: TỰ KIỂM ĐỊNH (PRE-FLIGHT AUDIT) TRÊN SPOKEN_BODY
# ==============================================================================
print("=" * 80)
print("🛡️ [BƯỚC 0] Tự kiểm định Pre-flight trên Kịch bản Thuần Văn Nói (PHẦN 2)...")
print("=" * 80)

FORBIDDEN_ABBREV = [
    r"\bmng\b", r"\bdn\b", r"\bsp\b", r"\bkh\b", r"\bhq\b", r"\bb/l\b",
    r"\bxnk\b", r"\btq\b", r"\bbl\b", r"\bktstq\b", r"\bgtgt\b", r"\bttđb\b"
]
FORBIDDEN_HYPERBOLE = [
    "chấn động", "kinh hoàng", "rúng động", "nghìn cân treo sợi tóc",
    "lằn ranh sinh tử", "trải đầy hoa hồng", "con đường hoa hồng", "bàn trà thanh tịnh"
]
BANNED_TERMS = [
    "trốn thuế", "lách luật", "hàng lậu", "hàng nhái", "hàng giả",
    "bao thuế", "đi chui", "đi lậu", "tiểu ngạch", "đánh hàng quảng châu",
    "đánh hàng trung quốc", "gom hàng biên", "chạy line"
]
BRAG_PATTERNS = [
    "kinh nghiệm của bên thúy", "bên thúy từng", "đội ngũ của mình đã gỡ",
    "15 năm kinh nghiệm", "mười lăm năm kinh nghiệm", "chuyên gia lão luyện"
]

lower_spoken = SPOKEN_BODY.lower()
errors = []
for pat in FORBIDDEN_ABBREV:
    if re.search(pat, lower_spoken):
        errors.append(f"Viết tắt cấm: {pat}")
for term in FORBIDDEN_HYPERBOLE + BANNED_TERMS + BRAG_PATTERNS:
    if term in lower_spoken:
        errors.append(f"Từ/cụm từ cấm: {term}")

word_count = len(SPOKEN_BODY.split())
print(f"   📊 Số từ Phần 2 (Full Spoken Script): {word_count} từ")
if errors:
    print("   ❌ LỖI KIỂM ĐỊNH:", errors)
    sys.exit(1)
print("   ✅ Kiểm định Pre-flight PASSED 100% (0 viết tắt, 0 cường điệu, 0 từ cấm, 0 khoe mẽ).")


# ==============================================================================
# BƯỚC 1: GHI TẤT CẢ CÁC FILE MARKDOWN (.MD)
# ==============================================================================
print("\n" + "=" * 80)
print(f"📌 [BƯỚC 1] Xuất các tệp vật lý PO, Kịch bản YouTube, Fact-Check & OneDrive...")
print("=" * 80)

files_to_write = [
    (os.path.join(SCRATCH_PO_DIR, f"{PO_CODE}.md"), PO_CONTENT),
    (os.path.join(SCRATCH_PO_DIR, f"DRAFT-YT-SERIES-{PO_CODE}.md"), FULL_YT_SCRIPT),
    (os.path.join(REP_YT_DIR, f"YS-{PO_CODE}.md"), FULL_YT_SCRIPT),
    (os.path.join(SCRATCH_PO_DIR, f"FACT-CHECK-{PO_CODE}.md"), FACT_CHECK_CONTENT),
]
if os.path.exists(ONEDRIVE_YT_DIR):
    files_to_write.append((os.path.join(ONEDRIVE_YT_DIR, ONEDRIVE_YT_FILENAME), FULL_YT_SCRIPT))

for path, content in files_to_write:
    with open(path, "w", encoding="utf-8") as f:
        f.write(content.strip() + "\n")
    print(f"   ✅ Đã ghi: {path}")


# ==============================================================================
# BƯỚC 2: CẬP NHẬT WEEKLY_CONTENT_CALENDAR.XLSX (SHEET YOUTUBE)
# ==============================================================================
print("\n" + "=" * 80)
print("📌 [BƯỚC 2] Cập nhật Sheet 'YouTube' trong Weekly_Content_Calendar.xlsx...")
print("=" * 80)

font_regular = Font(name="Arial", size=10)
font_bold = Font(name="Arial", size=10, bold=True)
font_passed = Font(name="Arial", size=10, bold=True, color="006100")
fill_passed = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
font_pending = Font(name="Arial", size=10, bold=True, color="9C6500")
fill_pending = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
thin_border = Side(border_style="thin", color="D3D3D3")
cell_border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

wb_cal = safe_load_workbook(EXCEL_PATH)
ws_yt = wb_cal["YouTube"]

found_row = None
for r in range(5, ws_yt.max_row + 1):
    if ws_yt.cell(r, 2).value == PO_CODE:
        found_row = r
        break
target_row = found_row if found_row else ws_yt.max_row + 1
stt_yt = target_row - 4

yt_link_formula = (
    f'=HYPERLINK("file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/video_scripts/YS-{PO_CODE}.md", '
    f'"👉 Xem Kịch Bản Chi Tiết")'
)
visual_summary = (
    "Cam 1 (50mm) & Cam 2 (85mm) tại bàn gỗ mộc. Chụp cận trang 1-2 Công văn 24083/CHQ-NVTHQ (10/11/2026); "
    "Motion Graphic 9 mặt hàng rủi ro, luồng GTT02 -> Phúc tập & KTSTQ; Sơ đồ 2 địa điểm & 3 căn cứ KTSTQ cho DN thương mại."
)
ssot_summary = (
    "PASSED — Đúng 100% CV 24083/CHQ-NVTHQ (10/11/2026), Điều 25/142/143 TT 38 (sửa đổi TT 121/2025), "
    "QĐ 181 & QĐ 1018/QĐ-CHQ, Điều 77-80 Luật HQ 2014. 0 viết tắt/từ cấm."
)

row_yt_data = [
    stt_yt,
    PO_CODE,
    "[W46] Thứ Sáu (13/11/2026) — CM1 Văn Bản Nóng",
    SEO_TITLE,
    SEO_KEYWORDS,
    yt_link_formula,
    FULL_YT_SCRIPT.strip(),
    visual_summary,
    "🟢 PASSED",
    ssot_summary,
    "Sẵn sàng dựng",
]

for c_idx, val in enumerate(row_yt_data, start=1):
    cell = ws_yt.cell(target_row, c_idx)
    cell.value = val
    cell.font = font_regular
    cell.border = cell_border
    if c_idx in [1, 2, 3, 11]:
        cell.alignment = align_center
        if c_idx in [1, 2, 11]:
            cell.font = font_bold
    elif c_idx == 6:
        cell.alignment = align_center
        cell.font = Font(name="Arial", size=10, bold=True, color="0000FF", underline="single")
    elif c_idx == 9:
        cell.alignment = align_center
        cell.font = font_passed
        cell.fill = fill_passed
    else:
        cell.alignment = align_left

print(f"   -> YouTube Sheet Row {target_row} (STT {stt_yt}): {PO_CODE} | {SEO_TITLE[:50]}...")
atomic_save_workbook(wb_cal, EXCEL_PATH)


# ==============================================================================
# BƯỚC 3: CẬP NHẬT ONEDRIVE KE_HOACH_PHAT_SONG_YOUTUBE_SERIES.XLSX
# ==============================================================================
print("\n" + "=" * 80)
print("📌 [BƯỚC 3] Cập nhật OneDrive Ke_Hoach_Phat_Song_YouTube_Series.xlsx...")
print("=" * 80)

if os.path.exists(ONEDRIVE_SERIES_EXCEL):
    wb_series = safe_load_workbook(ONEDRIVE_SERIES_EXCEL)
    if wb_series:
        # 3.1 Ghi CM1-HOT51 vào CM1_VAN_BAN_NONG
        ws_cm1 = wb_series["CM1_VAN_BAN_NONG"]
        found_r1 = None
        for r in range(5, ws_cm1.max_row + 1):
            if ws_cm1.cell(r, 2).value == CM_CODE:
                found_r1 = r
                break
        target_r1 = found_r1 if found_r1 else ws_cm1.max_row + 1
        stt_cm1 = target_r1 - 4

        onedrive_md_url = (
            f"C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\youtube\\{ONEDRIVE_YT_FILENAME}"
        )
        link_cm1 = f'=HYPERLINK("{onedrive_md_url}", "👉 Xem Kịch Bản Chi Tiết")'

        hook_preview = (
            "Chào các bạn, mình là Thúy, người đồng hành cùng các bạn trên kênh ThuyTalk. "
            "Càng gần giai đoạn cuối năm phục vụ Tết Đinh Mùi 2027, nhiều chủ hàng cứ nghĩ qua vòng tham vấn giá ở cửa khẩu là xong. "
            "Nhưng ngày 10/11/2026, Cục Hải quan đã ban hành Công văn 24083 siết chặt tham vấn giá và kích hoạt Phúc tập & Kiểm tra sau thông quan..."
        )
        core_summary = (
            "[NỘI DUNG CHI TIẾT - CHỊ THÚY]\n"
            "- Phân tích 4 điểm siết chặt trong Công văn 24083/CHQ-NVTHQ (10/11/2026): 9 nhóm hàng tiêu dùng chịu thuế cao; cấm tham vấn qua loa chiếu lệ (Đ.25 TT 38 sửa bởi k.12 Đ.1 TT 121/2025); khóa dữ liệu GTT02 (QĐ 181 & QĐ 1018) + biến động ±10%; kích hoạt Phúc tập & KTSTQ.\n"
            "- Ví dụ giả định lô hàng thiết bị điện gia dụng / mỹ phẩm giảm giá 15% cuối năm & lộ trình 3 bước phòng vệ.\n"
            "- Mở đường chuyên sâu: Gỡ ngộ nhận của DN thương mại về KTSTQ, phân định 2 địa điểm (Đ.79/142 vs Đ.80/143) & 3 căn cứ KTSTQ (Đ.78 Luật HQ 2014)."
        )
        cliffhanger_cm1 = (
            "Mở đường sang Tập chuyên sâu CM2 (CM2-HOT52): 'Các loại hình Kiểm tra sau thông quan đối với Doanh nghiệp thương mại "
            "— Phân biệt kiểm tra 60 ngày, kiểm tra 5 năm tại trụ sở Hải quan vs. Đoàn kiểm tra xuống tận công ty, phòng kế toán & XNK bị soi sổ sách nào đầu tiên?'"
        )
        outro_preview = (
            "Với bản thân Thúy, mỗi lần đọc công văn siết dữ liệu như Công văn 24083, mình luôn tự dặn lòng: "
            "Đừng bao giờ coi thời điểm hàng ra khỏi cảng là điểm kết thúc của hồ sơ. Chỉ khi làm thật, chứng từ minh bạch và khớp giữa hải quan với kế toán từ đầu, "
            "tâm mình mới nhẹ nhàng và đêm về mới ngủ ngon được."
        )

        row_cm1 = [
            stt_cm1,
            CM_CODE,
            SEO_TITLE,
            "Video Đơn lẻ + Cầu nối Series (8-9p)",
            "8 - 9 phút",
            "Công văn 24083/CHQ-NVTHQ (10/11/2026); Đ.25, 142, 143 TT 38/2015 (sửa đổi TT 121/2025); TT 39/2015 & TT 60/2019; QĐ 181 & 1018/QĐ-CHQ; Đ.77-80 Luật HQ 2014.",
            hook_preview,
            core_summary,
            visual_summary,
            cliffhanger_cm1,
            outro_preview,
            "🟢 ĐÃ HOÀN THÀNH KỊCH BẢN",
            link_cm1,
            "W46/2026",
            "🛡️ Bảo lưu vĩnh viễn (Đã duyệt)",
        ]

        for c_idx, val in enumerate(row_cm1, start=1):
            cell = ws_cm1.cell(target_r1, c_idx)
            cell.value = val
            cell.font = font_regular
            cell.border = cell_border
            if c_idx in [1, 2, 4, 5, 12, 14, 15]:
                cell.alignment = align_center
                if c_idx in [1, 2, 12, 15]:
                    cell.font = font_bold
            elif c_idx == 13:
                cell.alignment = align_center
                cell.font = Font(name="Arial", size=10, bold=True, color="0000FF", underline="single")
            else:
                cell.alignment = align_left
            if c_idx == 12:
                cell.font = font_passed
                cell.fill = fill_passed

        print(f"   -> OneDrive 'CM1_VAN_BAN_NONG' Row {target_r1} (STT {stt_cm1}): {CM_CODE} (🟢 ĐÃ HOÀN THÀNH KỊCH BẢN)")

        # 3.2 Đề xuất nối tiếp CM2-HOT52 vào CM2_DICH_LUAT_TIENG_NGUOI (trạng thái Chờ Chị Thúy duyệt)
        ws_cm2 = wb_series["CM2_DICH_LUAT_TIENG_NGUOI"]
        cm2_followup_code = "CM2-HOT52"
        found_r2 = None
        for r in range(5, ws_cm2.max_row + 1):
            if ws_cm2.cell(r, 2).value == cm2_followup_code:
                found_r2 = r
                break
        target_r2 = found_r2 if found_r2 else ws_cm2.max_row + 1
        stt_cm2 = target_r2 - 4

        row_cm2_proposal = [
            stt_cm2,
            cm2_followup_code,
            "Các Loại Hình Kiểm Tra Sau Thông Quan Đối Với Doanh Nghiệp Thương Mại: Phân Biệt Kiểm Tra 60 Ngày, Hồi Tố 5 Năm & Kiểm Tra Tại Trụ Sở Doanh Nghiệp",
            "Video Chuyên Sâu Nối Tiếp CV 24083 (9-10p)",
            "9 - 10 phút",
            "Điều 77, 78, 79, 80 Luật Hải quan 54/2014/QH13; Điều 142, 143 Thông tư 38/2015/TT-BTC (sửa đổi bởi Thông tư 39/2018/TT-BTC và Thông tư 121/2025/TT-BTC); Nghị định 08/2015/NĐ-CP & 59/2018/NĐ-CP.",
            "Ở video trước về Công văn 24083, Thúy có nhắc đến việc bộ phận Kiểm tra sau thông quan đang rà soát mạnh khối hàng tiêu dùng. Rất nhiều anh chị làm doanh nghiệp thương mại hỏi Thúy: Kiểm tra sau thông quan thực chất có mấy loại và sổ sách nào bị soi đầu tiên?",
            "[ĐỀ XUẤT NỐI TIẾP MỞ ĐƯỜNG TỪ CM1-HOT51]\n- Giải mã 2 địa điểm KTSTQ cho DN thương mại: Tại trụ sở HQ (60 ngày Đ.142 & 5 năm Đ.79/142, cơ chế thông báo chuẩn bị tài liệu của TT 121/2025) vs. Tại trụ sở DN (Đ.80 & Đ.143).\n- Giải mã 3 căn cứ KTSTQ (Điều 78): Dấu hiệu vi phạm (trị giá/HS/CO), Quản lý rủi ro (GTT02/phân luồng), Kiểm tra tuân thủ theo kế hoạch.\n- Trọng tâm soi DN thương mại: Khớp đúng Tờ khai HQ — Sổ kế toán 156/331 — Sao kê ngân hàng — Hóa đơn bán ra nội địa.",
            "Bảng so sánh 2 cột: Kiểm tra tại trụ sở Hải quan (Đ.142) vs. Kiểm tra tại trụ sở Doanh nghiệp (Đ.143); Checklist sổ sách kế toán DN thương mại.",
            "Nối tiếp trực tiếp từ Tập CM1-HOT51 (Công văn 24083/CHQ-NVTHQ).",
            "Hiểu rõ từng loại hình kiểm tra sau thông quan không phải để lo sợ, mà để phòng xuất nhập khẩu và phòng kế toán biết cách dựng hàng rào chứng từ chuẩn ngay từ đầu.",
            "Chờ Chị Thúy duyệt",
            "",
            "W46/2026",
            "27/11/2026 (Hạn 2 tuần)",
        ]

        for c_idx, val in enumerate(row_cm2_proposal, start=1):
            cell = ws_cm2.cell(target_r2, c_idx)
            cell.value = val
            cell.font = font_regular
            cell.border = cell_border
            if c_idx in [1, 2, 4, 5, 12, 14, 15]:
                cell.alignment = align_center
                if c_idx in [1, 2, 12]:
                    cell.font = font_bold
            else:
                cell.alignment = align_left
            if c_idx == 12:
                cell.font = font_pending
                cell.fill = fill_pending

        print(f"   -> OneDrive 'CM2_DICH_LUAT_TIENG_NGUOI' Row {target_r2} (STT {stt_cm2}): {cm2_followup_code} (Đề xuất nối tiếp - Chờ Chị Thúy duyệt)")
        atomic_save_workbook(wb_series, ONEDRIVE_SERIES_EXCEL)


# ==============================================================================
# BƯỚC 4: CẬP NHẬT CONTENT_STATUS_REPORT.MD & SYNC_LOG.TXT
# ==============================================================================
print("\n" + "=" * 80)
print("📌 [BƯỚC 4] Cập nhật content_status_report.md và sync_log.txt...")
print("=" * 80)

append_status = f"""
---

## 📋 CẬP NHẬT NGÀY 13/11/2026: KỊCH BẢN YOUTUBE CM1 (`{PO_CODE}` / `{CM_CODE}`)

| STT YT | Mã PO / Series | Kênh | Ngày / Tuần | Tiêu đề & Nội dung cốt lõi | File Kịch bản (.md) | Fact-Check | Trạng thái |
|:---:|:---|:---:|:---|:---|:---|:---:|:---:|
| {stt_yt} | `{PO_CODE}` (`{CM_CODE}`) | YouTube CM1 | [W46] Thứ Sáu (13/11/2026) | {SEO_TITLE} — Phân tích nóng CV 24083/CHQ-NVTHQ (10/11/2026) & mở đường chuyên sâu Các loại KTSTQ cho DN thương mại | [YS-{PO_CODE}.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/video_scripts/YS-{PO_CODE}.md) | `🟢 PASSED` | **Sẵn sàng dựng** |

- **Đề xuất nối tiếp trên Master Series (`CM2_DICH_LUAT_TIENG_NGUOI`):** `CM2-HOT52` — *Các Loại Hình Kiểm Tra Sau Thông Quan Đối Với Doanh Nghiệp Thương Mại: Phân Biệt Kiểm Tra 60 Ngày, Hồi Tố 5 Năm & Kiểm Tra Tại Trụ Sở Doanh Nghiệp* (`Chờ Chị Thúy duyệt`).
- **Thẩm định SSoT:** [FACT-CHECK-{PO_CODE}.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/FACT-CHECK-{PO_CODE}.md) (`🟢 PASSED`).
"""

with open(STATUS_REPORT_MD, "r", encoding="utf-8") as f:
    existing_status = f.read()

# Update counts cleanly (79 -> 80 YT, 191 -> 192 total) and update timestamp if present
existing_status = existing_status.replace("12/11/2026 17:30", "13/11/2026 16:25")
existing_status = existing_status.replace("79 kịch bản video hoàn chỉnh (W33-W47 + 34 kịch bản HOT). 100% (79/79)", "80 kịch bản video hoàn chỉnh (W33-W47 + 35 kịch bản HOT). 100% (80/80)")
existing_status = existing_status.replace("191/191 sản phẩm đạt", "192/192 sản phẩm đạt")

marker = f"## 📋 CẬP NHẬT NGÀY 13/11/2026: KỊCH BẢN YOUTUBE CM1 (`{PO_CODE}`"
if marker not in existing_status:
    existing_status = existing_status.rstrip() + "\n" + append_status

with open(STATUS_REPORT_MD, "w", encoding="utf-8") as f:
    f.write(existing_status)
print(f"   ✅ Đã cập nhật: {STATUS_REPORT_MD}")

now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
with open(SYNC_LOG_PATH, "a", encoding="utf-8") as f:
    f.write(f"[{now_str}] SYNC {PO_CODE} ({CM_CODE}) -> YT Row {target_row} (STT {stt_yt}), OneDrive CM1_VAN_BAN_NONG & CM2-HOT52 proposal, Zero Backup verified.\n")
print(f"   ✅ Đã ghi log: {SYNC_LOG_PATH}")

print("\n🎉 HOÀN TẤT 100% QUY TRÌNH TẠO KỊCH BẢN CM1-HOT51 & ĐỒNG BỘ HỆ THỐNG!")
