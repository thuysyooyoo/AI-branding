import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import subprocess
import os

# Stop Excel if running
subprocess.run(['powershell', '-Command', 'Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force'], check=False)

# Master definitions
# ANH.XLSX: 5 columns: STT, Tiêu đề, Prompt tạo ảnh (Image Prompt), Góc quay (Shot Type), Ghi chú
# PHAN CANH.XLSX: 5 columns: STT, Tiêu đề, Mô tả phân cảnh (Prompt), Lời thoại (Dialogue), Thời lượng (giây)

anh_rows = [
    [
        1,
        'Câu Đố Hóc Búa Bao Bì Thực Phẩm',
        'Medium shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive large doodle eyes with eyelashes, holding a cute cartoon magnifying glass and winking playfully at camera on clean pure white background. Three floating doodle items revolve around her with glowing yellow question marks: a clear plastic food container, an orange silicone cake mold, and a shiny metal spoon. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (bright yellow, orange, cyan), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'Khung hình bắt đầu (Start Frame): Cô bé người que Thúy nơ vàng mở đầu bằng câu đố hóc búa, 3 món đồ xoay quanh kèm dấu hỏi để hút tương tác ngay 3 giây đầu.'
    ],
    [
        2,
        'Vén Màn Bí Mật Quy Định Bao Bì',
        'Medium shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive doodle eyes, holding open a vibrant yellow doodle curtain on clean pure white background, revealing organized neat food containers and official regulatory scrolls radiating golden sparkle lines. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (bright yellow, royal blue, emerald green), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'Thúy kéo rèm vàng vén màn bí mật toàn bộ quy định kiểm soát bao bì chứa đựng thực phẩm.'
    ],
    [
        3,
        'Thách Thức Nhận Thức - Sự Thật Ngược Lại',
        'Close-up shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive large doodle eyes with eyelashes, playfully wagging an index finger with a knowing intriguing smile on clean pure white background. Above her floats an inverted lightbulb doodle with comic shock lines and bold red text "SỰ THẬT NGƯỢC LẠI!". Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (alert red, lemon yellow, crisp black), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Thúy lắc ngón tay bí hiểm cảnh báo những ngộ nhận pháp lý mà nhiều người tưởng mình đã biết rất rõ.'
    ],
    [
        4,
        'Nghịch Lý Kẹt Cảng & Kéo Dài Thời Gian',
        'Wide shot split-screen 2D doodle illustration on clean pure white background: Left half shows a cute shipping truck speeding past a bright green signal with "1 NGÀY" badge; Right half shows a cargo container locked by heavy black chains and an oversized red padlock with bold red comic text "KẸT CỨNG CẢ TUẦN!". Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (emerald green, alert red, dark navy), 8k resolution --ar 16:9',
        'Toàn cảnh (Wide Shot)',
        'Màn hình chia đôi đối lập trực quan: 1 bên thông quan 1 ngày bon bon, 1 bên bị xích khóa kẹt cứng cả tuần tại cảng.'
    ],
    [
        5,
        'Cận Cảnh Túi Màng Nhôm Vô Trùng',
        'Close-up shot 2D doodle illustration of a multi-layer aseptic aluminum foil milk pouch packaging bag, showing detailed hand-drawn cross-section layers (plastic film, aluminum barrier, aseptic coating) with technical doodle callout arrows, labeled \'TÚI MÀNG NHÔM TIỆT TRÙNG\', beside an official self-declaration certificate document stamped with blue ink on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (cyan blue, metallic silver, yellow), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Cận cảnh kỹ thuật túi màng nhôm vô trùng 4 lớp và bản Tự công bố minh họa case study thực tế.'
    ],
    [
        6,
        'Bẫy Ngộ Nhận Thông Tư 27',
        'Close-up shot 2D doodle illustration of an open official legal document titled \'THÔNG TƯ 27\', marked with a large red cartoon interrogation question mark and a bold red stamp \'DỪNG THÔNG QUAN - KIỂM TRA NHÀ NƯỚC\', with sharp comic stress lines indicating an unexpected regulatory roadblock on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (alert red, bold black, canary yellow), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Dấu mộc đỏ dừng hàng đập thẳng vào văn bản pháp lý thể hiện cú sốc khi bị dừng thông quan đột ngột.'
    ],
    [
        7,
        'Khẳng Định Bản Chất Văn Bản 22082',
        'Medium shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive large doodle eyes with eyelashes, holding a glowing golden legal gavel in one hand and presenting official dispatch documents labeled \'CV 22082\' and \'CV 1913\' with glowing green checkmarks on clean pure white background, giving an encouraging confident smile. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (emerald green, golden yellow, navy blue), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'Thúy cầm búa thẩm phán và 2 công văn giải cứu tạo niềm tin chuyên gia cho khán giả.'
    ],
    [
        8,
        'Đồ Họa Tháp 3 Tầng Pháp Lý',
        'Wide shot 2D doodle infographic illustration of a clean hand-drawn 3-tiered legal compliance pyramid on clean pure white background: Tier 1 at base labeled \'TẦNG 1: CỬA KHẨU (TT 15)\', Tier 2 in middle labeled \'TẦNG 2: TỰ CÔNG BỐ (NĐ 15)\', Tier 3 at apex labeled \'TẦNG 3: QCVN AN TOÀN\', with a glowing golden key floating above emitting comic sparkle lines. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (golden yellow, cyan blue, emerald green), 8k resolution --ar 16:9',
        'Toàn cảnh (Wide Shot)',
        'Tháp 3 tầng pháp lý dạng doodle phân cấp rõ ràng trên nền trắng, có chìa khóa vàng mở khóa toàn bộ bí quyết.'
    ],
    [
        9,
        'Giảng Giải Tầng 1 - Cửa Khẩu Hải Quan',
        'Medium shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive doodle eyes, standing beside a doodle blackboard presenting Tier 1, pointing a wooden pointer at an official customs gate icon labeled \'CỬA KHẨU HẢI QUAN - THÔNG TƯ 15/2018/TT-BYT\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (navy blue, warning yellow, bright red), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'Thúy chỉ bảng giải thích Tầng 1 kiểm tra cửa khẩu chỉ soi vào mã HS theo Thông tư 15.'
    ],
    [
        10,
        'Cận Cảnh Nhựa PP & Núm Vú Cao Su',
        'Close-up shot 2D doodle illustration of two regulated products: a transparent plastic container labeled \'NHỰA PP / PE\' with recycling logo, and an infant baby silicone bottle nipple labeled \'MÃ HS: 4014.90.10\', encircled by a bold red dashed warning circle and a red shield icon labeled \'BẮT BUỘC KIỂM TRA NHÀ NƯỚC\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (alert red, cyan blue, bright yellow), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Cận cảnh 2 nhóm hàng bị kiểm tra tại cửa khẩu: hộp nhựa thực phẩm và núm vú trẻ em.'
    ],
    [
        11,
        'Danh Mục Sản Phẩm Miễn Cửa Khẩu',
        'Medium shot 2D doodle illustration of exempt tableware items: a stainless steel pot, a ceramic bowl, a glass cup, and an adult rubber spatula, all framed inside a bright green shield with a large vibrant green checkmark and bold text \'MIỄN KIỂM TRA CỬA KHẨU - THÔNG QUAN NGAY\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (vibrant green, cool grey, sunny yellow), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'Bộ dụng cụ Inox, gốm sứ, thủy tinh được đóng khung khiên xanh thông quan ngay lập tức.'
    ],
    [
        12,
        'Giảng Giải Tầng 2 - Thủ Tục Tự Công Bố',
        'Medium shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive doodle eyes, holding an official dossier folder labeled \'HỒ SƠ TỰ CÔNG BỐ - NGHỊ ĐỊNH 15/2018\', sending a paper airplane document towards a cartoon government building labeled \'BAN QUẢN LÝ AN TOÀN THỰC PHẨM\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (cobalt blue, bright yellow, crisp black), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'Thúy ném máy bay giấy hồ sơ về Ban ATTP giải thích Tầng 2 quản lý nội địa để lưu hành hàng.'
    ],
    [
        13,
        'Phòng Lab Thử Nghiệm Chuẩn QCVN',
        'Close-up shot 2D doodle illustration of a cartoon scientific laboratory setup on clean pure white background: glass test tubes with bubbling colorful liquids, a test flask, a microscope, and an official analytical test certificate with an embossed golden seal labeled \'ĐẠT CHUẨN QCVN AN TOÀN THỰC PHẨM\'. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (amber gold, sky blue, emerald green), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Phòng lab doodle kiểm nghiệm thôi nhiễm an toàn theo QCVN - chốt chặn bắt buộc trước khi ra thị trường.'
    ],
    [
        14,
        'Cạm Bẫy Cao Su Trẻ Em vs Silicone',
        'Close-up shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive large doodle eyes with eyelashes, leaning forward with an intense engaged analytical expression, holding a baby pacifier in left hand and a silicone cake mold in right hand, with a flashing yellow warning badge \'CẠM BẪY PHÂN LOẠI 80% MẮC PHẢI\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (warning yellow, alert red, crisp black), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Thúy chỉ ra điểm thắt cổ chai: bẫy phân loại giữa cao su trẻ em và silicone người lớn.'
    ],
    [
        15,
        'Bóc Tách Nghịch Lý Khuôn Silicone Người Lớn',
        'Close-up split 2D doodle diagram comparison on clean pure white background: Left half shows silicone cake mold with a green customs barrier gate \'MIỄN KIỂM TRA CỬA KHẨU\'; Right half shows the same mold with an official document labeled \'MỤC 6 THÔNG TƯ 27: BẮT BUỘC TỰ CÔNG BỐ\' stamped with a blue seal. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (emerald green, navy blue, canary yellow), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Sơ đồ bóc tách nghịch lý: Khuôn silicone được miễn cửa khẩu nhưng vẫn bắt buộc làm Tự công bố.'
    ],
    [
        16,
        'Đồ Gia Dụng Inox & Gốm Sứ Nhóm 1',
        'Close-up shot 2D doodle illustration of a stainless steel cooking pan, a porcelain soup bowl, and a glass water tumbler, framed with a prominent green doodle ribbon badge labeled \'NHÓM 1: RỦI RO THẤP - MIỄN NỘP HỒ SƠ LÊN SỞ\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (vibrant green, cool grey, sunny yellow), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Đồ gia dụng Inox, gốm sứ rủi ro thấp được dán nhãn xanh miễn nộp hồ sơ lên Sở.'
    ],
    [
        17,
        'Hướng Dẫn Hồ Sơ TCCS Phòng Thủ Hậu Kiểm',
        'Close-up shot 2D doodle illustration of a neatly organized company legal archive binder labeled \'HỒ SƠ TIÊU CHUẨN CƠ SỞ (TCCS) NỘI BỘ\', containing laboratory test certificates and a sturdy blue cartoon defense shield labeled \'LƯU TRỮ PHÒNG THỦ HẬU KIỂM\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (royal blue, emerald green, warm yellow), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Hồ sơ TCCS nội bộ và khiên phòng thủ giúp doanh nghiệp vững tâm khi QLTT hậu kiểm.'
    ],
    [
        18,
        'Tối Ưu Tiết Kiệm Chi Phí Hàng Triệu Đồng',
        'Medium shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive doodle eyes, holding a cute cartoon calculator showing \'TIẾT KIỆM 100%\', with flying cartoon VND currency notes and a cheerful celebration expression, surrounded by floating green plus signs and coin icons on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (money green, golden yellow, crisp black), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'Thúy cầm máy tính bấm số tiết kiệm tiền thủ tục cho doanh nghiệp, biểu cảm rạng rỡ.'
    ],
    [
        19,
        'So Sánh Dụng Cụ Chứa Đựng vs Tiếp Xúc',
        'Medium shot 2D doodle comparison diagram on clean pure white background: Left side shows a deep food storage container labeled \'DỤNG CỤ CHỨA ĐỰNG (Container)\'; Right side shows a fork, spoon, and chopsticks labeled \'DỤNG CỤ TIẾP XÚC CƠ HỌC (Contact Only)\', separated by a bold yellow lightning bolt question mark. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (sunburst yellow, bright red, navy blue), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'So sánh trực quan giữa đồ chứa đựng và đồ tiếp xúc cơ học - điểm mấu chốt thứ hai.'
    ],
    [
        20,
        'Kim Bài Công Văn 4372 Bộ Y Tế (Aeon)',
        'Close-up shot 2D doodle illustration of a prestigious official government dispatch scroll titled \'CÔNG VĂN 4372/BYT-ATTP (2019)\', glowing with a golden aura and an official red Ministry seal, labeled \'CĂN CỨ VÀNG - BỘ Y TẾ GỬI AEON VIỆT NAM\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (imperial gold, official red, deep navy), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Công văn 4372/BYT như một kim bài pháp lý vàng tháo gỡ toàn bộ thủ tục cho dụng cụ bàn ăn.'
    ],
    [
        21,
        'Thìa Đũa Dụng Cụ Thao Tác Cơ Học',
        'Close-up shot 2D doodle illustration of dining cutlery: wooden chopsticks, stainless steel spoon, table fork, fruit paring knife, silicone cooking tongs, and food prep gloves, arranged neatly with animated comic sparkles and labeled \'DỤNG CỤ THAO TÁC CƠ HỌC - KHÔNG PHẢI BAO BÌ CHỨA ĐỰNG\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (bright cyan, lemon yellow, crisp black), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Cận cảnh thìa, đũa, dĩa, dao, kẹp gắp được xếp ngay ngắn dạng doodle.'
    ],
    [
        22,
        'Khẳng Định Miễn TCB & Miễn Kiểm Tra Cửa Khẩu',
        'Medium shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive doodle eyes, holding two giant green checkmark shields labeled \'MIỄN TỰ CÔNG BỐ\' and \'MIỄN KIỂM TRA CỬA KHẨU\' on clean pure white background, with a joyful confident smile. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (emerald green, golden yellow, crisp black), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'Thúy giơ 2 khiên xanh khẳng định miễn 100% cả 2 thủ tục cho nhóm thìa đũa tiếp xúc.'
    ],
    [
        23,
        'Xe Container Thông Quan Luồng Xanh Ra Cảng',
        'Wide shot 2D doodle illustration of a cute semi-truck carrying a shipping container zooming smoothly through a green customs gate with a glowing green traffic light, leaving green speed lines and a banner labeled \'THÔNG QUAN TRONG NGÀY - PHÍ KIỂM TRA = 0Đ\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (speed green, sunny yellow, cobalt blue), 8k resolution --ar 16:9',
        'Toàn cảnh (Wide Shot)',
        'Xe container bon bon qua luồng xanh hải quan, không tốn một đồng phí kiểm tra.'
    ],
    [
        24,
        'Cảnh Báo Mốc Thời Gian Lab 10-12 Ngày',
        'Close-up shot 2D doodle illustration of a flip calendar with pages tearing rapidly, stopping at the bold highlighted date range \'10 - 12 NGÀY LÀM VIỆC\', beside a ticking stopwatch with spinning hands and comic sweat drops indicating urgency on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (warning orange, bold red, royal blue), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Đồng hồ và tờ lịch chạy nước rút 10-12 ngày thể hiện áp lực thời gian thử nghiệm.'
    ],
    [
        25,
        'Ác Mộng Hóa Đơn Phạt Dem/Det 15-30 Triệu',
        'Close-up shot 2D doodle illustration of a giant menacing paper bill invoice labeled \'HÓA ĐƠN PHẠT DEM / DET LƯU CONT LƯU BÃI\' with a shocking total of \'15.000.000 - 30.000.000 VNĐ\', showing a cartoon container box crying comic tears under heavy chains on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (danger red, coal black, warning yellow), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Hóa đơn phạt lưu cont đỏ chót đè nặng lên thùng container khóc ròng.'
    ],
    [
        26,
        'Tuyệt Chiêu Gửi Mẫu CPN Trước 25 Ngày',
        'Medium shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive doodle eyes, holding a cartoon express courier package with wings and a timeline arrow marked \'GỬI MẪU TEST TRƯỚC 25 NGÀY TÀU CHẠY\' on clean pure white background, smiling brightly like a master strategist. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (express orange, golden yellow, crisp black), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'Thúy hướng dẫn tuyệt chiêu gửi mẫu CPN trước 25 ngày như một chiến lược gia logistics.'
    ],
    [
        27,
        'Màn Hình Cổng Một Cửa NSW Đạt Chuẩn',
        'Close-up shot 2D doodle illustration of a cartoon computer monitor displaying the National Single Window portal (\'CỔNG THÔNG TIN MỘT CỬA QUỐC GIA\'), showing a green approval banner \'HỒ SƠ HỢP LỆ - THÔNG QUAN 3 NGÀY\' with flying digital checkmarks on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (bright green, monitor blue, sunny yellow), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Màn hình Cổng Một Cửa Quốc gia NSW sáng đèn xanh duyệt hồ sơ trong 3 ngày.'
    ],
    [
        28,
        'Xe Cont Cập Kho Riêng Cắt Phí Lưu Bãi',
        'Wide shot 2D doodle illustration of a clean cartoon logistics warehouse building labeled \'KHO RIÊNG DOANH NGHIỆP - BẢO QUẢN HÀNG HÓA\', with a container truck safely parked inside unloading cargo, and a large red scissors cutting a chain labeled \'PHÍ LƯU BÃI = 0Đ\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (warehouse navy, cut red, sunny yellow), 8k resolution --ar 16:9',
        'Toàn cảnh (Wide Shot)',
        'Kéo đỏ cắt phăng sợi xích phí lưu kho bãi khi đưa hàng về kho riêng bảo quản.'
    ],
    [
        29,
        'Bản Tự Công Bố Có Giá Trị Vô Thời Hạn',
        'Medium shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive doodle eyes, holding an official certificate displaying a glowing golden infinity symbol \'∞\' labeled \'HIỆU LỰC VÔ THỜI HẠN CHO MỖI MODEL\' on clean pure white background, radiating warm golden light beams. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (radiant gold, royal blue, crisp black), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'Thúy giơ biểu tượng vô cực vàng khẳng định giá trị vĩnh viễn của bản Tự công bố.'
    ],
    [
        30,
        'Lưu Ý Phiếu Kiểm Nghiệm Hạn 12 Tháng',
        'Close-up shot 2D doodle illustration of an analytical test report document with a red warning badge stamped \'THỜI HẠN 12 THÁNG\', beside an animated doodle hourglass with blue sand running out on clean pure white background, showing comic clock hands and an alert exclamation mark. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (alert red, cyan blue, golden yellow), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Đồng hồ cát và dấu đỏ lưu ý hạn 12 tháng của phiếu test khi nộp hồ sơ.'
    ],
    [
        31,
        'Đúc Kết Thông Điệp Triết Lý Thực Chiến',
        'Close-up shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive large doodle eyes with eyelashes, with an inspiring, visionary, determined expression, holding a glowing golden lightbulb representing proactive logistics foresight, with a sturdy green fortress shield behind her on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (inspirational gold, fortress green, crisp black), 8k resolution --ar 16:9',
        'Cận cảnh (Close-up)',
        'Cận cảnh Thúy đúc kết triết lý: chủ động xây dựng hàng rào pháp lý trước khi hàng rời xưởng.'
    ],
    [
        32,
        'Outro Kêu Gọi Subscribe & Tặng Cẩm Nang',
        'Medium shot of cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive doodle eyes, waving warmly with both hands, surrounded by a red YouTube \'ĐĂNG KÝ KÊNH\' button, a glowing bell icon, and a floating golden gift book titled \'CẨM NANG BAO BÌ THỰC PHẨM\' on clean pure white background. Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors (youtube red, golden yellow, mint green), 8k resolution --ar 16:9',
        'Trung cảnh (Medium Shot)',
        'Outro Thúy vẫy tay chào, tặng cẩm nang và kêu gọi đăng ký kênh.'
    ]
]

pc_rows = [
    [
        1,
        'Scene 01',
        '2D doodle animation: Cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon and black dress silhouette holds a cartoon magnifying glass and winks playfully at camera on clean pure white background. Three doodle items (clear plastic box, orange silicone mold, metal spoon) rotate around her with glowing yellow question marks. A funny cartoon 3-second countdown timer ticks above her head as she poses her riddle. Smooth camera push-in. Ms. Thuy\'s Voiceover (Vietnamese, playful energetic female narrator, pacing 10s): "Đố bạn một câu cực hóc búa: Giữa hộp nhựa, khuôn silicone và chiếc thìa ăn cơm — cái nào hải quan cho thông quan ngay, cái nào bị xích lại ở cảng?" Audio & SFX: Mystery riddle harp intro, playful ticking timer clock, cartoon ding pop, energetic playful bassline.',
        'Đố bạn một câu cực hóc búa: Giữa hộp nhựa, khuôn silicone và chiếc thìa ăn cơm — cái nào hải quan cho thông quan ngay, cái nào bị xích lại ở cảng?',
        10
    ],
    [
        2,
        'Scene 02',
        '2D doodle animation: Cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon and black dress pulls open a bright yellow doodle theatrical curtain on clean pure white background, revealing organized stacks of food packaging containers with glowing golden sparkles and official regulatory scrolls. Gentle smooth camera tracking. Ms. Thuy\'s Voiceover (Vietnamese, confident welcoming female expert, pacing 8s): "Hôm nay cùng Thúy vén màn toàn bộ tất tần tật quy định về kiểm soát hàng nhập khẩu đối với sản phẩm là dụng cụ bao bì chứa đựng nhé." Audio & SFX: Fabric curtain unroll swoosh, magical chime sparkle, uplifting rhythm.',
        'Hôm nay cùng Thúy vén màn toàn bộ tất tần tật quy định về kiểm soát hàng nhập khẩu đối với sản phẩm là dụng cụ bao bì chứa đựng nhé.',
        8
    ],
    [
        3,
        'Scene 03',
        '2D doodle animation: Cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon and black dress wags an index finger playfully with an intriguing smile on clean pure white background. An animated doodle lightbulb above her head suddenly flips upside down, and bold red text "SỰ THẬT NGƯỢC LẠI!" pops up with comic exclamation marks. Slight playful camera zoom-in. Ms. Thuy\'s Voiceover (Vietnamese, intriguing teasing expert tone, pacing 8s): "Và sẽ có những bạn tưởng đã biết rất rõ nhưng sự thật lại hoàn toàn ngược lại đấy nhé!" Audio & SFX: Playful finger tap boing, cartoon lightbulb switch pop, comic suspense string swell.',
        'Và sẽ có những bạn tưởng đã biết rất rõ nhưng sự thật lại hoàn toàn ngược lại đấy nhé!',
        8
    ],
    [
        4,
        'Scene 04',
        'Wide shot split-screen 2D doodle animation on clean pure white background: On left side, a cute container truck zooms smoothly past a green traffic light with a "1 NGÀY" badge. On right side, a giant cargo container is wrapped in heavy black cartoon chains and an oversized red padlock, surrounded by comic smoke and bold red text "KẸT CỨNG CẢ TUẦN!". Fast dynamic whip pan between both sides. Ms. Thuy\'s Voiceover (Vietnamese, confident female expert, pacing 8s): "Cùng là bao bì, có doanh nghiệp thông quan trong 1 ngày, nhưng có bên bị kẹt cứng cả tuần!" Audio & SFX: Fast truck whoosh, loud metal chain rattle, comic alert buzzer.',
        'Cùng là bao bì, có doanh nghiệp thông quan trong 1 ngày, nhưng có bên bị kẹt cứng cả tuần!',
        8
    ],
    [
        5,
        'Scene 05',
        'Close-up 2D doodle animation: Focus on an aseptic aluminum foil milk packaging pouch on clean pure white background. The packaging smoothly unfolds into a technical cross-section doodle showing aluminum barrier layers, while an official document with a blue checkmark slides into view. Smooth camera tilt down. Ms. Thuy\'s Voiceover (Vietnamese, analytical tone, pacing 8s): "Điển hình là lô 4 container túi màng nhôm vô trùng đóng sữa vừa qua. Doanh nghiệp đinh ninh có bản Tự công bố là xong!" Audio & SFX: Paper sliding swish, high-tech digital ping, soft background beat.',
        'Điển hình là lô 4 container túi màng nhôm vô trùng đóng sữa vừa qua. Doanh nghiệp đinh ninh có bản Tự công bố là xong!',
        8
    ],
    [
        6,
        'Scene 06',
        'Close-up 2D doodle animation: An official Circular 27 document on clean pure white background. Suddenly, a giant red cartoon rubber stamp slams down forcefully stamping "DỪNG THÔNG QUAN", sending comic dust puffs and shockwaves across the frame. Energetic camera shake on stamp impact. Ms. Thuy\'s Voiceover (Vietnamese, dramatic warning tone, pacing 8s): "Họ ngộ nhận Thông tư 27 xếp bao bì vào mức rủi ro trung bình, cho tới khi hải quan dừng hàng kiểm tra nhà nước!" Audio & SFX: Heavy rubber stamp thud impact, comic glass breaking cue, dramatic violin swell.',
        'Họ ngộ nhận Thông tư 27 xếp bao bì vào mức rủi ro trung bình, cho tới khi hải quan dừng hàng kiểm tra nhà nước!',
        8
    ],
    [
        7,
        'Scene 07',
        'Medium shot 2D doodle animation: Female stick figure Thuy with high ponytail and yellow ribbon stands confidently on clean pure white background, tapping a glowing golden gavel on a doodle podium. She waves her hand and official dispatches 22082 and 1913 float into the air with glowing emerald checkmarks. Smooth slow camera pan. Ms. Thuy\'s Voiceover (Vietnamese, reassuring confident expert, pacing 8s): "Thực tế, Văn bản 22082 và 1913 đã giải thích rõ. Đừng lo, hôm nay Thúy sẽ gỡ rối toàn diện cho bạn!" Audio & SFX: Wooden gavel tap, magical shimmer chime, upbeat confident rhythm.',
        'Thực tế, Văn bản 22082 và 1913 đã giải thích rõ. Đừng lo, hôm nay Thúy sẽ gỡ rối toàn diện cho bạn!',
        8
    ],
    [
        8,
        'Scene 08',
        'Wide shot 2D doodle animation: A hand-drawn whiteboard animation where a 3-tier legal pyramid is sketched out step-by-step on clean pure white background: Tier 1 slides in from bottom, Tier 2 stacks on top, Tier 3 crowns the peak. A cartoon golden key drops from above, unlocking sparkling golden stars around the pyramid. Smooth camera crane up following key. Ms. Thuy\'s Voiceover (Vietnamese, pedagogical teacher tone, pacing 8s): "Chìa khóa vàng nằm ở việc nắm vững Tháp 3 tầng quản lý pháp lý mà hầu hết mọi người hay nhầm lẫn!" Audio & SFX: Pencil sketching sounds, building blocks click-clack, golden key magical bell ding.',
        'Chìa khóa vàng nằm ở việc nắm vững Tháp 3 tầng quản lý pháp lý mà hầu hết mọi người hay nhầm lẫn!',
        8
    ],
    [
        9,
        'Scene 09',
        'Medium shot 2D doodle animation: Female stick figure Thuy with yellow ribbon pointer explains at a doodle customs gate on clean pure white background. An animated magnifying glass glides over the title "THÔNG TƯ 15", highlighting specific HS codes with bright yellow highlighter lines. Smooth slow drift left to right. Ms. Thuy\'s Voiceover (Vietnamese, precise instructional tone, pacing 8s): "Tầng 1 là Thủ tục Kiểm tra nhà nước tại Cửa khẩu. Chỉ những mặt hàng có mã HS trong Thông tư 15 mới bị kiểm tra!" Audio & SFX: Chalkboard pointer tap, highlighter squeak, gentle educational melody.',
        'Tầng 1 là Thủ tục Kiểm tra nhà nước tại Cửa khẩu. Chỉ những mặt hàng có mã HS trong Thông tư 15 mới bị kiểm tra!',
        8
    ],
    [
        10,
        'Scene 10',
        'Close-up 2D doodle animation: A transparent plastic PP container and a baby pacifier nipple appear on clean pure white background. A bold red cartoon dashed boundary box draws itself around them, followed by an animated red warning siren spinning and flashing with "BẮT BUỘC KIỂM TRA". Smooth camera push-in. Ms. Thuy\'s Voiceover (Vietnamese, stern regulatory tone, pacing 8s): "Đó là Bao bì nhựa như PP, PE, PET và Núm vú cao su trẻ em mã HS 4014.90.10. Thiếu đăng ký kiểm tra là dừng thông quan!" Audio & SFX: Warning siren chirp, red neon hum, barcode beep.',
        'Đó là Bao bì nhựa như PP, PE, PET và Núm vú cao su trẻ em mã HS 4014.90.10. Thiếu đăng ký kiểm tra là dừng thông quan!',
        8
    ],
    [
        11,
        'Scene 11',
        'Medium shot 2D doodle animation: Doodle kitchenware (stainless pot, ceramic bowl, glass cup) pops onto clean pure white background with bouncy cartoon sounds. A giant green shield stamps over them with an animated sparkling green checkmark and "THÔNG QUAN NGAY". Slight camera tilt up. Ms. Thuy\'s Voiceover (Vietnamese, upbeat cheerful tone, pacing 8s): "Ngược lại, Inox, Gốm sứ, Thủy tinh và Cao su người lớn được miễn kiểm tra cửa khẩu, hải quan cho thông quan ngay!" Audio & SFX: Cartoon popping dings, triumphant brass fanfare, green shield whoosh.',
        'Ngược lại, Inox, Gốm sứ, Thủy tinh và Cao su người lớn được miễn kiểm tra cửa khẩu, hải quan cho thông quan ngay!',
        8
    ],
    [
        12,
        'Scene 12',
        'Medium shot 2D doodle animation: Female stick figure Thuy with high ponytail folds a doodle dossier into a paper airplane on clean pure white background and throws it smoothly; the paper airplane flies in a dotted arc into the entrance of a cute doodle Food Safety Authority building. Tracking camera following paper airplane. Ms. Thuy\'s Voiceover (Vietnamese, professional procedural tone, pacing 8s): "Tầng 2 là Thủ tục Tự công bố theo Nghị định 15. Hồ sơ này nộp về Ban Quản lý An toàn thực phẩm để lưu hành hợp pháp!" Audio & SFX: Paper folding rustle, airplane whoosh glide, official mailbox drop thud.',
        'Tầng 2 là Thủ tục Tự công bố theo Nghị định 37. Hồ sơ này nộp về Ban Quản lý An toàn thực phẩm để lưu hành hợp pháp!',
        8
    ],
    [
        13,
        'Scene 13',
        'Close-up 2D doodle animation: Cartoon test tubes bubble happily with animated chemical steam on clean pure white background. A test tube pours blue liquid onto a test sheet, which turns emerald green and displays a gleaming golden seal "ĐẠT CHUẨN QCVN". Macro zoom into golden seal. Ms. Thuy\'s Voiceover (Vietnamese, authoritative compliance tone, pacing 8s): "Tầng 3 là Nghĩa vụ QCVN. Mọi sản phẩm đưa ra thị trường đều bắt buộc đạt chuẩn an toàn thôi nhiễm để phòng thủ hậu kiểm!" Audio & SFX: Liquid bubbling sound, chemical reaction sizzle, solid metallic seal stamp.',
        'Tầng 3 là Nghĩa vụ QCVN. Mọi sản phẩm đưa ra thị trường đều bắt buộc đạt chuẩn an toàn thôi nhiễm để phòng thủ hậu kiểm!',
        8
    ],
    [
        14,
        'Scene 14',
        'Close-up 2D doodle animation: Female stick figure Thuy leans forward on clean pure white background, holding a baby pacifier in one hand and a silicone baking mold in the other. A yellow warning badge between them pulsates with comic lightning bolts as she compares them side-by-side. Subtle camera push-in. Ms. Thuy\'s Voiceover (Vietnamese, intriguing storytelling tone, pacing 8s): "Bây giờ là điểm thắt cổ chai khiến 80% doanh nghiệp mất tiền: Cạm bẫy phân loại giữa Cao su trẻ em và Silicone người lớn!" Audio & SFX: Low tension cello drone, electrical spark zap, comic warning clock tick.',
        'Bây giờ là điểm thắt cổ chai khiến 80% doanh nghiệp mất tiền: Cạm bẫy phân loại giữa Cao su trẻ em và Silicone người lớn!',
        8
    ],
    [
        15,
        'Scene 15',
        'Close-up 2D doodle animation: Split screen on clean pure white background divides with a hand-drawn black line. On left, the silicone mold drives through an open green customs barrier. On right, a dossier opens up and a blue ink stamp firmly marks "BẮT BUỘC TỰ CÔNG BỐ". Whip pan from left to right. Ms. Thuy\'s Voiceover (Vietnamese, sharp analytical tone, pacing 8s): "Khuôn silicone người lớn được miễn cửa khẩu thông quan ngay, nhưng Mục 6 Thông tư 27 vẫn bắt buộc làm Tự công bố!" Audio & SFX: Quick barrier gate ding, rapid ink stamp thud, rhythmic percussion.',
        'Khuôn silicone người lớn được miễn cửa khẩu thông quan ngay, nhưng Mục 6 Thông tư 27 vẫn bắt buộc làm Tự công bố!',
        8
    ],
    [
        16,
        'Scene 16',
        'Close-up 2D doodle animation: An inox pan, ceramic bowl, and glass tumbler shine with cartoon twinkle stars on clean pure white background. A green doodle banner unravels underneath with the words "RỦI RO THẤP - MIỄN NỘP SỞ" while unnecessary paperwork flies away with wings. Smooth camera tracking. Ms. Thuy\'s Voiceover (Vietnamese, refreshing relief tone, pacing 8s): "Trong khi đó, Nồi Inox, Bát gốm sứ và Ly thủy tinh thuộc nhóm rủi ro thấp. Bạn không phải nộp hồ sơ Tự công bố lên Sở!" Audio & SFX: Twinkle sparkle chimes, relaxing harp chord, winged paper fluttering away.',
        'Trong khi đó, Nồi Inox, Bát gốm sứ và Ly thủy tinh thuộc nhóm rủi ro thấp. Bạn không phải nộp hồ sơ Tự công bố lên Sở!',
        8
    ],
    [
        17,
        'Scene 17',
        'Close-up 2D doodle animation: An internal binder labeled TCCS snaps shut on clean pure white background. A cartoon blue knight\'s shield drops in front of it with a solid metallic clang, symbolizing rock-solid legal defense against market surveillance inspectors. Slow camera tilt down. Ms. Thuy\'s Voiceover (Vietnamese, firm advisory tone, pacing 8s): "Doanh nghiệp chỉ cần tự công bố Tiêu chuẩn cơ sở nội bộ, lưu phiếu test QCVN để xuất trình khi Quản lý thị trường kiểm tra!" Audio & SFX: Binder snap click, heavy shield lock impact, confident deep bass.',
        'Doanh nghiệp chỉ cần tự công bố Tiêu chuẩn cơ sở nội bộ, lưu phiếu test QCVN để xuất trình khi Quản lý thị trường kiểm tra!',
        8
    ],
    [
        18,
        'Scene 18',
        'Medium shot 2D doodle animation: Female stick figure Thuy with ponytail cheerfully punches buttons on a giant cartoon calculator. The calculator screen dings with "TIẾT KIỆM HÀNG TRIỆU ĐỒNG" and confetti/coins burst into the air around her on clean pure white background. Dynamic zoom out then in. Ms. Thuy\'s Voiceover (Vietnamese, enthusiastic inspiring tone, pacing 8s): "Hiểu rõ điểm này, bạn tiết kiệm ngay hàng triệu đồng chi phí làm thủ tục không cần thiết cho hàng Inox và Thủy tinh!" Audio & SFX: Calculator cash register cha-ching, cheerful confetti party horn, bouncy acoustic guitar.',
        'Hiểu rõ điểm này, bạn tiết kiệm ngay hàng triệu đồng chi phí làm thủ tục không cần thiết cho hàng Inox và Thủy tinh!',
        8
    ],
    [
        19,
        'Scene 19',
        'Medium shot 2D doodle animation: A large food storage box appears on left, while a spoon and chopsticks tap against each other on right on clean pure white background. A glowing yellow electric line splits the screen in half with animated question marks popping up. Subtle horizontal camera pan. Ms. Thuy\'s Voiceover (Vietnamese, intriguing revelation tone, pacing 8s): "Điểm huyệt thứ hai là phân biệt giữa Dụng cụ CHỨA ĐỰNG tiếp xúc trực tiếp và Dụng cụ TIẾP XÚC nhưng KHÔNG CHỨA ĐỰNG!" Audio & SFX: Cutlery tapping clink, electrical crackle zap, curious pizzicato strings.',
        'Điểm huyệt thứ hai là phân biệt giữa Dụng cụ CHỨA ĐỰNG tiếp xúc trực tiếp và Dụng cụ TIẾP XÚC nhưng KHÔNG CHỨA ĐỰNG!',
        8
    ],
    [
        20,
        'Scene 20',
        'Close-up 2D doodle animation: An official government scroll unrolls smoothly on clean pure white background, revealing the glowing header "CÔNG VĂN 4372/BYT-ATTP". An official red seal stamps with a golden sparkle effect. Slow camera push-in. Ms. Thuy\'s Voiceover (Vietnamese, historic authoritative tone, pacing 8s): "Căn cứ then chốt nằm tại Công văn 4372 năm 2019 của Bộ Y tế trả lời chính thức cho công ty Aeon Việt Nam!" Audio & SFX: Parchment unroll swoosh, heavy royal seal stamp, golden magical shimmer.',
        'Căn cứ then chốt nằm tại Công văn 4372 năm 2019 của Bộ Y tế trả lời chính thức cho công ty Aeon Việt Nam!',
        8
    ],
    [
        21,
        'Scene 21',
        'Close-up 2D doodle animation: Tableware playfully animates on clean pure white background: chopsticks click together, a spoon scoops air, and cooking tongs snap open and shut. An animated text tag "DỤNG CỤ CƠ HỌC - KHÔNG CHỨA ĐỰNG" bounces into view. Slight orbit camera pan. Ms. Thuy\'s Voiceover (Vietnamese, clear definitive tone, pacing 8s): "Bộ Y tế nêu rõ: Đũa, Thìa, Dĩa, Dao gọt hoa quả, Găng tay nấu ăn và Kẹp gắp KHÔNG PHẢI dụng cụ chứa đựng!" Audio & SFX: Cutlery clinking rhythm, cartoon bounce boing, energetic shaker groove.',
        'Bộ Y tế nêu rõ: Đũa, Thìa, Dĩa, Dao gọt hoa quả, Găng tay nấu ăn và Kẹp gắp KHÔNG PHẢI dụng cụ chứa đựng!',
        8
    ],
    [
        22,
        'Scene 22',
        'Medium shot 2D doodle animation: Female stick figure Thuy raises both hands triumphantly; two giant emerald green shields pop onto screen beside her with checkmarks on clean pure white background, while a cartoon red "CẤM" sign shatters into dust. Smooth center camera zoom. Ms. Thuy\'s Voiceover (Vietnamese, triumphant expert tone, pacing 8s): "Mặt hàng này HOÀN TOÀN MIỄN TỰ CÔNG BỐ, Sở từ chối tiếp nhận, và MIỄN LUÔN kiểm tra chuyên ngành tại cửa khẩu!" Audio & SFX: Double shield pop impact, glass shatter whoosh, victorious brass crescendo.',
        'Mặt hàng này HOÀN TOÀN MIỄN TỰ CÔNG BỐ, Sở từ chối tiếp nhận, và MIỄN LUÔN kiểm tra chuyên ngành tại cửa khẩu!',
        8
    ],
    [
        23,
        'Scene 23',
        'Wide shot 2D doodle animation: A cute cartoon container truck drives rapidly from right to left through an open green customs barrier on clean pure white background, leaving comic motion puffs and green speed lines, as a green traffic light flashes "PASSED". Tracking camera with truck. Ms. Thuy\'s Voiceover (Vietnamese, breezy cheerful tone, pacing 8s): "Bạn nhập khẩu chúng như hàng hóa thông thường, không mất một đồng phí kiểm tra chuyên ngành và thông quan trong ngày!" Audio & SFX: Truck horn honk, green traffic light chime, rapid wheel whoosh.',
        'Bạn nhập khẩu chúng như hàng hóa thông thường, không mất một đồng phí kiểm tra chuyên ngành và thông quan trong ngày!',
        8
    ],
    [
        24,
        'Scene 24',
        'Close-up 2D doodle animation: Calendar pages flip rapidly like wind on clean pure white background. A cartoon stopwatch ticks frantically with comic sweat drops flying off it, freezing at the number "10 - 12 NGÀY". Rapid camera push-in to clock face. Ms. Thuy\'s Voiceover (Vietnamese, urgent tactical warning, pacing 8s): "Chiến lược then chốt tiếp theo là kiểm soát thời gian. Thử nghiệm tại phòng lab mất từ 10 đến 12 ngày làm việc!" Audio & SFX: Wind blowing pages flap-flap, frantic ticking stopwatch, alarm bell ring.',
        'Chiến lược then chốt tiếp theo là kiểm soát thời gian. Thử nghiệm tại phòng lab mất từ 10 đến 12 ngày làm việc!',
        8
    ],
    [
        25,
        'Scene 25',
        'Close-up 2D doodle animation: An oversized red billing invoice slams down onto a crying stick-figure container on clean pure white background. Red alert sirens spin in the upper corners and bold numbers "15 - 30 TRIỆU" pulse menacingly. Dramatic camera shake on slam. Ms. Thuy\'s Voiceover (Vietnamese, alarming cost warning tone, pacing 8s): "Nếu chờ tàu cập cảng mới lấy mẫu gửi test, tiền phạt lưu cont lưu bãi sẽ lên tới 15 đến 30 triệu mỗi container!" Audio & SFX: Heavy invoice slam impact, crying cartoon whimpers, danger alert buzzer.',
        'Nếu chờ tàu cập cảng mới lấy mẫu gửi test, tiền phạt lưu cont lưu bãi sẽ lên tới 15 đến 30 triệu mỗi container!',
        8
    ],
    [
        26,
        'Scene 26',
        'Medium shot 2D doodle animation: Female stick figure Thuy presents an express courier box with cartoon wings on clean pure white background. The winged parcel flaps its wings and zooms along a dotted timeline curve labeled "TRƯỚC 25 NGÀY", leaving yellow stars in its wake. Smooth pan following timeline. Ms. Thuy\'s Voiceover (Vietnamese, tactical strategist tone, pacing 8s): "Kinh nghiệm thực chiến của Thúy: Hãy gửi chuyển phát nhanh 2 đến 3 mẫu từ xưởng về trước ngày tàu chạy 25 ngày!" Audio & SFX: Fluttering wings swoosh, jet engine zoom, magical starlight sparkles.',
        'Kinh nghiệm thực chiến của Thúy: Hãy gửi chuyển phát nhanh 2 đến 3 mẫu từ xưởng về trước ngày tàu chạy 25 ngày!',
        8
    ],
    [
        27,
        'Scene 27',
        'Close-up 2D doodle animation: On a doodle computer screen on clean pure white background, a mouse cursor clicks an animated button "NỘP HỒ SƠ". Instantly, a green progress bar loads to 100% and a giant green checkmark pops up with "THÔNG QUAN 3 NGÀY". Slight camera zoom into screen. Ms. Thuy\'s Voiceover (Vietnamese, smooth efficient tone, pacing 8s): "Khi tàu cập cảng, Tự công bố đã sẵn sàng. Bạn đăng ký kiểm tra nhà nước và thông quan chỉ mất đúng 3 ngày!" Audio & SFX: Mouse click clack, digital progress bar swoosh, success confirmation chime.',
        'Khi tàu cập cảng, Tự công bố đã sẵn sàng. Bạn đăng ký kiểm tra nhà nước và thông quan chỉ mất đúng 3 ngày!',
        8
    ],
    [
        28,
        'Scene 28',
        'Wide shot 2D doodle animation: A container truck reverses safely into a doodle warehouse bay on clean pure white background. A giant pair of red cartoon scissors snips a chain of storage fees in two, making coin bags drop safely into a vault labeled "0Đ PHÍ LƯU BÃI". Camera crane down to ground. Ms. Thuy\'s Voiceover (Vietnamese, practical problem solver tone, pacing 8s): "Trường hợp gấp, giải pháp là làm thủ tục xin đưa hàng về kho riêng bảo quản, triệt tiêu hoàn toàn phí lưu bãi!" Audio & SFX: Truck backing beep, crisp scissors snip, heavy safe door locking click.',
        'Trường hợp gấp, giải pháp là làm thủ tục xin đưa hàng về kho riêng bảo quản, triệt tiêu hoàn toàn phí lưu bãi!',
        8
    ],
    [
        29,
        'Scene 29',
        'Medium shot 2D doodle animation: Female stick figure Thuy proudly holds up a certificate on clean pure white background. The infinity symbol "∞" on the document glows brightly, rotating and pulsating with cartoon golden sparkles as containers in background pass freely. Slow zoom in on infinity symbol. Ms. Thuy\'s Voiceover (Vietnamese, high-value insight tone, pacing 8s): "Đặc biệt: Bản Tự công bố có hiệu lực vô thời hạn cho từng model. Từ lô thứ hai, bạn không tốn thêm một đồng nào!" Audio & SFX: Shimmering infinity rotation sound, golden bell toll, celebratory acoustic strum.',
        'Đặc biệt: Bản Tự công bố có hiệu lực vô thời hạn cho từng model. Từ lô thứ hai, bạn không tốn thêm một đồng nào!',
        8
    ],
    [
        30,
        'Scene 30',
        'Close-up 2D doodle animation: A doodle test certificate sits beside an animated hourglass on clean pure white background. The blue sand falls down to the last grain, and an animated red warning clock rings with "LƯU Ý: HẠN 12 THÁNG". Push-in to red stamp. Ms. Thuy\'s Voiceover (Vietnamese, critical compliance reminder, pacing 8s): "Chỉ cần lưu ý: Phiếu kiểm nghiệm có giá trị đúng 12 tháng khi nộp hồ sơ. Đừng để quá hạn kẻo phải làm lại từ đầu!" Audio & SFX: Hourglass trickling sand, ringing alarm clock, urgent warning ding.',
        'Chỉ cần lưu ý: Phiếu kiểm nghiệm có giá trị đúng 12 tháng khi nộp hồ sơ. Đừng để quá hạn kẻo phải làm lại từ đầu!',
        8
    ],
    [
        31,
        'Scene 31',
        'Close-up 2D doodle animation: Female stick figure Thuy looks directly at audience with an inspiring, thoughtful smile on clean pure white background. She holds up a golden doodle lightbulb that flashes on brightly, while a solid green legal fortress barrier rises behind her, shielding a passing container ship. Slow cinematic push-in. Ms. Thuy\'s Voiceover (Vietnamese, inspiring leader wisdom, pacing 8s): "Logistics thực chiến không phải chờ sự cố rồi chữa cháy, mà là xây hàng rào pháp lý trước khi container rời xưởng!" Audio & SFX: Warm lightbulb switch click, deep inspiring cello swell, resonant golden bell.',
        'Logistics thực chiến không phải chờ sự cố rồi chữa cháy, mà là xây hàng rào pháp lý trước khi container rời xưởng!',
        8
    ],
    [
        32,
        'Scene 32',
        'Medium shot 2D doodle animation: Female stick figure Thuy waves happily to audience on clean pure white background. Beside her, an animated red Subscribe button is clicked by a cartoon hand, a golden bell rings with soundwaves, and a gift book labeled "CẨM NANG BAO BÌ" drops into an open comment box. Smooth center hold. Ms. Thuy\'s Voiceover (Vietnamese, warm welcoming call-to-action, pacing 8s): "Hãy bấm Đăng ký kênh và bình luận BAO BÌ, Thúy sẽ gửi tặng bạn Cẩm nang nghiệp vụ nhập khẩu bao bì thực phẩm độc quyền nhé!" Audio & SFX: Mouse click pop, bell ringing ding-dong, cheerful acoustic outro theme.',
        'Hãy bấm Đăng ký kênh và bình luận BAO BÌ, Thúy sẽ gửi tặng bạn Cẩm nang nghiệp vụ nhập khẩu bao bì thực phẩm độc quyền nhé!',
        8
    ]
]

# Write to ANH.XLSX
file_anh = r'C:\Users\thuyn\Downloads\ảnh.xlsx'
wb_anh = openpyxl.Workbook()
ws_anh = wb_anh.active
ws_anh.title = 'TAO_ANH_HANG_LOAT'

headers_anh = ['STT', 'Tiêu đề', 'Prompt tạo ảnh (Image Prompt)', 'Góc quay (Shot Type)', 'Ghi chú']
ws_anh.append(headers_anh)
for r in anh_rows:
    ws_anh.append(r)

header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

for col_num in range(1, 6):
    cell = ws_anh.cell(row=1, column=col_num)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align

for row_num in range(2, len(anh_rows) + 2):
    for col_num in range(1, 6):
        cell = ws_anh.cell(row=row_num, column=col_num)
        cell.font = Font(name='Arial', size=10)
        cell.border = thin_border
        if col_num in [1, 4]:
            cell.alignment = center_align
        else:
            cell.alignment = left_align

ws_anh.column_dimensions['A'].width = 8
ws_anh.column_dimensions['B'].width = 32
ws_anh.column_dimensions['C'].width = 85
ws_anh.column_dimensions['D'].width = 26
ws_anh.column_dimensions['E'].width = 40

wb_anh.save(file_anh)
print(f'Successfully overwritten {file_anh} with {len(anh_rows)} rows.')

# Write to PHAN CANH.XLSX
file_pc = r'C:\Users\thuyn\Downloads\phân cảnh.xlsx'
wb_pc = openpyxl.Workbook()
ws_pc = wb_pc.active
ws_pc.title = 'KICH_BAN_FLOW'

headers_pc = ['STT', 'Tiêu đề', 'Mô tả phân cảnh (Prompt)', 'Lời thoại (Dialogue)', 'Thời lượng (giây)']
ws_pc.append(headers_pc)
for r in pc_rows:
    ws_pc.append(r)

for col_num in range(1, 6):
    cell = ws_pc.cell(row=1, column=col_num)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align

for row_num in range(2, len(pc_rows) + 2):
    for col_num in range(1, 6):
        cell = ws_pc.cell(row=row_num, column=col_num)
        cell.font = Font(name='Arial', size=10)
        cell.border = thin_border
        if col_num in [1, 5]:
            cell.alignment = center_align
        else:
            cell.alignment = left_align

ws_pc.column_dimensions['A'].width = 8
ws_pc.column_dimensions['B'].width = 15
ws_pc.column_dimensions['C'].width = 95
ws_pc.column_dimensions['D'].width = 65
ws_pc.column_dimensions['E'].width = 18

wb_pc.save(file_pc)
print(f'Successfully overwritten {file_pc} with {len(pc_rows)} rows.')
