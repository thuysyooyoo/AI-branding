# HƯỚNG DẪN HÌNH ẢNH VÀ FLOW

## Khóa nhất quán toàn video

Thay `[CHARACTER_REFERENCE]` bằng ảnh nhân vật người dùng cung cấp. Với mọi ảnh và mọi cảnh video, giữ nguyên:

- Khuôn mặt, kiểu tóc, tỷ lệ cơ thể, độ tuổi biểu kiến và trang phục của nhân vật tham chiếu.
- Phong cách minh họa của ảnh tham chiếu; không chuyển sang ảnh thật.
- Tỷ lệ 16:9, bố cục cinematic, vùng trống an toàn phía dưới để đặt phụ đề.
- Bảng màu: xanh teal lạnh cho cảng và hệ thống dữ liệu; cam ấm cho hàng hóa, tài liệu và các điểm hành động.
- Chuyển động vừa phải, có quán tính tự nhiên; tránh morphing khuôn mặt, tay và vật thể.
- Không tạo chữ nhỏ trong ảnh, không logo thương hiệu, không watermark. Chữ sẽ được thêm ở khâu dựng.

## Prompt khóa nhân vật và phong cách

`Use the uploaded [CHARACTER_REFERENCE] as the exact identity and style reference. Preserve the same face, hairstyle, age, body proportions, outfit, color palette, line quality and illustration medium in every frame. Hand-painted editorial 2.5D animation, textured gouache and colored-pencil details, cinematic lighting, subtle paper grain, teal-blue European logistics environments with warm amber accents, coherent scale and geography, 16:9, 4K, clean lower-third space for Vietnamese subtitles. No readable text, no logos, no watermark, no extra fingers, no facial drift, no outfit changes, no photorealism.`

## Cách dùng khung hình trong Flow

1. Tạo 31 khung hình K00 đến K30 theo bảng storyboard.
2. Cảnh S01 dùng K00 làm khung đầu và K01 làm khung cuối; S02 dùng K01 và K02; tiếp tục như vậy đến S30 dùng K29 và K30.
3. Với mỗi cảnh, đưa hai khung đầu–cuối cùng prompt video 10 giây vào Flow.
4. Giữ cùng seed hoặc cùng character reference nếu công cụ hỗ trợ.
5. Không yêu cầu AI tạo chữ trong khung hình. Thêm tiêu đề, số liệu và phụ đề sau khi tạo video.
6. Dùng cut mềm hoặc match cut 6–10 khung hình giữa hai cảnh; không dùng transition quá mạnh làm đứt mạch hình ảnh.

## Nhịp dựng đề xuất

- 00:00–01:20: bí ẩn, nhịp vừa, tiếng cảng và âm báo dữ liệu.
- 01:20–02:10: rõ ràng, mang tính giải thích, motion graphic dữ liệu.
- 02:10–03:00: timeline nhanh, điểm nhấn bằng ngày tháng.
- 03:00–04:40: thực hành, hình ảnh checklist và chuỗi cung ứng.
- 04:40–05:00: sáng hơn, kết thúc chủ động và tích cực.

## Negative prompt dùng chung

`photorealistic, 3D plastic look, inconsistent face, different hairstyle, wardrobe change, deformed hands, extra fingers, duplicate character, unreadable typography, random letters, logos, flags used inaccurately, shaky camera, abrupt teleportation, object morphing, flicker, over-saturated neon, low resolution, watermark`

## Tình trạng tạo ảnh

Bộ prompt đã được thiết kế để tạo các khung K00–K30 có tính nối tiếp. Cần ảnh nhân vật gốc (và ảnh phong cách cảnh vật nếu tách riêng) trước khi tạo ảnh thật để tránh sai khuôn mặt, trang phục và nét vẽ.

