"""
================================================================================
SCRIPT: daily_notes_reminder.py
MỤC TIÊU:
1. Nhắc nhở Chị Thúy cập nhật các câu chuyện hàng ngày trong công việc, cuộc sống
   vào D:\\AI AGENT THUY\\AI branding\\scratch\\daily_notes.md theo đúng cấu trúc của file.
2. Tự động kiểm tra xem ngày hôm nay đã có ghi chép nhật ký thực tế chưa.
   - Nếu chưa có tiêu đề ngày hôm nay, tự động chèn khung mẫu (template) chuẩn chỉnh lên đầu file.
3. Soạn tin nhắn nhắc nhở thân thiện, súc tích với 4 câu hỏi định hướng theo chuẩn PBAN.
4. Gửi trực tiếp vào Zalo 'Cloud của tôi' và hiển thị thông báo màn hình Windows.
================================================================================
"""

import os
import sys
import argparse
from datetime import datetime
import subprocess

# Đường dẫn thư mục gốc
BASE_DIR = r"D:\AI AGENT THUY\AI branding"
SCRATCH_DIR = os.path.join(BASE_DIR, "scratch")
DAILY_NOTES_PATH = os.path.join(SCRATCH_DIR, "daily_notes.md")
ZALO_PROJECT_DIR = r"D:\AI AGENT THUY\ZALO REPORT"

def get_today_str():
    return datetime.now().strftime("%d/%m/%Y")

def check_and_prepare_daily_notes(today_str):
    """
    Kiểm tra daily_notes.md:
    - Nếu chưa có '## Ngày dd/mm/yyyy', tự động chèn khung mẫu lên đầu danh sách ngày.
    - Trả về (has_today_header, is_already_filled, message)
    """
    if not os.path.exists(DAILY_NOTES_PATH):
        os.makedirs(os.path.dirname(DAILY_NOTES_PATH), exist_ok=True)
        initial_content = (
            "# NHẬT KÝ CÔNG VIỆC HẰNG NGÀY (DAILY NOTES)\n"
            "> 💡 **Hướng dẫn:** Chị Thúy chỉ cần ghi lại ngắn gọn những câu chuyện thực tế diễn ra trong ngày tại Eureka (sự cố, bài học quản trị, đàm phán, thành công/thất bại...). Không cần trau chuốt câu từ. Agent 01 sẽ tự đọc tệp này để phân tích và phác thảo ý tưởng.\n\n"
            "---\n\n"
        )
        with open(DAILY_NOTES_PATH, "w", encoding="utf-8") as f:
            f.write(initial_content)

    with open(DAILY_NOTES_PATH, "r", encoding="utf-8") as f:
        content = f.read()

    header_today = f"## Ngày {today_str}"
    has_today_header = header_today in content

    # Template mẫu chuẩn mực cho câu chuyện hàng ngày của Chị Thúy
    story_template = (
        f"### Câu chuyện 1: [Tiêu đề câu chuyện trong công việc / cuộc sống hôm nay]\n"
        f"- **Sự kiện / Câu chuyện nổi bật trong ngày:**\n"
        f"  - [Mô tả chi tiết những gì diễn ra: sự cố logistics, đàm phán nhà cung cấp, phản hồi khách hàng, nhân sự, hoặc một khoảnh khắc đời sống...]\n"
        f"- **Bài học rút ra hoặc suy ngẫm cá nhân (Góc nhìn Eureka):**\n"
        f"  - [Bài học quản trị, tư duy lãnh đạo Giáp Mộc - vững chãi che chở & Liêm Trinh - thẳng thắn liêm chính, chiêm nghiệm bản thân...]\n"
        f"- **Thông tin liên quan khác (nếu có):**\n"
        f"  - [Hành động tiếp theo, ảnh chụp chứng từ, tin nhắn đối chất...]\n\n"
        f"---\n\n"
    )
    sample_template = f"{header_today}\n\n" + story_template

    if not has_today_header:
        # Chèn lên đầu, ngay sau phần Header của file
        marker = "---\n\n"
        if marker in content:
            parts = content.split(marker, 1)
            updated_content = parts[0] + marker + sample_template + parts[1]
        else:
            updated_content = content + "\n\n" + sample_template

        with open(DAILY_NOTES_PATH, "w", encoding="utf-8") as f:
            f.write(updated_content)
        print(f"[+] Đã tự động chèn khung mẫu cho Ngày {today_str} vào daily_notes.md")
        return True, False, "Đã tạo khung mẫu mới cho ngày hôm nay."
    else:
        # Đã có header Ngày hôm nay (có thể do script cào FB tạo trước)
        today_section = content.split(header_today)[1].split("## Ngày ")[0]
        
        # Nếu chưa có mục câu chuyện cá nhân trong ngày hôm nay, chèn thêm vào ngay dưới header ngày
        if "### Câu chuyện" not in today_section:
            insert_marker = f"{header_today}\n\n"
            parts = content.split(insert_marker, 1)
            updated_content = parts[0] + insert_marker + story_template + parts[1]
            with open(DAILY_NOTES_PATH, "w", encoding="utf-8") as f:
                f.write(updated_content)
            print(f"[+] Đã chèn thêm khung mẫu câu chuyện cá nhân vào Ngày {today_str}")
            today_section = story_template + today_section

        # Phân biệt câu chuyện cá nhân với Case Study / Hot News cào tự động
        has_real_story = False
        for block in today_section.split("### "):
            if not block.strip():
                continue
            header_line = block.splitlines()[0].strip()
            # Bỏ qua các tin tức / case study cào tự động
            if header_line.startswith("Case Study:") or header_line.startswith("Hot News:"):
                continue
            if "[Tiêu đề" in header_line:
                continue
            # Kiểm tra nội dung chi tiết bên dưới
            for line in block.splitlines()[1:]:
                line_s = line.strip()
                if line_s.startswith("- ") and not line_s.startswith("- [") and len(line_s) > 10:
                    has_real_story = True
                    break
            if has_real_story:
                break
        
        return True, has_real_story, "Đã có mục ngày hôm nay."

def compose_reminder_message(today_str, is_already_filled=False):
    """
    Soạn nội dung tin nhắn nhắc nhở thân thiện và chuyên nghiệp.
    """
    if is_already_filled:
        msg = (
            f"🌟 [NHẬT KÝ DAILY NOTES {today_str}]\n"
            f"Chào Chị Thúy! Hệ thống ghi nhận chị đã cập nhật nhật ký cho ngày hôm nay ({today_str}).\n\n"
            f"Nếu trong chiều tối nay có thêm sự kiện, bài học hoặc cảm xúc mới nào phát sinh, chị có thể nhắn bổ sung để AI Agent tiếp tục cập nhật nhé!\n"
            f"📂 File lưu: D:\\AI AGENT THUY\\AI branding\\scratch\\daily_notes.md"
        )
    else:
        msg = (
            f"🌟 [NHẮC NHỞ HÀNG NGÀY] CẬP NHẬT NHẬT KÝ CÔNG VIỆC & CUỘC SỐNG\n"
            f"🗓 Ngày: {today_str}\n\n"
            f"Chào Chị Thúy! Một ngày làm việc nữa đã trôi qua, chị dành 2-3 phút ghi lại câu chuyện thực tế hoặc chiêm nghiệm hôm nay nhé.\n\n"
            f"💡 Chị chỉ cần gõ vắn tắt (hoặc voice-to-text) theo 4 nội dung chuẩn của Daily Notes:\n\n"
            f"1️⃣ TIÊU ĐỀ: Tên sự việc / tình huống trong ngày (Ví dụ: Đàm phán cước tàu, Xử lý khiếu nại khách hàng, Nhân sự mới, Chiêm nghiệm gia đình...)\n\n"
            f"2️⃣ SỰ KIỆN NỔI BẬT: Đã xảy ra chuyện gì? Ai tham gia? Diễn biến & nút thắt thực tế là gì? (Viết mộc mạc, không cần trau chuốt câu từ).\n\n"
            f"3️⃣ BÀI HỌC & SUY NGẪM (Góc nhìn Eureka): Chị đúc kết được gì về quản trị, lãnh đạo, con người hay quy trình? (Kết hợp tinh thần Giáp Mộc vững chãi & Liêm Trinh minh bạch).\n\n"
            f"4️⃣ THÔNG TIN LIÊN QUAN: Có hình ảnh chứng từ, tin nhắn đối chất hay việc cần làm tiếp theo không?\n\n"
            f"📲 Chị có thể nhắn trực tiếp lại tin nhắn này (hoặc mở file d:\\AI AGENT THUY\\AI branding\\scratch\\daily_notes.md). Đội ngũ AI Agent (01 & 03) sẽ tự động bóc tách thành chất liệu xây dựng thương hiệu cá nhân!\n"
        )
    return msg

def send_zalo_reminder(message_text):
    """
    Gửi tin nhắn qua Zalo Bot (Cloud của tôi)
    """
    try:
        sys.path.insert(0, ZALO_PROJECT_DIR)
        from core.zalo_sender import send_zalo_message
        print("[*] Đang gửi tin nhắn nhắc nhở qua Zalo Bot...")
        success = send_zalo_message(message_text, target_chat="Cloud của tôi")
        return success
    except Exception as e:
        print(f"[!] Lỗi khi gửi Zalo: {e}")
        return False

def show_desktop_popup(today_str):
    """
    Hiển thị thông báo màn hình Windows (tự động đóng sau 15 giây)
    """
    try:
        popup_text = f"Chị Thúy ơi, đừng quên ghi lại câu chuyện công việc & cuộc sống hôm nay ({today_str}) vào Daily Notes nhé!"
        popup_title = "Eureka Daily Notes Reminder"
        ps_cmd = f"$wshell = New-Object -ComObject Wscript.Shell; $wshell.Popup('{popup_text}', 15, '{popup_title}', 64)"
        subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], check=False)
        return True
    except Exception as e:
        print(f"[!] Lỗi hiển thị thông báo màn hình: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Eureka Daily Notes Reminder")
    parser.add_argument("--no-zalo", action="store_true", help="Không gửi tin nhắn qua Zalo")
    parser.add_argument("--no-popup", action="store_true", help="Không hiện thông báo màn hình")
    parser.add_argument("--force", action="store_true", help="Bắt buộc gửi nhắc nhở dù đã ghi")
    parser.add_argument("--check-only", action="store_true", help="Chỉ kiểm tra trạng thái file")
    args = parser.parse_args()

    today_str = get_today_str()
    print("=" * 70)
    print(f"⏰ EUREKA DAILY NOTES REMINDER - NGÀY: {today_str}")
    print("=" * 70)

    has_header, is_filled, status_msg = check_and_prepare_daily_notes(today_str)
    print(f"[*] Trạng thái file: {status_msg}")
    print(f"[*] Đã điền câu chuyện thực tế: {'Có' if is_filled else 'Chưa'}")

    if args.check_only:
        return

    if is_filled and not args.force:
        print("[*] Hôm nay chị Thúy đã cập nhật câu chuyện. Bỏ qua gửi nhắc nhở (dùng --force nếu muốn gửi).")
        return

    msg = compose_reminder_message(today_str, is_already_filled=is_filled)
    print("\n" + "-" * 70)
    print("NỘI DUNG NHẮC NHỞ:")
    print("-" * 70)
    print(msg)
    print("-" * 70 + "\n")

    # 1. Hiện thông báo màn hình
    if not args.no_popup:
        print("[*] Đang hiển thị thông báo màn hình Windows...")
        show_desktop_popup(today_str)

    # 2. Gửi Zalo Bot
    if not args.no_zalo:
        zalo_success = send_zalo_reminder(msg)
        if zalo_success:
            print("✅ Đã gửi nhắc nhở vào Zalo 'Cloud của tôi' thành công!")
        else:
            print("⚠️ Chưa gửi được Zalo (vui lòng kiểm tra lại phiên đăng nhập Zalo).")
    else:
        print("[*] Đã bỏ qua gửi Zalo (--no-zalo).")

    print("\n[+] HOÀN TẤT QUY TRÌNH NHẮC NHỞ DAILY NOTES!")

if __name__ == "__main__":
    main()
