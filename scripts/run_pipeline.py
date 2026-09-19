# -*- coding: utf-8 -*-
import os
import sys
import io
import json
import argparse
import urllib.request
import urllib.parse

# Đảm bảo in tiếng Việt trên console không bị lỗi
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

WORKSPACE_DIR = r"d:\AI AGENT THUY\AI branding"
GEMINI_KEY_ENV = "GEMINI_API_KEY"

def call_gemini(system_instruction, user_content, api_key):
    """Gọi Gemini 1.5 Flash API bằng thư viện chuẩn urllib của Python (Zero-dependency)"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
    headers = {"Content-Type": "application/json"}
    
    # Payload cấu trúc v1beta hỗ trợ systemInstruction
    payload = {
        "contents": [
            {
                "parts": [
                    {"text": user_content}
                ]
            }
        ],
        "systemInstruction": {
            "parts": [
                {"text": system_instruction}
            ]
        },
        "generationConfig": {
            "temperature": 0.3,
            "maxOutputTokens": 8192
        }
    }
    
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=60) as response:
            res_data = json.loads(response.read().decode("utf-8"))
            # Trích xuất text từ response
            text = res_data["candidates"][0]["content"]["parts"][0]["text"]
            return text
    except Exception as e:
        print(f"Lỗi kết nối Gemini API: {e}")
        if hasattr(e, 'read'):
            print(e.read().decode("utf-8"))
        return None

def read_file(path):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return ""

def write_file(path, content):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

def run_pban_pipeline(topic, run_llm, api_key):
    print("="*60)
    print("🚀 BẮT ĐẦU VẬN HÀNH PIPELINE TÁC TỰ THƯƠNG HIỆU PBAN")
    print("="*60)
    
    # --- Cập nhật Content Pillars ---
    print("\n📦 Đang tự động dọn dẹp và làm dày kho ý tưởng Content Pillars...")
    if run_llm:
        key_arg = f" --key {api_key}" if api_key else ""
        os.system(f"python \"{os.path.join(WORKSPACE_DIR, 'scripts', 'update_content_pillars.py')}\"{key_arg}")
    else:
        os.system(f"python \"{os.path.join(WORKSPACE_DIR, 'scripts', 'update_content_pillars.py')}\"")
        
    # 0. Đọc tri thức và prompt nền tảng dùng chung
    base_prompt = read_file(os.path.join(WORKSPACE_DIR, "agents", "_shared", "system_prompt_base.md"))
    
    # Đọc input từ daily notes nếu không truyền topic
    if not topic:
        daily_notes_path = os.path.join(WORKSPACE_DIR, "scratch", "daily_notes.md")
        if os.path.exists(daily_notes_path):
            print(f"Đang đọc nội dung câu chuyện thô từ: {daily_notes_path}")
            topic = read_file(daily_notes_path)
        else:
            topic = "Ứng dụng AI tối ưu quy trình hải quan và HS Code."
            
    print(f"Chủ đề đầu vào: {topic[:150]}...")
    
    # Đọc prompt của từng Agent
    agent1_prompt = read_file(os.path.join(WORKSPACE_DIR, "agents", "01_market_researcher.md"))
    agent2_prompt = read_file(os.path.join(WORKSPACE_DIR, "agents", "02_policy_seo_consultant.md"))
    agent3_prompt = read_file(os.path.join(WORKSPACE_DIR, "agents", "03_creative_director.md"))
    agent4_prompt = read_file(os.path.join(WORKSPACE_DIR, "agents", "04_content_planner.md"))
    agent5_prompt = read_file(os.path.join(WORKSPACE_DIR, "agents", "05_copywriter.md"))
    agent6_prompt = read_file(os.path.join(WORKSPACE_DIR, "agents", "06_scriptwriter.md"))
    agent8_prompt = read_file(os.path.join(WORKSPACE_DIR, "agents", "08_fact_checker.md"))
    
    if not run_llm:
        print("\n💡 Chế độ chạy thủ công (Manual Mode): Hệ thống sẽ đóng gói các file chỉ dẫn Prompt để bạn copy-paste.")
        # Tạo file tổng hợp các bước thủ công
        instructions = f"""# 📑 BÁO CÁO CÔNG VIỆC PIPELINE PBAN (BÁN TỰ ĐỘNG)
Chủ đề đầu vào: {topic}

Hãy sao chép các prompt dưới đây và dán vào cửa sổ chat AI của bạn để chạy lần lượt các Agent:

## Bước 1: Agent 01 (Market Researcher)
**System Instruction:**
{base_prompt}
---
{agent1_prompt}

**User Input:**
Nghiên cứu chủ đề này: {topic}

---
## Bước 2: Agent 02 (Chính sách & SEO)
**System Instruction:**
{base_prompt}
---
{agent2_prompt}

**User Input:**
Hãy áp từ khóa SEO cho kết quả nghiên cứu của Agent 01.

---
## Bước 3: Agent 03 (Creative Director)
**System Instruction:**
{base_prompt}
---
{agent3_prompt}

**User Input:**
Hãy phát triển 3 hướng tiếp cận (Hooks) dựa trên định hướng của Agent 02.
"""
        manual_output_path = os.path.join(WORKSPACE_DIR, "scratch", "next_agent_prompts.md")
        write_file(manual_output_path, instructions)
        print(f"Đã tạo file hướng dẫn chi tiết tại: {manual_output_path}")
        return

    # Chế độ chạy tự động bằng LLM API
    if not api_key:
        api_key = os.environ.get(GEMINI_KEY_ENV)
    if not api_key:
        print("❌ Lỗi: Bạn cần cung cấp GEMINI_API_KEY để chạy ở chế độ tự động.")
        return
        
    print("\n🤖 Khởi động Pipeline tự động bằng Gemini 1.5 Flash...")
    
    # --- AGENT 01 ---
    print("\n1. Đang chạy Agent 01 (Market Researcher)...")
    system_1 = f"{base_prompt}\n\n{agent1_prompt}"
    user_1 = f"Hãy nghiên cứu chủ đề đầu vào sau và cung cấp dữ liệu SSoT:\n{topic}"
    res_1 = call_gemini(system_1, user_1, api_key)
    if not res_1: return
    write_file(os.path.join(WORKSPACE_DIR, "scratch", "pipeline_outputs", "01_research_output.md"), res_1)
    print("🟢 Hoàn thành Agent 01.")

    # --- AGENT 02 ---
    print("2. Đang chạy Agent 02 (Policy & SEO Consultant)...")
    system_2 = f"{base_prompt}\n\n{agent2_prompt}"
    user_2 = f"Dựa trên báo cáo nghiên cứu dưới đây, hãy đề xuất từ khóa SEO và quy định cấm:\n\n{res_1}"
    res_2 = call_gemini(system_2, user_2, api_key)
    if not res_2: return
    write_file(os.path.join(WORKSPACE_DIR, "scratch", "pipeline_outputs", "02_seo_output.md"), res_2)
    print("🟢 Hoàn thành Agent 02.")

    # --- AGENT 03 ---
    print("3. Đang chạy Agent 03 (Creative Director)...")
    system_3 = f"{base_prompt}\n\n{agent3_prompt}"
    user_3 = f"Hãy phát triển 3 hướng tiếp cận (Góc viết) từ đầu ra của Agent 02:\n\n{res_2}"
    res_3 = call_gemini(system_3, user_3, api_key)
    if not res_3: return
    write_file(os.path.join(WORKSPACE_DIR, "scratch", "pipeline_outputs", "03_creative_output.md"), res_3)
    print("🟢 Hoàn thành Agent 03.")

    # --- AGENT 04 ---
    print("4. Đang chạy Agent 04 (Content Planner)...")
    system_4 = f"{base_prompt}\n\n{agent4_prompt}\n\nLưu ý quan trọng: Trả về kết quả lịch biên tập dưới định dạng JSON thô duy nhất, khớp với trường 'DATA' để nhập vào excel calendar script."
    user_4 = f"Dựa trên các hướng tiếp cận sau, hãy lập lịch đăng tuần (gồm 4 ngày đăng) và xuất ra mảng JSON chứa các trường: 'Ngày đăng', 'Nền tảng', 'Pillar', 'Persona Giả định', 'Định dạng', 'Góc tiếp cận', 'Option A (Chủ đề chính)', 'Option B (Dự phòng 1)', 'Option C (Dự phòng 2)', 'Nguồn SSoT / Cốt truyện', 'Trạng thái Duyệt' (đặt mặc định 'Chờ duyệt').\n\n{res_3}"
    res_4 = call_gemini(system_4, user_4, api_key)
    if not res_4: return
    
    # Thử bóc tách JSON ra khỏi markdown code block
    json_str = res_4.strip()
    if json_str.startswith("```json"):
        json_str = json_str[7:]
    elif json_str.startswith("```"):
        json_str = json_str[3:]
    if json_str.endswith("```"):
        json_str = json_str[:-3]
    json_str = json_str.strip()
    
    write_file(os.path.join(WORKSPACE_DIR, "scratch", "weekly_calendar_input.json"), json_str)
    print("🟢 Hoàn thành Agent 04 (Đã ghi file JSON đầu vào cho Lịch Excel).")
    
    # Chạy script tạo Excel Calendar tự động
    print("\n📊 Đang tự động cập nhật Lịch biểu biên tập sang Excel...")
    os.system(f"python \"{os.path.join(WORKSPACE_DIR, 'scripts', 'create_excel_calendar.py')}\"")
    
    print("\n🎉 PIPELINE TỰ ĐỘNG CHẠY THÀNH CÔNG!")
    print(f"Đầu ra đã được cập nhật tại: {os.path.join(WORKSPACE_DIR, 'reports', 'plan1', 'Weekly_Content_Calendar.xlsx')}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Trình điều phối (Orchestrator) tự động hóa hệ thống PBAN.")
    parser.add_argument("--topic", default="", help="Chủ đề hoặc ý tưởng thô đầu vào.")
    parser.add_argument("--run-llm", action="store_true", help="Chạy tự động gọi LLM API.")
    parser.add_argument("--key", default="", help="SerpAPI hoặc Gemini API key.")
    
    args = parser.parse_args()
    run_pban_pipeline(args.topic, args.run_llm, args.key)
