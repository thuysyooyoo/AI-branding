# -*- coding: utf-8 -*-
import os
import sys
import io
import re

# Đảm bảo in tiếng Việt trên console không bị lỗi
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

WORKSPACE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILLS_DIR = os.path.join(WORKSPACE_DIR, "skills")
AGENTS_DIR = os.path.join(WORKSPACE_DIR, "agents")
KNOWLEDGE_DIR = os.path.join(WORKSPACE_DIR, "knowledge")

MAX_SKILL_LINES = 500  # Tiêu chuẩn Superpowers Framework: SKILL.md < 500 lines để tối ưu token

def parse_yaml_frontmatter(file_path):
    """Đọc file và trích xuất YAML frontmatter ở đầu file"""
    if not os.path.exists(file_path):
        return None
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()
    
    match = re.match(r"^---\s*\n(.*?)\n---\s*\n", content, re.DOTALL)
    if not match:
        return {}
    
    yaml_text = match.group(1)
    metadata = {}
    for line in yaml_text.splitlines():
        if ":" in line:
            key, val = line.split(":", 1)
            metadata[key.strip()] = val.strip()
    return metadata

def count_file_lines(file_path):
    """Đếm số dòng của file"""
    with open(file_path, "r", encoding="utf-8") as f:
        return len(f.readlines())

def verify_skills():
    print("="*70)
    print("🔍 KIỂM THỬ TỰ ĐỘNG CHẤT LƯỢNG HỆ THỐNG SKILLS (SUPERPOWERS FRAMEWORK)")
    print("="*70)
    
    if not os.path.exists(SKILLS_DIR):
        print("❌ Lỗi: Thư mục 'skills/' không tồn tại!")
        return False
    
    all_passed = True
    skills_found = []
    
    # 1. Rà soát thư mục skills/
    print("\n📁 1. Kiểm tra các thư mục Skills & Giới hạn Token (< 500 dòng):")
    for skill_name in sorted(os.listdir(SKILLS_DIR)):
        skill_path = os.path.join(SKILLS_DIR, skill_name)
        if os.path.isdir(skill_path):
            print(f"  - Skill: {skill_name}")
            skill_md_path = os.path.join(skill_path, "SKILL.md")
            
            # Kiểm tra sự tồn tại của SKILL.md
            if not os.path.exists(skill_md_path):
                print(f"    ❌ Lỗi: Thiếu file SKILL.md tại {skill_path}")
                all_passed = False
                continue
            
            # Đếm số dòng
            lines_count = count_file_lines(skill_md_path)
            if lines_count > MAX_SKILL_LINES:
                print(f"    ⚠️ Cảnh báo Token: SKILL.md có {lines_count} dòng (> {MAX_SKILL_LINES} dòng tiêu chuẩn Superpowers)")
                all_passed = False
            else:
                print(f"    ✓ Độ dài Token OK: {lines_count} dòng (chuẩn < {MAX_SKILL_LINES} dòng)")
            
            # Đọc frontmatter metadata
            meta = parse_yaml_frontmatter(skill_md_path)
            if not meta or 'name' not in meta or 'description' not in meta:
                print(f"    ⚠️ Cảnh báo: YAML Frontmatter chưa đạt chuẩn trong {skill_md_path}")
                all_passed = False
            else:
                print(f"    ✓ Frontmatter OK: name='{meta.get('name')}' | desc='{meta.get('description', '')[:45]}...'")
            
            skills_found.append(skill_name)
            
    # 2. Rà soát các liên kết và tham chiếu trong Agent
    print("\n🤖 2. Kiểm tra tham chiếu Skill & Tri thức trong 8 Agent:")
    if not os.path.exists(AGENTS_DIR):
        print("❌ Lỗi: Thư mục 'agents/' không tồn tại!")
        return False
        
    for agent_file in sorted(os.listdir(AGENTS_DIR)):
        if agent_file.endswith(".md") and agent_file != "README.md":
            agent_path = os.path.join(AGENTS_DIR, agent_file)
            with open(agent_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
            
            # Tìm xem có chứa link dẫn tới thư mục skills không
            referenced = []
            for skill in skills_found:
                pattern = rf"skills/{skill}"
                if re.search(pattern, content):
                    referenced.append(skill)
            
            if referenced:
                print(f"  ✓ Agent {agent_file:30s} -> Liên kết {len(referenced)} skill: {referenced}")
            else:
                print(f"  ❌ Agent {agent_file:30s} -> Không phát hiện liên kết tới Skill nào!")
                all_passed = False

    # 3. Kiểm tra tính toàn vẹn của Kho Tri thức Knowledge
    print("\n📚 3. Kiểm tra tính toàn vẹn của Kho Tri thức (Knowledge Files):")
    if os.path.exists(KNOWLEDGE_DIR):
        k_files = [f for f in os.listdir(KNOWLEDGE_DIR) if f.endswith('.md')]
        print(f"  ✓ Đã xác minh {len(k_files)} tệp Tri thức nền tảng trong knowledge/")
    else:
        print("  ❌ Lỗi: Thư mục 'knowledge/' không tồn tại!")
        all_passed = False

    print("\n" + "="*70)
    if all_passed:
        print("🟢 KẾT LUẬN: HỆ THỐNG SKILLS ĐẠT TIÊU CHUẨN XÁC MINH SUPERPOWERS!")
        print("="*70)
        return True
    else:
        print("🔴 KẾT LUẬN: CÓ LỖI HOẶC CẢNH BÁO TRONG HỆ THỐNG SKILLS.")
        print("="*70)
        return False

if __name__ == "__main__":
    verify_skills()
