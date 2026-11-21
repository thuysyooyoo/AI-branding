import asyncio
import os
import openpyxl
import edge_tts

async def gen_single(text, voice, out_path, max_retries=4):
    for attempt in range(1, max_retries + 1):
        try:
            comm = edge_tts.Communicate(text, voice, rate="+0%", pitch="+0Hz")
            await comm.save(out_path)
            if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                return True
        except Exception as e:
            print(f"  -> Lỗi lần {attempt}: {e}. Đang thử lại sau {attempt}s...")
            await asyncio.sleep(attempt * 1.0)
    return False

async def generate_all_audio():
    excel_path = r'C:\Users\thuyn\Downloads\phân cảnh.xlsx'
    audio_dir = r'C:\Users\thuyn\Documents\Flow_Storyboards\audio'
    os.makedirs(audio_dir, exist_ok=True)

    wb = openpyxl.load_workbook(excel_path)
    ws = wb['KICH_BAN_FLOW']

    voice = 'vi-VN-HoaiMyNeural'
    print(f"Bắt đầu sinh thoại hàng loạt với giọng: {voice}")
    print(f"Thư mục đích: {audio_dir}")

    scenes = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        stt = row[0]
        title = row[1]
        dialogue = str(row[3]).strip()
        expected_dur = row[4]
        scenes.append((stt, title, dialogue, expected_dur))

    print(f"Tổng số phân cảnh: {len(scenes)}")

    full_text_parts = []
    success_count = 0

    for stt, title, dialogue, expected_dur in scenes:
        out_filename = f"Scene_{stt:02d}.mp3"
        out_path = os.path.join(audio_dir, out_filename)
        
        # If file already exists and > 10KB, skip or re-generate
        if os.path.exists(out_path) and os.path.getsize(out_path) > 10000:
            print(f"[{stt:02d}/{len(scenes):02d}] Đã có sẵn: {out_filename} ({os.path.getsize(out_path)/1024:.1f} KB)")
            full_text_parts.append(dialogue)
            success_count += 1
            continue

        ok = await gen_single(dialogue, voice, out_path)
        if ok:
            file_size = os.path.getsize(out_path)
            full_text_parts.append(dialogue)
            success_count += 1
            print(f"[{stt:02d}/{len(scenes):02d}] Đã sinh {out_filename} ({file_size/1024:.1f} KB)")
        else:
            print(f"[{stt:02d}/{len(scenes):02d}] THẤT BẠI: {out_filename}")
        
        await asyncio.sleep(0.4)

    # Generate full audio file
    full_audio_path = os.path.join(audio_dir, "Full_Voiceover_Bao_Bi_Thuc_Pham.mp3")
    print(f"\nĐang sinh file tổng hợp Full Voiceover cho {len(full_text_parts)} cảnh...")
    full_text = " ... ".join(full_text_parts)
    full_ok = await gen_single(full_text, voice, full_audio_path, max_retries=5)
    if full_ok:
        full_size = os.path.getsize(full_audio_path)
        print(f"HOÀN THÀNH TẤT CẢ! File tổng hợp: Full_Voiceover_Bao_Bi_Thuc_Pham.mp3 ({full_size/1024:.1f} KB)")
    else:
        print("Không tạo được file tổng hợp.")

    print(f"\nTổng kết: {success_count}/{len(scenes)} cảnh đã sẵn sàng tại {audio_dir}")

if __name__ == "__main__":
    asyncio.run(generate_all_audio())
