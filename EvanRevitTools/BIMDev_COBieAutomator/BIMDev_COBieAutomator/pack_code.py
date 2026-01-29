import os
import pyperclip

# --- 設定區 ---
# 1. 專案靈魂文件
RULES_FILE = '_AI_RULES.md'

# 2. 要抓取的副檔名 (根據你的專案截圖調整)
TARGET_EXTENSIONS = ['.cs', '.xaml', '.config', '.xml'] 

# 3. 忽略清單 (忽略編譯產出的垃圾桶)
IGNORE_FOLDERS = ['obj', 'bin', '.git', '.vs', 'packages', 'Properties'] 

def pack_project():
    output_text = ""
    current_dir = os.getcwd()
    file_count = 0
    
    print(f"🚀 正在為 BIMDev_COBieAutomator 打包上下文...")
    
    # --- PART 1: 讀取靈魂 (規則書) ---
    rules_path = os.path.join(current_dir, RULES_FILE)
    if os.path.exists(rules_path):
        print(f"✅ 讀取專案規則: {RULES_FILE}")
        output_text += f"=== 🛑 PROJECT CONTEXT & RULES (READ THIS FIRST) ===\n"
        try:
            with open(rules_path, 'r', encoding='utf-8') as f:
                output_text += f.read()
        except Exception as e:
            output_text += f"[讀取規則錯誤: {e}]"
        output_text += f"\n=== END OF RULES ===\n\n"
    else:
        print(f"⚠️  (尚未找到 {RULES_FILE}，目前只打包程式碼)")

    # --- PART 2: 讀取肉體 (程式碼) ---
    print(f"📂 掃描路徑: {current_dir}")
    for root, dirs, files in os.walk(current_dir):
        # 排除忽略的資料夾
        dirs[:] = [d for d in dirs if d not in IGNORE_FOLDERS]
        
        for file in files:
            file_extension = os.path.splitext(file)[1].lower()
            if file_extension in TARGET_EXTENSIONS:
                file_path = os.path.join(root, file)
                relative_path = os.path.relpath(file_path, current_dir)
                
                output_text += f"\n--- FILE: {relative_path} ---\n"
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        output_text += f.read()
                    file_count += 1
                except Exception as e:
                    output_text += f"[讀取錯誤: {e}]"
                output_text += f"\n--- END OF FILE ---\n"

    # --- 輸出結果 ---
    output_filename = "_context_for_ai.txt"
    with open(output_filename, "w", encoding="utf-8") as f:
        f.write(output_text)

    try:
        pyperclip.copy(output_text)
        print("📋 內容已複製到剪貼簿！")
    except:
        print("⚠️ 無法自動複製到剪貼簿 (可能未安裝 pyperclip)，但檔案已儲存。")

    print(f"📦 打包完成！共處理 {file_count} 個檔案。")
    print(f"💾 檔案已儲存為: {output_filename}")
    input("按 Enter 鍵結束...")

if __name__ == "__main__":
    pack_project()