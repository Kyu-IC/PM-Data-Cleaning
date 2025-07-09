import os
import re

INPUT_DIR = "Input"
MAX_LENGTH = 50  # จำกัดความยาวชื่อไฟล์/โฟลเดอร์

def clean_name(name):
    # ตัดชื่อด้านหลังตามคำที่มักไม่สำคัญ
    split_keywords = [" - ", " ระหว่าง", "(", " ปริมาณ", " ช่วง", " ตั้งแต่", " ถึง", " กม."]
    for keyword in split_keywords:
        if keyword in name:
            name = name.split(keyword)[0]
            break

    # ล้างอักขระพิเศษ ยกเว้นอักษร/ตัวเลข/ช่องว่าง
    name = re.sub(r"[^\w\s]", "", name)
    # ช่องว่าง → _
    name = re.sub(r"\s+", "_", name)
    # จำกัดความยาว
    return name[:MAX_LENGTH]

def rename_files_and_folders(base_dir):
    for root, dirs, files in os.walk(base_dir, topdown=False):
        # 🔁 ไฟล์
        for filename in files:
            old_path = os.path.join(root, filename)
            name, ext = os.path.splitext(filename)
            new_filename = clean_name(name) + ext
            new_path = os.path.join(root, new_filename)
            if old_path != new_path:
                try:
                    os.rename(old_path, new_path)
                    print(f"✅ Renamed file: {old_path} -> {new_path}")
                except PermissionError:
                    print(f"⚠️ Skipped file (in use): {old_path}")
                except Exception as e:
                    print(f"❌ Error renaming file {old_path}: {e}")

        # 🔁 โฟลเดอร์
        for dirname in dirs:
            old_dir = os.path.join(root, dirname)
            new_dirname = clean_name(dirname)
            new_dir = os.path.join(root, new_dirname)
            if old_dir != new_dir:
                try:
                    os.rename(old_dir, new_dir)
                    print(f"✅ Renamed folder: {old_dir} -> {new_dir}")
                except PermissionError:
                    print(f"⚠️ Skipped folder (in use): {old_dir}")
                except Exception as e:
                    print(f"❌ Error renaming folder {old_dir}: {e}")

def main():
    print("🧼 Cleaning filenames and folder names in:", INPUT_DIR)
    if not os.path.exists(INPUT_DIR):
        print(f"❌ Folder '{INPUT_DIR}' not found.")
        return
    rename_files_and_folders(INPUT_DIR)
    print("✅ Filename cleaning complete.")

# สำหรับเรียกจาก command-line โดยตรง
if __name__ == "__main__":
    main()
