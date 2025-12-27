import os
import sys
from pptx_to_txt_batch import extract_text_from_pptx

def print_progress(done, total, current_file):
    msg = f"[ {done:>4} / {total:<4} ] Converted: {current_file}"
    # Pad message to fully overwrite previous output
    print("\r" + msg.ljust(100), end="", flush=True)

def convert_all_pptx_to_single_folder(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    # 1️⃣ Collect all pptx files first
    pptx_files = []
    for root, _, files in os.walk(input_folder):
        for file in files:
            if file.lower().endswith(".pptx"):
                pptx_files.append(os.path.join(root, file))

    total = len(pptx_files)

    if total == 0:
        print("No .pptx files found.")
        return

    print(f"Found {total} .pptx files.\n")

    converted = 0

    # 2️⃣ Convert one by one with persistent progress view
    for pptx_path in pptx_files:
        relative_path = os.path.relpath(pptx_path, input_folder)

        # Flatten filename to avoid collisions
        safe_name = relative_path.replace(os.sep, "_")
        txt_name = os.path.splitext(safe_name)[0] + ".txt"
        output_path = os.path.join(output_folder, txt_name)

        try:
            text = extract_text_from_pptx(pptx_path)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(text)

            converted += 1
            print_progress(converted, total, safe_name)

        except Exception as e:
            print(f"\n❌ Error processing {pptx_path}: {e}")

    print("\n\n✅ Conversion complete.")
    print(f"Total converted: {converted} / {total}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python convert_all_to_single_folder.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_folder = sys.argv[2]

    convert_all_pptx_to_single_folder(input_folder, output_folder)
