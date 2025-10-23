import subprocess
import sys
import os

# Using system() method to
# execute shell commands
def main():
    pptx_path = sys.argv[1];
    print(pptx_path)

    if not os.path.exists(pptx_path):
        sys.exit(1)

    txt_path = os.path.splitext(pptx_path)[0] + ".txt"
    extractor_script = os.path.join(os.path.dirname(__file__), "ppt-to-txt.py")

    if os.path.exists(txt_path):
        print(f"Output file already exists — overwriting: {txt_path}")
        try:
            os.remove(txt_path)
        except Exception as e:
            print(f"Could not remove existing file: {e}")
            sys.exit(1)

    result = subprocess.run(
        ["python", extractor_script, pptx_path, txt_path],
        capture_output=True,
        text=True
    )

    print(result.stdout)
    if result.stderr:
        print(result.stderr)


    if os.name == "nt":
        subprocess.Popen(["notepad", txt_path])

if __name__ == "__main__":
    main()
