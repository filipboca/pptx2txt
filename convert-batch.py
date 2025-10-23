import os
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

def select_or_create_folder():
    """
    Opens a GUI to select a folder and optionally create a new subfolder.
    Returns the full path to the selected/new folder.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window

    # Ask user to select an existing folder
    selected_folder = filedialog.askdirectory(title="Select a folder for the destination")
    if not selected_folder:
        messagebox.showinfo("Cancelled", "No folder selected.")
        return None
    return selected_folder

def main():
    # Check for source folder argument
    if len(sys.argv) != 2:
        print("Usage: python main_script.py <absolute_path_to_source_folder>")
        sys.exit(1)

    source_folder = sys.argv[1]
    if not os.path.isdir(source_folder):
        print(f"Error: Source folder '{source_folder}' does not exist.")
        sys.exit(1)

    # Get destination folder from GUI
    destination_folder = select_or_create_folder()
    if not destination_folder:
        print("Operation cancelled.")
        sys.exit(1)

    # Run the secondary script
    try:
        subprocess.run(
            ["python", "pptx-to-txt-batch.py", source_folder, destination_folder],
            check=True
        )
        print("Secondary script executed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error running secondary script: {e}")

if __name__ == "__main__":
    main()
