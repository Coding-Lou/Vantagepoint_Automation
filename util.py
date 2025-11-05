import json
from pypdf import PdfReader, PdfWriter
from pathlib import Path
import os

def show_welcome_banner():
    banner = r"""
     ██████╗  ██████═╗ █████╗ 
    ██╔═══██╗██      ║██╔══██╗
    ██║   ██║██      ║███████║
    ██║▄▄ ██║██      ║██╔══██║
    ╚██████╔╝╚██████╔╝██║  ██║
     ╚══▀▀═╝  ╚═════╝ ╚═╝  ╚═╝
        🔧 QCA Accounting Team Automation Script
    """
    print(banner)
    print("🚀 Welcome! First-time setup in progress...\n")

def update_config(key, value):
    try:
        with open("config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
        
        config[key] = value
        
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
            f.flush()

        print(f"💾 {key} been updated to {value}. ")

    except Exception as e:
        print(f"⚠️ {key} updated failed: ", e)

def merge_amazon_invoices():
    path = input("Please input the folder path: ")
    input_folder = Path(path)  
    output_file = input_folder / "combined_output.pdf"

    combined_writer = PdfWriter()

    for pdf_file in input_folder.glob("*.pdf"):
        try:
            reader = PdfReader(str(pdf_file))
            if len(reader.pages) >= 1:
                combined_writer.add_page(reader.pages[0])
            if len(reader.pages) == 2:
                combined_writer.add_page(reader.pages[1])
            if len(reader.pages) >= 3:
                combined_writer.add_page(reader.pages[2])
            print(f"✅ Done: {pdf_file.name}")
        except Exception as e:
            print(f"⚠️ Error {pdf_file.name}: {e}")

    with open(output_file, "wb") as f:
        combined_writer.write(f)

    print(f"\n🎉 Success: {output_file}")

def merge_pdfs(folder_path, output_filename):
    folder_path = input("📂 Please input folder path of the pdf files: ").strip()
    output_filename = "merged_output.pdf"

    if not os.path.isdir(folder_path):
        print("❌ Invalid folder path")

    pdf_writer = PdfWriter()

    pdf_files = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith('.pdf')
    ]

    # Sort by created time
    pdf_files.sort(key=lambda f: os.path.getctime(f))

    if not pdf_files:
        print("❌ No pdf files")
        return

    for pdf_path in pdf_files:
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                pdf_writer.add_page(page)
            print(f"✅ Add: {os.path.basename(pdf_path)}")
        except Exception as e:
            print(f"⚠️ Skip {pdf_path}: {e}")

    output_path = os.path.join(folder_path, output_filename)
    with open(output_path, "wb") as out_file:
        pdf_writer.write(out_file)

    print(f"\n🎉 Success the merged pdf file：{output_path}")

