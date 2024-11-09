# 🖼️ PPTX to PDF Converter

This Python script automates the conversion of `.pptx` (PowerPoint) files to `.pdf`. It scans a specified folder, converts all `.pptx` files found, and saves the converted PDFs in a separate `PDFs` folder.

## 🚀 Features
- **Batch Conversion**: Automatically converts all `.pptx` files in a folder to `.pdf`.
- **Organized Output**: Converted files are saved in a `PDFs` subfolder.
- **Error Handling**: Skips problematic files and provides clear error messages.
- **Easy to Use**: Just modify the folder path, run the script, and you're good to go!

## 📦 Prerequisites
Make sure you have the following installed:

- **Python** (version 3.6 or higher)
- **comtypes** library for PowerPoint interaction

Install `comtypes` with:
```bash
pip install comtypes
```

## 🛠️ How to Use

### 1. Clone the Repository
```bash
git clone https://github.com/TAUSEEF-01/.pptx_to_.pdf_converter.git
cd pptx-to-pdf-converter
```

### 2. Set Up a Virtual Environment (Optional but Recommended)
```bash
python -m venv myenv
myenv\Scripts\activate  # For Windows
```

### 3. Install Dependencies
```bash
pip install comtypes
```

### 4. Update the Folder Path
1. Go to myenv/files/ directory

2. In the script, update the `folder_path` variable with the path to the folder containing your `.pptx` files:
```python
folder_path = r'D:\Downloads\folder'
```

### 5. Run the Script
```bash
python pptx_to_pdf.py
```

### 6. 🎉 Check the Output
Converted PDF files will be available in a `PDFs` folder within the specified directory.

## 📝 Example Output
```text
Converted: Lecture 2 BFS.pptx to PDF.
Converted: Lecture 3 DFS.pptx to PDF.
Failed to convert Lecture 4.pptx: <error message>
```

## ⚙️ Customization
Feel free to modify the script to fit your specific needs, such as:
- Changing the output folder name
- Adding support for other file formats

## 🐛 Troubleshooting
- **COMError**: Ensure PowerPoint is installed and properly configured.
- **Permission Denied**: Check if you have write permissions for the output folder.
- **Corrupted File**: Manually verify problematic `.pptx` files.

## 📄 License
This project is licensed under the MIT License. Feel free to use, modify, and distribute as needed.

---

Made with ❤️ by [Tauseef](https://github.com/TAUSEEF-01) ✨
