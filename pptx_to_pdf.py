import os
import comtypes.client

def pptx_to_pdf(folder_path):
    # Create output folder for PDFs if not exists
    output_folder = os.path.join(folder_path, 'PDFs')
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    
    for filename in os.listdir(folder_path):
        if filename.endswith('.pptx'):
            pptx_path = os.path.join(folder_path, filename)
            pdf_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '.pdf')
            
            try:
                presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)
                presentation.SaveAs(pdf_path, 32)  # 32 is the format for PDF
                presentation.Close()
                print(f"Converted: {filename} to PDF.")
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")
    
    powerpoint.Quit()

folder_path = r'D:\Downloads\DUCSE Documents\3rd year\Mine\3-1\CSE 3101 - Computer Networking\Theory\Slides' # enter your file path
pptx_to_pdf(folder_path)
