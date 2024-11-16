import os
import win32com.client

def convert_ppt_to_pdf(directory):
    # Initialize PowerPoint application
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1  # Ensure PowerPoint is visible

    try:
        # Loop through all files in the given directory
        for filename in os.listdir(directory):
            if filename.lower().endswith((".ppt", ".pptx")):
                ppt_path = os.path.join(directory, filename)
                pdf_path = os.path.join(directory, os.path.splitext(filename)[0] + ".pdf")
                
                # Open the PowerPoint file
                presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
                
                # Save as PDF
                presentation.SaveAs(pdf_path, 32)  # 32 is the constant for PDF format
                presentation.Close()
                
                print(f"Converted: {ppt_path} to {pdf_path}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Quit PowerPoint application
        powerpoint.Quit()

# Example usage
directory_path = r"D:\Downloads\DUCSE Documents\2nd year\2-2\Mine\Data_and_Telecom\Slides"
convert_ppt_to_pdf(directory_path)
