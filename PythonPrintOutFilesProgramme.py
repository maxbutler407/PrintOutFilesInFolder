import os
import win32com.client

def print_docx(file_path, printer_name):
    try:
        # Create a COM object for Word application
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(file_path)
        
        # Set the printer
        word.ActivePrinter = printer_name
        
        # Print the document
        doc.PrintOut()
        
        # Close the document and quit Word application
        doc.Close(False)
        word.Quit()
        
        print(f"Successfully printed: {file_path}")
    except Exception as e:
        print(f"Failed to print {file_path}: {e}")

def main():
    directory = "C:\\Programming\\Automation Projects\\Awake\\PrintOutFilesProgramme\\PrintTheseFiles"  # Adjust to your directory path
    printer_name = "EPSON242106 (ET-4750 Series)"  # Replace with your printer's name

    if not os.path.isdir(directory):
        print(f"Directory not found: {directory}")
        return

    files = [os.path.join(directory, f) for f in os.listdir(directory) if f.lower().endswith('.docx')]

    if not files:
        print("No .docx files to print.")
        return

    print("Printing files...")
    for file_path in files:
        print_docx(file_path, printer_name)

if __name__ == "__main__":
    main()