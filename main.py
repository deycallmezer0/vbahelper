from openpyxl import load_workbook
import win32com.client
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pythoncom
from datetime import datetime

def create_excel_instance():
    """
    Creates an Excel instance with proper COM initialization
    """
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        return excel
    except Exception as e:
        raise Exception(f"Could not create Excel instance: {str(e)}")

def analyze_excel_file(file_path):
    """
    Analyzes an Excel file to count pages and VBA modules.
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        dict: Analysis results including pages, modules, and code
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    
    results = {
        'filename': os.path.basename(file_path),
        'full_path': file_path,
        'analysis_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'worksheets': [],
        'modules': [],
        'module_code': {}
    }
    
    # Load workbook for page count
    wb = load_workbook(file_path)
    
    # Get worksheet information
    for sheet in wb.worksheets:
        sheet_info = {
            'name': sheet.title,
            'has_content': sheet.calculate_dimension() != 'A1:A1',
            'print_area': sheet.print_area if sheet.print_area else 'Not Set'
        }
        results['worksheets'].append(sheet_info)
    
    # Get VBA information
    excel = None
    try:
        excel = create_excel_instance()
        wb_vba = excel.Workbooks.Open(os.path.abspath(file_path), ReadOnly=True)
        
        for component in wb_vba.VBProject.VBComponents:
            module_info = {
                'name': component.Name,
                'type': component.Type,  # 1=Standard, 2=Class, 3=Form, etc.
                'code_lines': component.CodeModule.CountOfLines
            }
            results['modules'].append(module_info)
            
            # Get the actual code
            if component.CodeModule.CountOfLines > 0:
                code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                results['module_code'][component.Name] = code
            else:
                results['module_code'][component.Name] = "(Empty Module)"
        
        wb_vba.Close(SaveChanges=False)
    except Exception as e:
        if "Programmatic access to Visual Basic Project is not trusted" in str(e):
            results['vba_error'] = "VBA access not enabled in Excel Trust Center"
        else:
            results['vba_error'] = str(e)
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass
            pythoncom.CoUninitialize()
    
    return results

def save_analysis_to_file(results, output_dir=None):
    """
    Saves analysis results to a text file.
    """
    if output_dir is None:
        output_dir = os.path.dirname(results['full_path'])
    
    # Create filename based on Excel file name and timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = os.path.splitext(results['filename'])[0]
    output_file = os.path.join(output_dir, f"{base_name}_analysis_{timestamp}.txt")
    
    with open(output_file, 'w', encoding='utf-8') as f:
        # Write header
        f.write("=" * 80 + "\n")
        f.write("Excel File Analysis Report\n")
        f.write("=" * 80 + "\n\n")
        
        # File information
        f.write("FILE INFORMATION\n")
        f.write("-" * 50 + "\n")
        f.write(f"Filename: {results['filename']}\n")
        f.write(f"Full Path: {results['full_path']}\n")
        f.write(f"Analysis Date: {results['analysis_date']}\n\n")
        
        # Worksheet information
        f.write("WORKSHEET INFORMATION\n")
        f.write("-" * 50 + "\n")
        f.write(f"Total Worksheets: {len(results['worksheets'])}\n")
        for sheet in results['worksheets']:
            f.write(f"\nSheet Name: {sheet['name']}\n")
            f.write(f"Has Content: {sheet['has_content']}\n")
            f.write(f"Print Area: {sheet['print_area']}\n")
        f.write("\n")
        
        # VBA Module information
        if 'vba_error' in results:
            f.write("VBA MODULE INFORMATION\n")
            f.write("-" * 50 + "\n")
            f.write(f"Error: {results['vba_error']}\n\n")
        else:
            f.write("VBA MODULE INFORMATION\n")
            f.write("-" * 50 + "\n")
            f.write(f"Total Modules: {len(results['modules'])}\n")
            for module in results['modules']:
                f.write(f"\nModule Name: {module['name']}\n")
                f.write(f"Type: {module['type']}\n")
                f.write(f"Code Lines: {module['code_lines']}\n")
            f.write("\n")
            
            # Module Code
            f.write("VBA MODULE CODE\n")
            f.write("-" * 50 + "\n")
            for module_name, code in results['module_code'].items():
                f.write(f"\n{module_name}\n")
                f.write("-" * len(module_name) + "\n")
                f.write(code + "\n")
                f.write("\n")
    
    return output_file

def select_file():
    """
    Opens a file dialog to select an Excel file.
    """
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[
            ("Excel files", "*.xlsx *.xlsm *.xls"),
            ("All files", "*.*")
        ]
    )
    
    return file_path

def main():
    try:
        file_path = select_file()
        if not file_path:
            print("No file selected")
            return
        
        # Analyze file
        results = analyze_excel_file(file_path)
        
        # Save results
        output_file = save_analysis_to_file(results)
        
        # Show completion message
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "Analysis Complete", 
            f"Analysis has been saved to:\n{output_file}"
        )
        
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()