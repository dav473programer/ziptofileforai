import zipfile
import tkinter as tk
from tkinter import Tk, Canvas
from tkinter import filedialog, messagebox
from tkinter import ttk
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os
import time

def write_zip_contents_to_pdf(zip_filename, pdf_filename, progress_callback):
    # Create a PDF file
    c = canvas.Canvas(pdf_filename, pagesize=letter)
    
    # Set the font for the title (ZIP file name)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(100, 750, f"Contents of ZIP file: {zip_filename}")
    
    # Set the font for the rest of the content (smaller)
    c.setFont("Helvetica", 8)
    y_position = 730  # Initial Y position for writing contents
    
    # Open the zip file
    with zipfile.ZipFile(zip_filename, 'r') as zip_ref:
        file_list = zip_ref.infolist()
        total_files = len(file_list)
        
        start_time = time.time()  # Track the start time for progress calculation
        
        # Iterate through each file in the zip
        for index, file_info in enumerate(file_list):
            file_name = file_info.filename
            
            # Set the font for file names (larger)
            c.setFont("Helvetica-Bold", 12)
            c.drawString(100, y_position, f"File: {file_name}")
            y_position -= 20
            
            # Set the font for the contents (smaller)
            c.setFont("Helvetica", 8)
            
            # Extract the file contents and write it to the PDF
            with zip_ref.open(file_info.filename) as file:
                file_content = file.read().decode('utf-8', errors='ignore')  # Read file content
                content_lines = file_content.splitlines()
                for line in content_lines:
                    c.drawString(100, y_position, line)
                    y_position -= 10  # Move down for the next line of content
                    
                    # Check if we need a new page
                    if y_position < 100:
                        c.showPage()
                        c.setFont("Helvetica", 8)
                        y_position = 750  # Reset Y position for new page

            y_position -= 20  # Add extra space after each file content
            
            # Calculate elapsed time and estimate remaining time
            elapsed_time = time.time() - start_time
            avg_time_per_file = elapsed_time / (index + 1)
            remaining_files = total_files - (index + 1)
            estimated_time_left = avg_time_per_file * remaining_files
            
            # Update progress
            progress_callback(index + 1, total_files, file_name, estimated_time_left)
            
    # Save the PDF
    c.save()

def write_zip_contents_to_text(zip_filename, text_filename, progress_callback):
    with open(text_filename, 'w', encoding='utf-8') as f:
        with zipfile.ZipFile(zip_filename, 'r') as zip_ref:
            file_list = zip_ref.infolist()
            total_files = len(file_list)
            
            start_time = time.time()  # Track the start time for progress calculation
            
            for index, file_info in enumerate(file_list):
                file_name = file_info.filename
                f.write(f"File: {file_name}\n")
                
                with zip_ref.open(file_info.filename) as file:
                    file_content = file.read().decode('utf-8', errors='ignore')
                    f.write(file_content + "\n\n")
                
                # Update progress
                elapsed_time = time.time() - start_time
                avg_time_per_file = elapsed_time / (index + 1)
                remaining_files = total_files - (index + 1)
                estimated_time_left = avg_time_per_file * remaining_files
                progress_callback(index + 1, total_files, file_name, estimated_time_left)

def write_zip_contents_to_csv(zip_filename, csv_filename, progress_callback):
    import csv
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow(["File Name", "Content"])  # Header row
        
        with zipfile.ZipFile(zip_filename, 'r') as zip_ref:
            file_list = zip_ref.infolist()
            total_files = len(file_list)
            
            start_time = time.time()  # Track the start time for progress calculation
            
            for index, file_info in enumerate(file_list):
                file_name = file_info.filename
                csv_writer.writerow([file_name, ''])
                
                with zip_ref.open(file_info.filename) as file:
                    file_content = file.read().decode('utf-8', errors='ignore')
                    csv_writer.writerow([file_name, file_content])
                
                # Update progress
                elapsed_time = time.time() - start_time
                avg_time_per_file = elapsed_time / (index + 1)
                remaining_files = total_files - (index + 1)
                estimated_time_left = avg_time_per_file * remaining_files
                progress_callback(index + 1, total_files, file_name, estimated_time_left)

def write_zip_contents_to_html(zip_filename, html_filename, progress_callback):
    with open(html_filename, 'w', encoding='utf-8') as f:
        f.write("<html><body><h1>Contents of ZIP file: {}</h1>".format(zip_filename))
        
        with zipfile.ZipFile(zip_filename, 'r') as zip_ref:
            file_list = zip_ref.infolist()
            total_files = len(file_list)
            
            start_time = time.time()  # Track the start time for progress calculation
            
            for index, file_info in enumerate(file_list):
                file_name = file_info.filename
                f.write(f"<h2>File: {file_name}</h2><pre>")
                
                with zip_ref.open(file_info.filename) as file:
                    file_content = file.read().decode('utf-8', errors='ignore')
                    f.write(file_content)
                
                f.write("</pre><hr>")
                
                # Update progress
                elapsed_time = time.time() - start_time
                avg_time_per_file = elapsed_time / (index + 1)
                remaining_files = total_files - (index + 1)
                estimated_time_left = avg_time_per_file * remaining_files
                progress_callback(index + 1, total_files, file_name, estimated_time_left)
        
        f.write("</body></html>")

def start_conversion(zip_filename, output_filename, progress_var, progress_label, root, export_type):
    def progress_callback(current, total, current_file, estimated_time_left):
        progress_var.set(current / total * 100)
        
        # Limit the filename to 30 characters
        truncated_filename = (current_file[:30] + '...') if len(current_file) > 30 else current_file
        
        # Update the label text with the truncated filename and estimated time
        progress_label.config(text=f"Processing: {truncated_filename} (ETA: {int(estimated_time_left)}s)")
        root.update_idletasks()

    if not zip_filename or not output_filename:
        messagebox.showerror("Error", "Please select both a ZIP file and a save location.")
        return

    try:
        if export_type == 'pdf':
            write_zip_contents_to_pdf(zip_filename, output_filename, progress_callback)
        elif export_type == 'text':
            write_zip_contents_to_text(zip_filename, output_filename, progress_callback)
        elif export_type == 'csv':
            write_zip_contents_to_csv(zip_filename, output_filename, progress_callback)
        elif export_type == 'html':
            write_zip_contents_to_html(zip_filename, output_filename, progress_callback)
        
        messagebox.showinfo("Success", f"{export_type.capitalize()} generated successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def select_zip_file():
    zip_filename = filedialog.askopenfilename(title="Select ZIP file", filetypes=[("ZIP files", "*.zip")])
    if zip_filename:
        zip_filename = zip_filename.strip()  # Remove any leading/trailing whitespace
        zip_label.config(text=f"Selected ZIP file | {zip_filename}")
        return zip_filename
    return None

def select_output_location(export_type):
    filetypes = [("Text files", "*.txt")] if export_type == 'text' else \
                [("CSV files", "*.csv")] if export_type == 'csv' else \
                [("HTML files", "*.html")] if export_type == 'html' else \
                [("PDF files", "*.pdf")]
    
    output_filename = filedialog.asksaveasfilename(title=f"Save {export_type.capitalize()} file", defaultextension=f".{export_type}", filetypes=filetypes)
    if output_filename:
        output_filename = output_filename.strip()  # Remove any leading/trailing whitespace
        output_label.config(text=f"Save {export_type.capitalize()} as | {output_filename}")
        return output_filename
    return None

def show_about():
    # Display a messagebox with the about information
    about_message = """
    ZIP to File Converter
    Version 1.0
    
    This program allows you to extract the contents of ZIP files
    and export them to various formats including PDF, Text,
    CSV, andHTML.The use of this program is primarily to be
    able to upload source code to an AI Assistant which doesnt
    support zip analyzing,for them to be able to analyze it and
    help you.

    In the export tab, you should be able to select the file
    formatin which you want the program to extract the
    information.
    
    Developed by: dav473programer
    https://github.com/dav473programer
    """
    messagebox.showinfo("About", about_message)

def run_gui():
    # Initialize Tkinter
    root = tk.Tk()  # Use TkinterDnD to enable drag-and-drop

    root.title("ZIP to File Converter")
    root.geometry("520x240")
    root.resizable(False, False)  # Disable both horizontal and vertical resizing
    
    # Initialize variables for file paths
    zip_filename = None
    output_filename = None
    export_type=None
    export_type = 'pdf'  # Default export type is PDF
    
    def set_zip_file(path):
        nonlocal zip_filename
        zip_filename = path
        zip_label.config(text=f"Selected ZIP file | {zip_filename}")

    def set_output_location(path):
        nonlocal output_filename
        output_filename = path
        output_label.config(text=f"Save file as | {output_filename}")

    def set_export_type(file_type):
        nonlocal export_type
        export_type = file_type
        type_label.config(text=f"Selected format | {export_type}")
      
    global zip_label, output_label
    tk.Label(root, text="ZIP to File Converter", font=("Helvetica", 16)).place(x=170, y=10)

    # Label to display the selected file type
    type_label = tk.Label(root, text="No format selected")
    type_label.place(x=10, y=105)

    # Button to select ZIP file
    zip_button = tk.Button(root, text="Select ZIP file", command=lambda: set_zip_file(select_zip_file()))
    zip_button.place(x=10, y=40)

    # Label to display the selected ZIP file
    zip_label = tk.Label(root, text="No ZIP file selected")
    zip_label.place(x=100, y=43)

    # Button to select the location to save the output
    output_button = tk.Button(root, text="Select Output Location", command=lambda: set_output_location(select_output_location(export_type)))
    output_button.place(x=10, y=70)

    # Label to display the output file location
    output_label = tk.Label(root, text="No save location selected")
    output_label.place(x=150, y=73)

    # Progress bar
    style = ttk.Style()

    # Change the color of the progress bar (light green for the fill, green for the background)
    style.configure("TProgressbar",
                    barcolor="red",  # Color of the filled progress bar
                    background="black")  # Color of the empty part of the progress bar
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, style="TProgressbar")
    progress_bar.place(x=10, y=130, width=500, height=35)  # Custom height and width for the progress bar

    # Label to show current file and estimated time
    progress_label = tk.Label(root, text="Processing: None (ETA: 0s)", bg="green", fg="white")
    progress_label.place(x=20, y=137)

    # Menu bar
    menu_bar = tk.Menu(root)

    # File menu
    file_menu = tk.Menu(menu_bar, tearoff=0)
    file_menu.add_command(label="Exit", command=root.quit)
    file_menu.add_command(label="About", command=show_about)  # Add the About option
    menu_bar.add_cascade(label="File", menu=file_menu)

    # Export menu
    export_menu = tk.Menu(menu_bar, tearoff=0)
    export_menu.add_command(label="Export as PDF", command=lambda: set_export_type('pdf'))
    export_menu.add_command(label="Export as Text", command=lambda: set_export_type('text'))
    export_menu.add_command(label="Export as CSV", command=lambda: set_export_type('csv'))
    export_menu.add_command(label="Export as HTML", command=lambda: set_export_type('html'))
    menu_bar.add_cascade(label="Export", menu=export_menu)

    # Display the menu bar
    root.config(menu=menu_bar)

    # Start button to begin conversion
    start_button = tk.Button(root, text="Start Conversion", command=lambda: start_conversion(zip_filename, output_filename, progress_var, progress_label, root, export_type), fg="white", bg="dark green", width=20, height=3)
    start_button.place(x=200, y=170)

    # Run the GUI
    root.mainloop()

# Run the GUI
run_gui()
