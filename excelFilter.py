import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from tkinter import ttk

def filter_and_copy_excel(file_path, sheet_name, column_name, filter_value, output_sheet_name):
    try:
        # Load the Excel file
        xls = pd.ExcelFile(file_path)

        # Check if the specified sheet exists
        if sheet_name not in xls.sheet_names:
            messagebox.showerror("Error", f"Sheet '{sheet_name}' does not exist in the Excel file.")
            return

        # Read the specified sheet into a DataFrame
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Check if the specified column exists
        if column_name not in df.columns:
            messagebox.showerror("Error", f"Column '{column_name}' does not exist in the sheet '{sheet_name}'.")
            return

        # Add "Filtered" column if it does not exist
        if 'Filtered' not in df.columns:
            df['Filtered'] = None

        # Filter the DataFrame based on the specified column and value (partial match) and not previously filtered
        filtered_df = df[df[column_name].str.contains(filter_value, case=False, na=False)]

        # Check if any rows are filtered
        num_rows_copied = len(filtered_df)

        if num_rows_copied == 0:
            messagebox.showinfo("No Results", f"No new rows found containing '{filter_value}' in column '{column_name}'.")
            return

        # Mark filtered rows with the output sheet name in the original DataFrame
        def append_sheet_name(existing_value):
            if pd.isna(existing_value):
                return output_sheet_name
            else:
                existing_sheets = existing_value.split(", ")
                if output_sheet_name not in existing_sheets:
                    existing_sheets.append(output_sheet_name)
                return ", ".join(existing_sheets)

        df.loc[filtered_df.index, 'Filtered'] = df.loc[filtered_df.index, 'Filtered'].apply(append_sheet_name)

        # Remove the "Filtered" column from the filtered DataFrame before copying
        filtered_df = filtered_df.drop(columns=['Filtered'])

        # Create a Pandas Excel writer using the same file path
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            # Write the filtered DataFrame to a new sheet in the same Excel file
            filtered_df.to_excel(writer, sheet_name=output_sheet_name, index=False)
            # Write the updated original DataFrame back to its original sheet
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        messagebox.showinfo("Success", f"Filtered data written to sheet '{output_sheet_name}' in the Excel file.\n{num_rows_copied} rows copied.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def auto_detect_and_copy(file_path, sheet_name, column_name):
    try:
        # Load the Excel file
        xls = pd.ExcelFile(file_path)

        # Check if the specified sheet exists
        if sheet_name not in xls.sheet_names:
            messagebox.showerror("Error", f"Sheet '{sheet_name}' does not exist in the Excel file.")
            return

        # Read the specified sheet into a DataFrame
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Check if the specified column exists
        if column_name not in df.columns:
            messagebox.showerror("Error", f"Column '{column_name}' does not exist in the sheet '{sheet_name}'.")
            return

        # Add "Filtered" column if it does not exist
        if 'Filtered' not in df.columns:
            df['Filtered'] = None

        # Process each sheet in the Excel file except the input sheet
        for output_sheet_name in xls.sheet_names:
            if output_sheet_name == sheet_name:
                continue

            # Filter the DataFrame based on the column and sheet name (partial match)
            filtered_df = df[df[column_name].str.contains(output_sheet_name, case=False, na=False)]

            # Check if any rows are filtered
            num_rows_copied = len(filtered_df)

            if num_rows_copied == 0:
                continue

            # Mark filtered rows with the output sheet name in the original DataFrame
            def append_sheet_name(existing_value):
                if pd.isna(existing_value):
                    return output_sheet_name
                else:
                    existing_sheets = existing_value.split(", ")
                    if output_sheet_name not in existing_sheets:
                        existing_sheets.append(output_sheet_name)
                    return ", ".join(existing_sheets)

            df.loc[filtered_df.index, 'Filtered'] = df.loc[filtered_df.index, 'Filtered'].apply(append_sheet_name)

            # Remove the "Filtered" column from the filtered DataFrame before copying
            filtered_df = filtered_df.drop(columns=['Filtered'])

            # Create a Pandas Excel writer using the same file path
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Write the filtered DataFrame to a new sheet in the same Excel file
                filtered_df.to_excel(writer, sheet_name=output_sheet_name, index=False)
                # Write the updated original DataFrame back to its original sheet
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        messagebox.showinfo("Success", f"Auto-detection and copying completed successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_file_dialog_manual():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    if file_path:
        try:
            # Load the Excel file
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names
            sheet_name = simpledialog.askstring("Input", f"Enter the sheet name:\nAvailable sheets: {', '.join(sheet_names)}")
            if sheet_name not in xls.sheet_names:
                messagebox.showerror("Error", f"Sheet '{sheet_name}' does not exist in the Excel file.")
                return

            # Read the specified sheet into a DataFrame
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            column_names = df.columns.tolist()
            column_name = simpledialog.askstring("Input", f"Enter the column name to filter by:\nAvailable columns: {', '.join(column_names)}")
            if column_name not in df.columns:
                messagebox.showerror("Error", f"Column '{column_name}' does not exist in the sheet '{sheet_name}'.")
                return

            # Get the filter value and output sheet name from the user
            filter_value = simpledialog.askstring("Input", "Enter the value to filter for:")
            output_sheet_name = simpledialog.askstring("Input", "Enter the output sheet name:")
            if filter_value and output_sheet_name:
                summary = (
                    f"File Path: {file_path}\n"
                    f"Sheet Name: {sheet_name}\n"
                    f"Column Name: {column_name}\n"
                    f"Filter Value: {filter_value}\n"
                    f"Output Sheet Name: {output_sheet_name}"
                )
                if messagebox.askokcancel("Confirm Details", f"Please confirm the details:\n\n{summary}"):
                    filter_and_copy_excel(file_path, sheet_name, column_name, filter_value, output_sheet_name)
            else:
                messagebox.showerror("Error", "Filter value and output sheet name must be provided.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

def open_file_dialog_auto():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    if file_path:
        try:
            # Load the Excel file
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names
            sheet_name = simpledialog.askstring("Input", f"Enter the input sheet name:\nAvailable sheets: {', '.join(sheet_names)}")
            if sheet_name not in xls.sheet_names:
                messagebox.showerror("Error", f"Sheet '{sheet_name}' does not exist in the Excel file.")
                return

            # Read the specified sheet into a DataFrame
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            column_names = df.columns.tolist()
            column_name = simpledialog.askstring("Input", f"Enter the column name to filter by:\nAvailable columns: {', '.join(column_names)}")
            if column_name not in df.columns:
                messagebox.showerror("Error", f"Column '{column_name}' does not exist in the sheet '{sheet_name}'.")
                return

            summary = (
                f"File Path: {file_path}\n"
                f"Sheet Name: {sheet_name}\n"
                f"Column Name: {column_name}"
            )
            if messagebox.askokcancel("Confirm Details", f"Please confirm the details:\n\n{summary}"):
                auto_detect_and_copy(file_path, sheet_name, column_name)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

def show_readme():
    readme_text = (
        "Excel Filter and Copy Tool\n\n"
        "Manual Mode:\n"
        "1. Click 'Select Excel File (Manual)'.\n"
        "2. Choose the Excel file to process.\n"
        "3. Enter the sheet name containing the data.\n"
        "4. Enter the column name to filter by.\n"
        "5. Enter the value to filter for.\n"
        "6. Enter the output sheet name to save the filtered data.\n\n"
        "Auto Detect Mode:\n"
        "1. Click 'Select Excel File (Auto Detect)'.\n"
        "2. Choose the Excel file to process.\n"
        "3. Enter the sheet name containing the data.\n"
        "4. Enter the column name to filter by.\n"
        "5. The program will automatically filter and copy data to sheets named after the detected values.\n\n"
        "Note: The program will update the original sheet by marking the filtered rows."
    )
    messagebox.showinfo("Readme", readme_text)

# Create the main window
root = tk.Tk()
root.title("Excel Filter and Copy")
root.geometry("300x200")
root.configure(bg="#f0f0f0")

style = ttk.Style()
style.configure("TButton", padding=6, relief="flat", background="#ccc")

# Add buttons to open the file dialog for manual and auto modes
open_button_manual = ttk.Button(root, text="Select Excel File (Manual)", command=open_file_dialog_manual)
open_button_manual.pack(pady=10)

open_button_auto = ttk.Button(root, text="Select Excel File (Auto Detect)", command=open_file_dialog_auto)
open_button_auto.pack(pady=10)

readme_button = ttk.Button(root, text="Readme", command=show_readme)
readme_button.pack(pady=10)

# Run the Tkinter event loop
root.mainloop()
