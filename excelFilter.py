import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox


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
            messagebox.showinfo("No Results",
                                f"No new rows found containing '{filter_value}' in column '{column_name}'.")
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

        messagebox.showinfo("Success",
                            f"Filtered data written to sheet '{output_sheet_name}' in the Excel file.\n{num_rows_copied} rows copied.")
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

            # Get the sheet name from the user
            sheet_name = simpledialog.askstring("Input", "Enter the sheet name:")
            if sheet_name not in xls.sheet_names:
                messagebox.showerror("Error", f"Sheet '{sheet_name}' does not exist in the Excel file.")
                return

            # Read the specified sheet into a DataFrame
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # Get the column name from the user
            column_name = simpledialog.askstring("Input", "Enter the column name to filter by:")
            if column_name not in df.columns:
                messagebox.showerror("Error", f"Column '{column_name}' does not exist in the sheet '{sheet_name}'.")
                return

            # Get the filter value and output sheet name from the user
            filter_value = simpledialog.askstring("Input", "Enter the value to filter for:")
            output_sheet_name = simpledialog.askstring("Input", "Enter the output sheet name:")
            if filter_value and output_sheet_name:
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

            # Get the sheet name from the user
            sheet_name = simpledialog.askstring("Input", "Enter the input sheet name:")
            if sheet_name not in xls.sheet_names:
                messagebox.showerror("Error", f"Sheet '{sheet_name}' does not exist in the Excel file.")
                return

            # Get the column name from the user
            column_name = simpledialog.askstring("Input", "Enter the column name to filter by:")
            if column_name not in pd.read_excel(file_path, sheet_name=sheet_name).columns:
                messagebox.showerror("Error", f"Column '{column_name}' does not exist in the sheet '{sheet_name}'.")
                return

            auto_detect_and_copy(file_path, sheet_name, column_name)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


# Create the main window
root = tk.Tk()
root.title("Excel Filter and Copy")
root.geometry("300x200")

# Add buttons to open the file dialog for manual and auto modes
open_button_manual = tk.Button(root, text="Select Excel File (Manual)", command=open_file_dialog_manual)
open_button_manual.pack(pady=10)

open_button_auto = tk.Button(root, text="Select Excel File (Auto Detect)", command=open_file_dialog_auto)
open_button_auto.pack(pady=10)

# Run the Tkinter event loop
root.mainloop()
