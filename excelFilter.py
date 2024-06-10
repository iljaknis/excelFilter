import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox


class ExcelFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Filter and Copy")
        self.root.geometry("400x300")

        self.file_path = None
        self.sheet_name = None
        self.column_name = None
        self.filter_value = None
        self.output_sheet_name = None

        self.inputs = {}
        self.current_step = 0


        self.steps = [
            self.select_excel_file,
            self.input_sheet_name,
            self.input_column_name,
            self.input_filter_value,
            self.input_output_sheet_name
        ]

        self.create_widgets()

    def create_widgets(self):
        self.frame = tk.Frame(self.root)
        self.frame.pack(pady=20)

        self.label = tk.Label(self.frame, text="Select an option:")
        self.label.pack(pady=10)

        self.manual_button = tk.Button(self.frame, text="Select Excel File (Manual)", command=self.start_manual)
        self.manual_button.pack(pady=5)

        self.auto_button = tk.Button(self.frame, text="Select Excel File (Auto Detect)", command=self.start_auto)
        self.auto_button.pack(pady=5)

        self.back_button = tk.Button(self.frame, text="Back", command=self.go_back, state=tk.DISABLED)
        self.back_button.pack(side=tk.LEFT, padx=10)

        self.cancel_button = tk.Button(self.frame, text="Cancel", command=self.cancel)
        self.cancel_button.pack(side=tk.RIGHT, padx=10)

    def start_manual(self):
        self.reset_inputs()
        self.manual_mode = True
        self.current_step = 0
        self.next_step()

    def start_auto(self):
        self.reset_inputs()
        self.manual_mode = False
        self.current_step = 0
        self.next_step()

    def reset_inputs(self):
        self.inputs = {}
        self.update_displayed_inputs()

    def next_step(self):
        if self.current_step < len(self.steps):
            self.steps[self.current_step]()
            self.current_step += 1

    def go_back(self):
        if self.current_step > 1:  # Ensure we don't go back to before the first step
            self.current_step -= 2  # To counteract the increment in next_step
            self.next_step()

    def cancel(self):
        self.root.destroy()

    def update_displayed_inputs(self):
        input_text = "\n".join([f"{key}: {value}" for key, value in self.inputs.items()])
        self.label.config(text=input_text)

    def select_excel_file(self):
        self.file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if self.file_path:
            self.inputs["File Path"] = self.file_path
            self.update_displayed_inputs()
            self.next_step()

    def input_sheet_name(self):
        self.sheet_name = simpledialog.askstring("Input", "Enter the sheet name:")
        if self.sheet_name:
            self.inputs["Sheet Name"] = self.sheet_name
            self.update_displayed_inputs()
            self.next_step()

    def input_column_name(self):
        self.column_name = simpledialog.askstring("Input", "Enter the column name to filter by:")
        if self.column_name:
            self.inputs["Column Name"] = self.column_name
            self.update_displayed_inputs()
            self.next_step()

    def input_filter_value(self):
        self.filter_value = simpledialog.askstring("Input", "Enter the value to filter for:")
        if self.filter_value:
            self.inputs["Filter Value"] = self.filter_value
            self.update_displayed_inputs()
            self.next_step()

    def input_output_sheet_name(self):
        self.output_sheet_name = simpledialog.askstring("Input", "Enter the output sheet name:")
        if self.output_sheet_name:
            self.inputs["Output Sheet Name"] = self.output_sheet_name
            self.update_displayed_inputs()
            if self.manual_mode:
                self.filter_and_copy_excel()
            else:
                self.auto_detect_and_copy()

    def filter_and_copy_excel(self):
        try:
            # Load the Excel file
            xls = pd.ExcelFile(self.file_path)

            # Check if the specified sheet exists
            if self.sheet_name not in xls.sheet_names:
                messagebox.showerror("Error", f"Sheet '{self.sheet_name}' does not exist in the Excel file.")
                return

            # Read the specified sheet into a DataFrame
            df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)

            # Check if the specified column exists
            if self.column_name not in df.columns:
                messagebox.showerror("Error",
                                     f"Column '{self.column_name}' does not exist in the sheet '{self.sheet_name}'.")
                return

            # Add "Filtered" column if it does not exist
            if 'Filtered' not in df.columns:
                df['Filtered'] = None

            # Filter the DataFrame based on the specified column and value (partial match) and not previously filtered
            filtered_df = df[df[self.column_name].str.contains(self.filter_value, case=False, na=False)]

            # Check if any rows are filtered
            num_rows_copied = len(filtered_df)

            if num_rows_copied == 0:
                messagebox.showinfo("No Results",
                                    f"No new rows found containing '{self.filter_value}' in column '{self.column_name}'.")
                return

            # Mark filtered rows with the output sheet name in the original DataFrame
            def append_sheet_name(existing_value):
                if pd.isna(existing_value):
                    return self.output_sheet_name
                else:
                    existing_sheets = existing_value.split(", ")
                    if self.output_sheet_name not in existing_sheets:
                        existing_sheets.append(self.output_sheet_name)
                    return ", ".join(existing_sheets)

            df.loc[filtered_df.index, 'Filtered'] = df.loc[filtered_df.index, 'Filtered'].apply(append_sheet_name)

            # Remove the "Filtered" column from the filtered DataFrame before copying
            filtered_df = filtered_df.drop(columns=['Filtered'])

            # Create a Pandas Excel writer using the same file path
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Write the filtered DataFrame to a new sheet in the same Excel file
                filtered_df.to_excel(writer, sheet_name=self.output_sheet_name, index=False)
                # Write the updated original DataFrame back to its original sheet
                df.to_excel(writer, sheet_name=self.sheet_name, index=False)

            messagebox.showinfo("Success",
                                f"Filtered data written to sheet '{self.output_sheet_name}' in the Excel file.\n{num_rows_copied} rows copied.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def auto_detect_and_copy(self):
        try:
            # Load the Excel file
            xls = pd.ExcelFile(self.file_path)

            print(self.current_step)


            # Check if the specified sheet exists
            if self.sheet_name not in xls.sheet_names:
                messagebox.showerror("Error", f"Sheet '{self.sheet_name}' does not exist in the Excel file.")
                return

            # Read the specified sheet into a DataFrame
            df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)

            # Check if the specified column exists
            if self.column_name not in df.columns:
                messagebox.showerror("Error",
                                     f"Column '{self.column_name}' does not exist in the sheet '{self.sheet_name}'.")
                return

            # Add "Filtered" column if it does not exist
            if 'Filtered' not in df.columns:
                df['Filtered'] = None

            # Process each sheet in the Excel file except the input sheet
            for output_sheet_name in xls.sheet_names:
                if output_sheet_name == self.sheet_name:
                    continue

                # Filter the DataFrame based on the column and sheet name (partial match)
                filtered_df = df[df[self.column_name].str.contains(output_sheet_name, case=False, na=False)]

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
                with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    # Write the filtered DataFrame to a new sheet in the same Excel file
                    filtered_df.to_excel(writer, sheet_name=output_sheet_name, index=False)
                    # Write the updated original DataFrame back to its original sheet
                    df.to_excel(writer, sheet_name=self.sheet_name, index=False)

            messagebox.showinfo("Success", f"Auto-detection and copying completed successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelFilterApp(root)
    root.mainloop()
