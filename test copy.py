import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from datetime import datetime
import pandas as pd

def select_excel_file():
    file_path = filedialog.askopenfilename(initialdir=".", title="Select Excel File", filetypes=(("Excel", "*.xlsx"), ("Excel", "*.xls")))
    if file_path:
        read_excel(file_path)

def read_excel(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Group values in column B based on categories in column A
        grouped_data = df.groupby('Service')['Container Name'].apply(list).reset_index()

        # Call the function to select surgery service
        select_surgery_service(grouped_data)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file: {e}")

def select_surgery_service(grouped_data):
    service_window = tk.Toplevel(root)
    service_window.title("Select Service")

    # Set the size of the window
    service_window.geometry("400x300")

    selected_services = tk.Listbox(service_window, selectmode=tk.MULTIPLE)
    selected_services.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

    services = grouped_data['Service']  # Access 'Service' column directly
    # Insert service options into the listbox
    for service in services:
        selected_services.insert(tk.END, service)

    def select():
        # Retrieve the indices of selected items
        selected_indices = selected_services.curselection()
        # Retrieve the selected services based on the indices
        selected = [selected_services.get(idx) for idx in selected_indices]
        # Pass the selected services to the next function
        select_instruments(grouped_data, selected)

    select_button = tk.Button(service_window, text="Select", command=select)
    select_button.pack(padx=10, pady=10)

def select_instruments(grouped_data, selected_services):
    container_window = tk.Toplevel(root)
    container_window.title("Select Containers")

    # Set the size of the window
    container_window.geometry("800x600")

    selected_containers = tk.Listbox(container_window, selectmode=tk.MULTIPLE)
    selected_containers.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

    for selected_service in selected_services:
        # Get the index of the row where 'Service' matches selected_service
        service_index = grouped_data[grouped_data['Service'] == selected_service].index.tolist()
        # Retrieve container names for the specified service index
        container_names = grouped_data.loc[service_index, 'Container Name'].iloc[0]
        # Insert each container name into the listbox
        for container_name in container_names:
            selected_containers.insert(tk.END, f"{selected_service}: {container_name}")

    def export_to_excel():
        # Retrieve the indices of selected items
        selected_indices = selected_containers.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning", "No containers selected.")
            return

        # Retrieve the selected containers based on the indices
        selected_instruments = [selected_containers.get(idx) for idx in selected_indices]

        # Create a DataFrame to store selected instruments
        selected_df = pd.DataFrame([instrument.split(': ') for instrument in selected_instruments],
                                   columns=['Service', 'Container Name'])

        # Prompt user for doctor's name
        doctor_name = simpledialog.askstring("Enter Doctor's Name", "Please enter the doctor's name:")
        if doctor_name is not None:
            # Create the file path based on doctor's name
            file_path = f"{doctor_name}.xlsx"

            # Check if the file already exists
            if os.path.isfile(file_path):
                # File exists, ask user if they want to append to existing file
                choice = messagebox.askyesno("File Exists", f"File '{doctor_name}.xlsx' already exists. Do you want to append to this file?")
                if choice:
                    # Append to existing file
                    with pd.ExcelWriter(file_path, mode='a') as writer:
                        sheet_name = f"{selected_df['Service'].iloc[0]}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
                        selected_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    messagebox.showinfo("Export Successful", f"Selected instruments appended to '{sheet_name}' in '{doctor_name}.xlsx' successfully.")
                else:
                    # Ask for new file name
                    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                             filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")),
                                                             initialfile=f"{doctor_name}.xlsx",
                                                             title="Save As")
                    if file_path:
                        selected_df.to_excel(file_path, index=False)
                        messagebox.showinfo("Export Successful", "Selected instruments exported successfully.")
                        os.startfile(file_path)  # Open the saved Excel file
                    else:
                        messagebox.showwarning("Warning", "Operation canceled.")
            else:
                # File does not exist, ask for confirmation to create new file
                choice = messagebox.askyesno("File Not Found", f"File '{doctor_name}.xlsx' does not exist. Do you want to create a new file?")
                if choice:
                    sheet_name = f"{selected_df['Service'].iloc[0]}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
                    selected_df.to_excel(file_path, sheet_name=sheet_name, index=False)
                    messagebox.showinfo("Export Successful", f"Selected instruments exported to '{sheet_name}' in '{doctor_name}.xlsx' successfully.")
                    os.startfile(file_path)  # Open the saved Excel file
                else:
                    messagebox.showwarning("Warning", "Operation canceled.")

    export_button = tk.Button(container_window, text="Export to Excel", command=export_to_excel)
    export_button.pack(padx=10, pady=10)

root = tk.Tk()
root.title("Surgery Instrument Selector")

select_button = tk.Button(root, text="Select Excel File", command=select_excel_file)
select_button.pack(padx=10, pady=10)

root.mainloop()
