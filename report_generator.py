import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox

def generate_reports():
    try:
        # Excel file assumed in the same folder as this script or executable
        base_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(base_dir, 'sales_data.xlsx')

        df = pd.read_excel(file_path)
        filtered_sales = df[df['Sales'] > 100]
        filtered_region = df[df['Region'] == "East"]

        filtered_sales.to_excel(os.path.join(base_dir, 'filtered_sales_over_100.xlsx'), index=False)
        filtered_region.to_excel(os.path.join(base_dir, 'filtered_region_east.xlsx'), index=False)

        messagebox.showinfo("Success", "Reports saved successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate reports:\n{e}")

root = tk.Tk()
root.title("Report Generator")

# Set window size and center it on screen
window_width, window_height = 300, 120
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_cordinate = int((screen_width/2) - (window_width/2))
y_cordinate = int((screen_height/2) - (window_height/2))
root.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")

btn = tk.Button(root, text="Generate Reports", command=generate_reports, width=25, height=3)
btn.pack(padx=20, pady=20)

root.mainloop()
