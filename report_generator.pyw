import tkinter as tk
from tkinter import messagebox
import pandas as pd
import datetime

def generate_summary_report():
    try:
        # Load all reports
        sales_df = pd.read_excel('sales_by_region_report.xlsx')
        pl_df = pd.read_excel('profit_&_loss_report.xlsx')
        inventory_df = pd.read_excel('inventory_levels_report.xlsx')
        timesheets_df = pd.read_excel('employee_timesheets_report.xlsx')
        interactions_df = pd.read_excel('customer_interactions_report.xlsx')
        budget_df = pd.read_excel('project_budget_tracking_report.xlsx')

        # Filter sales: sales > 100 in East & South regions
        filtered_sales = sales_df[(sales_df['Sales'] > 100) & (sales_df['Region'].isin(['East', 'South']))]

        # Profit & Loss summary sorted by amount descending
        pl_summary = pl_df.sort_values(by='Amount', ascending=False)

        # Inventory: products below reorder level
        low_stock = inventory_df[inventory_df['Stock Level'] < inventory_df['Reorder Level']]

        # Timesheets: employees working > 40 hours/week
        timesheets_df['Week'] = timesheets_df['Date'].dt.isocalendar().week
        hours_per_week = timesheets_df.groupby(['Employee', 'Week'])['Hours Worked'].sum().reset_index()
        overworked = hours_per_week[hours_per_week['Hours Worked'] > 40]

        # Customer interactions: "Follow up" in last 30 days
        cutoff_date = datetime.datetime.today() - pd.Timedelta(days=30)
        recent_followups = interactions_df[
            (interactions_df['Notes'] == 'Follow up') & (interactions_df['Date'] >= cutoff_date)
        ]

        # Budget tracking: projects > 90% budget used
        budget_df['Usage %'] = budget_df['Budget Used'] / budget_df['Budget Allocated'] * 100
        over_budget_projects = budget_df[budget_df['Usage %'] > 90]

        # Write summaries into one Excel file
        with pd.ExcelWriter('management_summary_report.xlsx') as writer:
            filtered_sales.to_excel(writer, sheet_name='Sales Summary', index=False)
            pl_summary.to_excel(writer, sheet_name='Profit & Loss', index=False)
            low_stock.to_excel(writer, sheet_name='Inventory Alerts', index=False)
            overworked.to_excel(writer, sheet_name='Overworked Employees', index=False)
            recent_followups.to_excel(writer, sheet_name='Customer Follow-ups', index=False)
            over_budget_projects.to_excel(writer, sheet_name='Budget Alerts', index=False)

        messagebox.showinfo("Success", "Summary report generated successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate report:\n{e}")

# Set up the simple GUI window with a button
root = tk.Tk()
root.title("Management Report Generator")
root.geometry("350x150")

btn = tk.Button(root, text="Generate Summary Report", command=generate_summary_report, padx=15, pady=10)
btn.pack(expand=True)

root.mainloop()
