from tkinter import ttk
import tkinter as tk
import tkinter.messagebox as tkmb
import openpyxl
from openpyxl import *
from datetime import datetime
from tkinter import filedialog
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from openpyxl.utils import range_boundaries
import docx
from docx.shared import Inches
import math
wb = openpyxl.load_workbook('Sizing_template.xlsx')
sheet = wb.active
wb2 = openpyxl.load_workbook('Standard costing sheet.xlsx', data_only=True)
sheet2 = wb2.active

root=ttkb.Window(themename="darkly")
root.title("Sizing System")
root.geometry("960x860")
root.resizable(False,False)

global ups_make_entry, ups_model_entry, ups_rating_entry, actual_loadkva_entry, actual_loadkw_entry
global power_factor_entry, inverter_efficiency_entry, nominal_dc_voltage_entry, backup_requirement_entry
global cell_chemisrty_combobox

def mainscreen():
    title_label = ttkb.Label(root, text="Sizing Software", font=("Segoe UI", 32, "bold"))
    title_label.place(relx=0.5, rely=0.1, anchor="center")
    def next():
        global customername
        customername=customer_entry.get()
        global providername
        providername=provider_entry.get()
        global date
        date=date_entry.get()
        customer_name = customer_entry.get()
        solution_provider = provider_entry.get()
        date = date_entry.get()
        if not customer_name or not solution_provider or not date:
            tkmb.showerror("Error", "Please fill in all fields.")
            return
        try:
            datetime.strptime(date, "%d-%m-%Y")
            input()
        except ValueError:
            tkmb.showerror("Error", "Date must be in DD-MM-YYYY format.")
            return
    global center_frame
    center_frame = ttkb.Frame(root)
    center_frame.place(relx=0.5, rely=0.5, anchor="center")
    customer_label = ttkb.Label(center_frame, text="Customer Name:", font=("Segoe UI", 14))
    customer_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")
    customer_entry = ttkb.Entry(center_frame, font=("Segoe UI", 14), width=30)
    customer_entry.grid(row=0, column=1, padx=10, pady=10)
    provider_label = ttkb.Label(center_frame, text="Solution Provider:", font=("Segoe UI", 14))
    provider_label.grid(row=1, column=0, padx=10, pady=10, sticky="e")
    provider_entry = ttkb.Entry(center_frame, font=("Segoe UI", 14), width=30)
    provider_entry.grid(row=1, column=1, padx=10, pady=10)
    date_label = ttkb.Label(center_frame, text="Date:", font=("Segoe UI", 14))
    date_label.grid(row=2, column=0, padx=10, pady=10, sticky="e")
    date_entry = ttkb.Entry(center_frame, font=("Segoe UI", 14), width=30)
    date_entry.insert(0, datetime.now().strftime("%d-%m-%Y"))
    date_entry.grid(row=2, column=1, padx=10, pady=10)
    next_button = ttkb.Button(center_frame, text="Next", command=lambda: next())
    next_button.grid(row=3, column=0, columnspan=2, pady=20)
    search_button = ttkb.Button(center_frame, text="Search", command=costing_screen)
    search_button.grid(row=4, column=0, columnspan=2, pady=10)

def input():
    global ups_make_entry, ups_model_entry, ups_rating_entry, actual_loadkva_entry, actual_loadkw_entry
    global power_factor_entry, inverter_efficiency_entry, nominal_dc_voltage_entry, backup_requirement_entry
    global cell_chemisrty_combobox
    global input_frame
    def size():
        global ups_make
        global ups_model
        global ups_rating
        global actual_loadkva
        global power_factor
        global inverter_efficiency
        global nominal_dc_voltage
        global backup_requirement
        global calc_load
        global actual_loadkw
        global maximum_charging_voltage
        global endcell_voltage
        global energy_required
        global capacity_required
        global noofcells
        global cell_chemisrty
        ups_make = ups_make_entry.get()
        ups_model = ups_model_entry.get()
        ups_rating = float(ups_rating_entry.get())
        actual_loadkva = float(actual_loadkva_entry.get())
        actual_loadkw = float(actual_loadkw_entry.get())
        power_factor = float(power_factor_entry.get())
        inverter_efficiency = float(inverter_efficiency_entry.get())
        nominal_dc_voltage = float(nominal_dc_voltage_entry.get())
        backup_requirement = float(backup_requirement_entry.get())
        cell_chemisrty = cell_chemisrty_combobox.get()
        if cell_chemisrty == "LFP":
            nominal_vdc = 3.2
        elif cell_chemisrty == "NPM":
            nominal_vdc = 3.6
        if actual_loadkw:
            calc_load = round((actual_loadkw/inverter_efficiency),1)
        elif actual_loadkva:
            calc_load = round(((actual_loadkva * power_factor)/inverter_efficiency),1)
        elif ups_rating:
            calc_load = round(((ups_rating * power_factor)/inverter_efficiency),1)
        if nominal_dc_voltage == 12:
            noofcells = 4
        elif nominal_dc_voltage == 24:
            noofcells = 8
        elif nominal_dc_voltage == 36:
            noofcells = 11
        elif nominal_dc_voltage == 48:
            noofcells = 15
        elif nominal_dc_voltage == 72:
            noofcells = 23
        elif nominal_dc_voltage == 96:
            noofcells = 30
        elif nominal_dc_voltage == 120:
            noofcells = 38
        elif nominal_dc_voltage == 144:
            noofcells = 45
        elif nominal_dc_voltage == 192:
            noofcells = 60
        elif nominal_dc_voltage == 240:
            noofcells = 75
        elif nominal_dc_voltage == 336:
            noofcells = 105
        elif nominal_dc_voltage == 360:
            noofcells = 112
        elif nominal_dc_voltage == 384:
            noofcells = 120
        elif nominal_dc_voltage == 408:
            noofcells = 128
        elif nominal_dc_voltage == 480:
            noofcells = 150
        elif nominal_dc_voltage == 512:
            noofcells = 160
        elif nominal_dc_voltage == 528:
            noofcells = 165
        elif nominal_dc_voltage == 576:
            noofcells = 180
        global maximum_charging_voltage_calc
        global endcell_voltage_calc
        global energy_required_calc
        global capacity_required_calc
        maximum_charging_voltage_calc = (noofcells* 3.6)
        endcell_voltage_calc = noofcells* 2.8
        energy_required_calc = (calc_load * backup_requirement) / 60
        capacity_required_calc = (energy_required_calc * 1000) / endcell_voltage_calc
        maximum_charging_voltage = maximum_charging_voltage_calc
        endcell_voltage = round(endcell_voltage_calc,1)
        energy_required = round(energy_required_calc,1)
        capacity_required = round(capacity_required_calc,1)
        input2()
    center_frame.destroy()
    input_frame = ttkb.Frame(root)
    input_frame.place(relx=0.5, rely=0.5, anchor="center")
    ups_make_label = ttkb.Label(input_frame, text="UPS Make:", font=("Segoe UI", 12))
    ups_make_label.grid(row=0, column=0, padx=10, pady=5, sticky="e")
    ups_make_entry = ttkb.Entry(input_frame, font=("Segoe UI", 12), width=30)
    ups_make_entry.grid(row=0, column=1, padx=10, pady=5)
    ups_model_label = ttkb.Label(input_frame, text="UPS Model:", font=("Segoe UI", 12))
    ups_model_label.grid(row=1, column=0, padx=10, pady=5, sticky="e")
    ups_model_entry = ttkb.Entry(input_frame, font=("Segoe UI", 12), width=30)
    ups_model_entry.grid(row=1, column=1, padx=10, pady=5)
    ups_rating_label = ttkb.Label(input_frame, text="UPS Rating (kVA):", font=("Segoe UI", 12))
    ups_rating_label.grid(row=2, column=0, padx=10, pady=5, sticky="e")
    ups_rating_entry = ttkb.Entry(input_frame, font=("Segoe UI", 12), width=30)
    ups_rating_entry.insert(0, "0")
    ups_rating_entry.grid(row=2, column=1, padx=10, pady=5)
    actual_loadkva_label = ttkb.Label(input_frame, text="Actual Load (kVA):", font=("Segoe UI", 12))
    actual_loadkva_label.grid(row=3, column=0, padx=10, pady=5, sticky="e")
    actual_loadkva_entry = ttkb.Entry(input_frame, font=("Segoe UI", 12), width=30)
    actual_loadkva_entry.insert(0, "0")
    actual_loadkva_entry.grid(row=3, column=1, padx=10, pady=5)
    actual_loadkw_label = ttkb.Label(input_frame, text="Actual Load (kW):", font=("Segoe UI", 12))
    actual_loadkw_label.grid(row=4, column=0, padx=10, pady=5, sticky="e")
    actual_loadkw_entry = ttkb.Entry(input_frame, font=("Segoe UI", 12), width=30)
    actual_loadkw_entry.insert(0, "0")
    actual_loadkw_entry.grid(row=4, column=1, padx=10, pady=5)
    power_factor_label = ttkb.Label(input_frame, text="Power Factor:", font=("Segoe UI", 12))
    power_factor_label.grid(row=5, column=0, padx=10, pady=5, sticky="e")
    power_factor_entry = ttkb.Entry(input_frame, font=("Segoe UI", 12), width=30)
    power_factor_entry.insert(0, "0")
    power_factor_entry.grid(row=5, column=1, padx=10, pady=5)
    inverter_efficiency_label = ttkb.Label(input_frame, text="Inverter Efficiency:", font=("Segoe UI", 12))
    inverter_efficiency_label.grid(row=6, column=0, padx=10, pady=5, sticky="e")
    inverter_efficiency_entry = ttkb.Entry(input_frame, font=("Segoe UI", 12), width=30)
    inverter_efficiency_entry.insert(0, "0")
    inverter_efficiency_entry.grid(row=6, column=1, padx=10, pady=5)
    nominal_dc_voltage_label = ttkb.Label(input_frame, text="Nominal DC Voltage (V):", font=("Segoe UI", 12))
    nominal_dc_voltage_label.grid(row=7, column=0, padx=10, pady=5, sticky="e")
    nominal_dc_voltage_entry = ttkb.Entry(input_frame, font=("Segoe UI", 12), width=30)
    nominal_dc_voltage_entry.insert(0, "0")
    nominal_dc_voltage_entry.grid(row=7, column=1, padx=10, pady=5)
    backup_requirement_label = ttkb.Label(input_frame, text="Backup Requirement (Min):", font=("Segoe UI", 12))
    backup_requirement_label.grid(row=8, column=0, padx=10, pady=5, sticky="e")
    backup_requirement_entry = ttkb.Entry(input_frame, font=("Segoe UI", 12), width=30)
    backup_requirement_entry.insert(0, "0")
    backup_requirement_entry.grid(row=8, column=1, padx=10, pady=5)
    cell_chemisrty_label = ttkb.Label(input_frame, text="Cell Chemistry:", font=("Segoe UI", 12))
    cell_chemisrty_label.grid(row=9, column=0, padx=10, pady=5, sticky="e")
    cell_chemisrty_var = tk.StringVar()
    cell_chemisrty_combobox = ttkb.Combobox(input_frame, textvariable=cell_chemisrty_var, font=("Segoe UI", 12), width=28, state="readonly")
    cell_chemisrty_combobox['values'] = ("LFP", "NPM")
    cell_chemisrty_combobox.current(0)
    cell_chemisrty_combobox.grid(row=9, column=1, padx=10, pady=5)
    size_button = ttkb.Button(input_frame, text="Size", width=20,command=lambda: size())
    size_button.grid(row=10, column=0, columnspan=2, pady=20)

def input2():
    global input_frame2
    global offered_battery_config_entry
    global calc_load_entry
    global noofcells_entry
    global maximum_charging_voltage_entry
    global endcell_voltage_entry
    global minimum_charging_current_entry
    global maximum_charging_current_entry
    global energy_required_entry
    global capacity_required_entry
    global nearest_available_capacity_entry
    global total_available_energy_entry
    global backup_time_entry

    input_frame.destroy()
    input_frame2 = ttkb.Frame(root)
    input_frame2.place(relx=0.5, rely=0.5, anchor="center")
    calc_load_label = ttkb.Label(input_frame2, text="Calculated Load in kW:", font=("Segoe UI", 12))
    calc_load_label.grid(row=0, column=0, padx=10, pady=5, sticky="e")
    calc_load_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    calc_load_entry.insert(0,calc_load)
    calc_load_entry.grid(row=0, column=1, padx=10, pady=5)
    noofcells_label = ttkb.Label(input_frame2, text="Number of Cells:", font=("Segoe UI", 12))
    noofcells_label.grid(row=2, column=0, padx=10, pady=5, sticky="e")
    noofcells_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    noofcells_entry.insert(0,noofcells)
    noofcells_entry.grid(row=2, column=1, padx=10, pady=5)
    maximum_charging_voltage_label = ttkb.Label(input_frame2, text="Maximum Charging Voltage(V):", font=("Segoe UI", 12))
    maximum_charging_voltage_label.grid(row=3, column=0, padx=10, pady=5, sticky="e")
    maximum_charging_voltage_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    maximum_charging_voltage_entry.insert(0,maximum_charging_voltage)
    maximum_charging_voltage_entry.grid(row=3, column=1, padx=10, pady=5)
    endcell_voltage_label = ttkb.Label(input_frame2, text="End Cell Voltage(V):", font=("Segoe UI", 12))
    endcell_voltage_label.grid(row=4, column=0, padx=10, pady=5, sticky="e")
    endcell_voltage_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    endcell_voltage_entry.insert(0,endcell_voltage)
    endcell_voltage_entry.grid(row=4, column=1, padx=10, pady=5)
    minimum_charging_current_label = ttkb.Label(input_frame2, text="Minimum Charging Current:", font=("Segoe UI", 12))
    minimum_charging_current_label.grid(row=5, column=0, padx=10, pady=5, sticky="e")
    minimum_charging_current_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    minimum_charging_current_entry.grid(row=5, column=1, padx=10, pady=5)
    minimum_charging_current_entry.insert(0,"0.1C")
    maximum_charging_current_label = ttkb.Label(input_frame2, text="Maximum Charging Current:", font=("Segoe UI", 12))
    maximum_charging_current_label.grid(row=6, column=0, padx=10, pady=5, sticky="e")
    maximum_charging_current_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    maximum_charging_current_entry.grid(row=6, column=1, padx=10, pady=5)
    maximum_charging_current_entry.insert(0,"0.5C")
    energy_required_label = ttkb.Label(input_frame2, text="Energy Required(kWh):", font=("Segoe UI", 12))
    energy_required_label.grid(row=7, column=0, padx=10, pady=5, sticky="e")
    energy_required_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    energy_required_entry.insert(0,energy_required)
    energy_required_entry.grid(row=7, column=1, padx=10, pady=5)
    ageing_label = ttkb.Label(input_frame2, text="Ageing:", font=("Segoe UI", 12))
    ageing_label.grid(row=8, column=0, padx=10, pady=5, sticky="e")
    ageing_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    ageing_entry.insert(0, 0)
    ageing_entry.grid(row=8, column=1, padx=10, pady=5)
    def on_ageing_change(event):
        age_str = ageing_entry.get()
        try:
            age = float(age_str)
        except ValueError:
            capacity_required_entry.delete(0, tk.END)
            return
        capacity_required_calc = ((energy_required_calc * 1000) / endcell_voltage_calc) + ((energy_required_calc * 1000) / endcell_voltage_calc) * age
        capacity_required = round(capacity_required_calc, 1)
        capacity_required_entry.delete(0, tk.END)
        capacity_required_entry.insert(0, capacity_required)
        try:
            backup_time = (backup_requirement / capacity_required_calc) * float(nearest_available_capacity_entry.get())
            backup_time_entry.delete(0, tk.END)
            backup_time_entry.insert(0, round(backup_time,1))
        except ValueError:
            backup_time_entry.delete(0, tk.END)
            return
    ageing_entry.bind("<KeyRelease>", on_ageing_change)
    capacity_required_label = ttkb.Label(input_frame2, text="Capacity Required(Ah):", font=("Segoe UI", 12))
    capacity_required_label.grid(row=9, column=0, padx=10, pady=5, sticky="e")
    capacity_required_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    capacity_required_entry.insert(0,capacity_required)
    capacity_required_entry.grid(row=9, column=1, padx=10, pady=5)
    nearest_available_capacity_label = ttkb.Label(input_frame2, text="Nearest Available Capacity(Ah):", font=("Segoe UI", 12))
    nearest_available_capacity_label.grid(row=10, column=0, padx=10, pady=5, sticky="e")
    nearest_available_capacity_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    nearest_available_capacity_entry.grid(row=10, column=1, padx=10, pady=5)
    def on_nearest_capacity_change(event):
        global value_str
        global total_available_energy
        global backup_time
        value_str = nearest_available_capacity_entry.get()
        try:
            value = float(value_str)
        except ValueError:
            backup_time_entry.delete(0, tk.END)
            total_available_energy_entry.delete(0, tk.END)
            return
        backup_time = (backup_requirement / float(capacity_required_entry.get())) * value
        backup_time_entry.delete(0, tk.END)
        backup_time_entry.insert(0, math.floor(backup_time))
        total_available_energy = ((nominal_dc_voltage * value) / 1000)
        total_available_energy_entry.delete(0, tk.END)
        total_available_energy_entry.insert(0,total_available_energy)
        offered_battery_config=f"{int(nominal_dc_voltage)}V {int(value)}Ah"
        offered_battery_config_entry.delete(0, tk.END)
        offered_battery_config_entry.insert(0, offered_battery_config)
    nearest_available_capacity_entry.bind("<KeyRelease>", on_nearest_capacity_change)
    offered_battery_config_label = ttkb.Label(input_frame2, text="Offered Battery Configuration:", font=("Segoe UI", 12))
    offered_battery_config_label.grid(row=11, column=0, padx=10, pady=5, sticky="e")
    offered_battery_config_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    offered_battery_config_entry.grid(row=11, column=1, padx=10, pady=5)
    total_available_energy_label = ttkb.Label(input_frame2, text="Total Available Energy(kWh):", font=("Segoe UI", 12))
    total_available_energy_label.grid(row=12, column=0, padx=10, pady=5, sticky="e")
    total_available_energy_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    total_available_energy_entry.grid(row=12, column=1, padx=10, pady=5)
    backup_time_label = ttkb.Label(input_frame2, text="Backup Time(Aproximate Minutes):", font=("Segoe UI", 12))
    backup_time_label.grid(row=13, column=0, padx=10, pady=5, sticky="e")
    backup_time_entry = ttkb.Entry(input_frame2, font=("Segoe UI", 12), width=30)
    backup_time_entry.grid(row=13, column=1, padx=10, pady=5)
    def save_to_excel():
        if actual_loadkva:
            actual_load= actual_loadkva
        elif actual_loadkw:
            actual_load= actual_loadkw
        elif ups_rating:
            actual_load= 0
        data = {
            "Customer Name": customername,
            "Solution Provider": providername,
            "Date": date,
            "UPS Make": ups_make,
            "UPS Model": ups_model,
            "UPS Rating (KVA)": ups_rating,
            "Actual Load (KVA)": actual_load,
            "Power Factor": power_factor,
            "Inverter Efficiency": inverter_efficiency,
            "Nominal DC Voltage (V)": nominal_dc_voltage,
            "Backup Requirement (Min)": backup_requirement,
            "Cell Chemistry": cell_chemisrty,
            "Calculated Load (kW)": calc_load_entry.get(),
            "Maximum Charging Voltage(V)": maximum_charging_voltage_entry.get(),
            "End Cell Voltage(V)": endcell_voltage_entry.get(),
            "Minimum Charging Current": minimum_charging_current_entry.get(),
            "Maximum Charging Current": maximum_charging_current_entry.get(),
            "Energy Required(kWh)": energy_required_entry.get(),
            "Capacity Required(Ah)": capacity_required_entry.get(),
            "Nearest Available Capacity(Ah)": nearest_available_capacity_entry.get(),
            "Offered Battery Configuration": offered_battery_config_entry.get(),
            "Total Available Energy(kWh)": total_available_energy_entry.get(),
            "Backup Time(Aproximate minimum)": backup_time_entry.get()
        }   
        cell_map = {
            "Customer Name": "C4",
            "Solution Provider": "C5",
            "Date": "H3",
            "UPS Make": "A8",
            "UPS Model": "B6",
            "UPS Rating (KVA)": "C8",
            "Actual Load (KVA)": "D8",
            "Power Factor": "E8",
            "Inverter Efficiency": "F8",
            "Nominal DC Voltage (V)": "G8",
            "Backup Requirement (Min)": "H8",
            "Cell Chemistry": "E11",
            "Calculated Load (kW)": "E12",
            "Maximum Charging Voltage(V)": "E13",
            "End Cell Voltage(V)": "E14",
            "Minimum Charging Current": "E15",
            "Maximum Charging Current": "E16",
            "Energy Required(kWh)": "E17",
            "Capacity Required(Ah)": "E18",
            "Nearest Available Capacity(Ah)": "E19",
            "Offered Battery Configuration": "E20",
            "Total Available Energy(kWh)": "E21",
            "Backup Time(Aproximate minimum)": "E22"
        }
        merged_ranges = list(sheet.merged_cells.ranges)
        for key, cell in cell_map.items():
            merged = None
            for mrange in merged_ranges:
                if cell in mrange:
                    merged = mrange
                    break
            if merged:
                sheet.unmerge_cells(str(merged))
                sheet[cell] = data[key]
                sheet.merge_cells(str(merged))
            else:
                sheet[cell] = data[key]
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            wb.save(save_path)
            tkmb.showinfo("Success", "Data saved to Excel successfully.")
    save_button = ttkb.Button(input_frame2, text="Save to Excel", command=save_to_excel)
    save_button.grid(row=14, column=0, pady=20, sticky="e")
    clear_button = ttkb.Button(input_frame2, text="Make Costing", command=costing_screen)
    clear_button.grid(row=14, column=1, pady=20, sticky="w")

    def back_to_input():
        input_frame2.destroy()
        input()
        ups_make_entry.insert(0, ups_make)
        ups_model_entry.insert(0, ups_model)
        ups_rating_entry.insert(0, ups_rating)
        actual_loadkva_entry.insert(0, actual_loadkva)
        actual_loadkw_entry.insert(0, actual_loadkw)
        power_factor_entry.insert(0, power_factor)
        inverter_efficiency_entry.insert(0, inverter_efficiency)
        nominal_dc_voltage_entry.insert(0, nominal_dc_voltage)
        backup_requirement_entry.insert(0, backup_requirement)       
        if cell_chemisrty == "LFP":
            cell_chemisrty_combobox.current(0)
        else:
            cell_chemisrty_combobox.current(1)
        cell_chemisrty_combobox.grid(row=9, column=1, padx=10, pady=5)

    back_button = ttkb.Button(input_frame2, text="Back", command=back_to_input)
    back_button.grid(row=14, column=2, pady=20, sticky="w")

def costing_screen():
    global costing_frame
    def export_costing():
        template_wb = openpyxl.load_workbook('costing_sheet_template.xlsx')
        template_ws = template_wb.active
        option_count = 0
        for col in range(1, 4):
            if any(tree.set(item, f"Option {col}") for item in tree.get_children()):
                option_count += 1
        for col in range(2, 2 + option_count):
            template_ws.cell(row=3, column=col).value = backup_time_var.get()
            template_ws.cell(row=4, column=col).value = battery_config_entry.get()
        for idx, item_id in enumerate(tree.get_children()):
            excel_row = idx + 5
            if excel_row > 45:
                break
            values = tree.item(item_id, 'values')
            for col_offset, val in enumerate(values[1:1+option_count], start=2):
                template_ws.cell(row=excel_row, column=col_offset).value = val
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            template_wb.save(save_path)
            tkmb.showinfo("Success", "Costing exported successfully.")
    global offered_battery_config_value
    try:
        offered_battery_config_value = offered_battery_config_entry.get()
        input_frame2.destroy()
    except:
        offered_battery_config_value = ""
    try:
        center_frame.destroy()
    except:
        pass
    costing_frame = ttkb.Frame(root)
    costing_frame.place(relx=0.5, rely=0.5, anchor="center")
    backup_time_label = ttkb.Label(costing_frame, text="Select Backup Time:", font=("Segoe UI", 12))
    backup_time_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")
    backup_time_var = tk.StringVar(value="")  # Set default to empty
    battery_config_label = ttkb.Label(costing_frame, text="Battery Configuration:", font=("Segoe UI", 12))
    battery_config_label.grid(row=1, column=0, padx=10, pady=10, sticky="e")
    battery_config_entry = ttkb.Entry(costing_frame, font=("Segoe UI", 12), width=30)
    battery_config_entry.insert(0, offered_battery_config_value)
    battery_config_entry.grid(row=1, column=1, padx=10, pady=10)
    backup_time_combobox = ttkb.Combobox(costing_frame, textvariable=backup_time_var, font=("Segoe UI", 12), width=28, state="readonly")
    backup_time_combobox['values'] = ("15min", "30min", "60min", "120min")
    # Do NOT call backup_time_combobox.current(0)
    backup_time_combobox.grid(row=0, column=1, padx=10, pady=10)
    def on_backup_time_select(event):
        global sheet3
        selection = backup_time_var.get()
        sheet3 = wb2[selection]
        found_columns = []
        search_value = battery_config_entry.get().strip().lower()
        for col in range(1, sheet3.max_column + 1):
            cell_value = sheet3.cell(row=4, column=col).value
            if cell_value and str(cell_value).strip().lower() == search_value:
                col_letter = openpyxl.utils.get_column_letter(col)
                found_columns.append(col_letter)
        for idx in range(len(tree.get_children())):
            tree.set(tree.get_children()[idx], "Option 1", "")
            tree.set(tree.get_children()[idx], "Option 2", "")
            tree.set(tree.get_children()[idx], "Option 3", "")
        for option_idx, col_letter in enumerate(found_columns[:3]):
            for row_idx, excel_row in enumerate(range(5, 46)):
                value = sheet3[f"{col_letter}{excel_row}"].value
                if isinstance(value, (float, int)):
                    value = round(value, 2)
                tree.set(tree.get_children()[row_idx], f"Option {option_idx+1}", value if value is not None else "")
        highlight_excel_rows = [21, 34, 17, 40, 42]
        global prices, centretapping
        prices = []
        for option_idx, col_letter in enumerate(found_columns[:3]):
            price_cell_value = sheet3[f"{col_letter}40"].value
            prices.append(price_cell_value)
        centretapping=[]
        for option_idx, col_letter in enumerate(found_columns[:3]):
            centretapping_cell_value = sheet3[f"{col_letter}1"].value
            centretapping.append(centretapping_cell_value)
        highlight_tree_indices = [row - row_start for row in highlight_excel_rows]
        tree.tag_configure('highlight', background='yellow', foreground='black')
        for idx in highlight_tree_indices:
            item_id = tree.get_children()[idx]
            tree.item(item_id, tags=('highlight',))
    backup_time_combobox.bind("<<ComboboxSelected>>", on_backup_time_select)
    row_start = 5
    row_end = 45
    row_labels = []
    for i in range(row_start, row_end + 1):
        cell_value = sheet2[f"A{i}"].value
        row_labels.append(cell_value if cell_value is not None else "")
    columns = ("Description", "Option 1", "Option 2", "Option 3")
    tree = ttkb.Treeview(costing_frame, columns=columns, show="headings", height=15)
    # Set columns as wide as possible
    tree.heading("Description", text="Description")
    tree.column("Description", width=300, anchor="w", stretch=True)
    for col in columns[1:]:
        tree.heading(col, text=col)
        tree.column(col, width=200, anchor="center", stretch=True)
    for label in row_labels:
        tree.insert("", "end", values=(label, "", "", ""))
    vsb = ttkb.Scrollbar(costing_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    vsb.grid(row=2, column=4, sticky="ns")
    hsb = ttkb.Scrollbar(costing_frame, orient="horizontal", command=tree.xview)
    tree.configure(xscrollcommand=hsb.set)
    hsb.grid(row=3, column=0, columnspan=3, sticky="ew")
    tree.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
    costing_frame.grid_rowconfigure(1, weight=1)
    costing_frame.grid_columnconfigure(0, weight=1)
    costing_frame.grid_columnconfigure(1, weight=1)
    export_button = ttkb.Button(costing_frame, text="Export Costing", command=export_costing)
    export_button.grid(row=5, column=0, pady=20)
    def back_to_input2():
        costing_frame.destroy()
        input2()
        calc_load_entry.delete(0, tk.END)
        noofcells_entry.delete(0, tk.END)
        maximum_charging_voltage_entry.delete(0, tk.END)
        endcell_voltage_entry.delete(0, tk.END)
        minimum_charging_current_entry.delete(0, tk.END)
        maximum_charging_current_entry.delete(0, tk.END)
        energy_required_entry.delete(0, tk.END)
        capacity_required_entry.delete(0, tk.END)
        nearest_available_capacity_entry.delete(0, tk.END)
        total_available_energy_entry.delete(0, tk.END)
        backup_time_entry.delete(0, tk.END)
        calc_load_entry.insert(0, calc_load)
        noofcells_entry.insert(0, noofcells)
        maximum_charging_voltage_entry.insert(0, maximum_charging_voltage)
        endcell_voltage_entry.insert(0, endcell_voltage)
        minimum_charging_current_entry.insert(0, "0.1C")
        maximum_charging_current_entry.insert(0, "0.5C")
        energy_required_entry.insert(0, energy_required)
        capacity_required_entry.insert(0, capacity_required)
        nearest_available_capacity_entry.insert(0,value_str)
        offered_battery_config_entry.insert(0, offered_battery_config_value)
        total_available_energy_entry.insert(0, total_available_energy)
        backup_time_entry.insert(0, backup_time)

    button_frame = ttkb.Frame(costing_frame)
    button_frame.grid(row=4, column=1 ,pady=10)
    option1_button = ttkb.Button(button_frame, text="Option 1", command=option1)
    option1_button.grid(row=0, column=0, padx=5)
    option2_button = ttkb.Button(button_frame, text="Option 2", command=option2)
    option2_button.grid(row=0, column=1, padx=5)
    option3_button = ttkb.Button(button_frame, text="Option 3", command=option3)
    option3_button.grid(row=0, column=2, padx=5)
    make_quote_button = ttkb.Button(costing_frame, text="Make Quote", command=quotation)
    make_quote_button.grid(row=5, column=1, pady=20)
    back_button = ttkb.Button(costing_frame, text="Back", command=back_to_input2)
    back_button.grid(row=5, column=2, pady=20)

def quotation():
    quote_window = ttkb.Toplevel()
    quote_window.title("Quotation")
    global main_frame
    main_frame = ttkb.Frame(quote_window)
    main_frame.pack(expand=True, fill="both", padx=20, pady=20)
    global row_frames
    global text_areas
    row_frames = []
    text_areas = []
    def save_to_word():
        doc = docx.Document("Quote_format_High_Vtg.docx")
        if not doc.tables:
            tkmb.showerror("Error", "No table found in the template document.")
            return
        table = doc.tables[0]
        for row_idx, areas in enumerate(text_areas, start=1):
            while len(table.rows) <= row_idx:
                table.add_row()
            row_data = [area.get("1.0", tk.END).strip() for area in areas]
            if any(row_data):
                table.rows[row_idx].cells[0].text = str(row_idx)
            for col_idx, val in enumerate(row_data, start=1):
                if len(table.rows[row_idx].cells) > col_idx:
                    table.rows[row_idx].cells[col_idx].text = val
            # Calculate and set the product in the 6th column
            try:
                quantity = float(row_data[2]) if row_data[2] else 0
                price = float(row_data[3]) if row_data[3] else 0
                total = quantity * price
                if len(table.rows[row_idx].cells) > 5:
                    table.rows[row_idx].cells[5].text = str(round(total, 2))
            except Exception:
                if len(table.rows[row_idx].cells) > 5:
                    table.rows[row_idx].cells[5].text = ""
        save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if save_path:
            doc.save(save_path)
            tkmb.showinfo("Success", "Quotation saved to Word document.")

    # Create the button frame below main_frame
    button_frame = ttkb.Frame(quote_window)
    button_frame.pack(fill="x", padx=20, pady=(0, 20))
    save_button = ttkb.Button(button_frame, text="Save to Word", command=save_to_word)
    save_button.pack(padx=10, pady=10)

def option1():
    global option
    option=1
    add_row_frame()
def option2():
    global option
    option=2
    add_row_frame()
def option3():
    global option
    option=3
    add_row_frame()

def add_row_frame():
        row_frame = ttkb.Frame(main_frame)
        row_frame.pack(side="top", fill="x", pady=10)

        # Sr No label
        sr_no = len(row_frames) + 1
        sr_no_label = ttkb.Label(row_frame, text=f"{sr_no}", width=4, anchor="center")
        sr_no_label.pack(side="left", padx=(0, 5), pady=5)

        col1_frame = ttkb.Frame(row_frame)
        col1_frame.pack(side="left", expand=True, fill="both", padx=10)
        label1 = ttkb.Label(col1_frame, text="System")
        label1.pack(side="top", anchor="w")
        text_area1 = tk.Text(col1_frame, height=5, width=25, wrap="word")
        text_area1.pack(side="top", fill="both", expand=True)
        col2_frame = ttkb.Frame(row_frame)
        col2_frame.pack(side="left", expand=True, fill="both", padx=10)
        label2 = ttkb.Label(col2_frame, text="Solution")
        label2.pack(side="top", anchor="w")
        text_area2 = tk.Text(col2_frame, height=5, width=25, wrap="word")
        text_area2.pack(side="top", fill="both", expand=True)
        col3_frame = ttkb.Frame(row_frame)
        col3_frame.pack(side="left", expand=True, fill="both", padx=10)
        label3 = ttkb.Label(col3_frame, text="Quantity")
        label3.pack(side="top", anchor="w")
        text_area3 = tk.Text(col3_frame, height=5, width=25, wrap="word")
        text_area3.pack(side="top", fill="both", expand=True)
        col4_frame = ttkb.Frame(row_frame)
        col4_frame.pack(side="left", expand=True, fill="both", padx=10)
        label4 = ttkb.Label(col4_frame, text="Price")
        label4.pack(side="top", anchor="w")
        text_area4 = tk.Text(col4_frame, height=5, width=25, wrap="word")
        text_area4.pack(side="top", fill="both", expand=True)
        row_frames.append(row_frame)
        text_areas.append((text_area1, text_area2, text_area3, text_area4))
        text_area2.insert("1.0", f"Solution1: Lithium Battery Pack\n(HVL {offered_battery_config_value}) with\n Approximate Backup Time: {math.floor(backup_time)}Mins At BOL \n With Cabinet and inbuilt BMS")
        text_area3.insert("1.0", "1")
        if option == 1:
            if centretapping[0] == "centre tap":
                centretapping_text = "With Centre Tapping"
            else:
                centretapping_text = "Without Centre Tapping"
            text_area1.insert("1.0", f"{ups_rating}KVA: {backup_requirement}Min Backup \n(Load: {calc_load}kW)\n(Cell type:)\n({centretapping_text})")
            text_area4.insert("1.0", round(prices[0],2))
        elif option == 2:
            if centretapping[1] == "centre tap":
                centretapping_text = "With Centre Tapping"
            else:
                centretapping_text = "Without Centre Tapping"
            text_area1.insert("1.0", f"{ups_rating}KVA: {backup_requirement}Min Backup \n(Load: {calc_load}kW)\n(Cell type:)\n({centretapping_text})")
            text_area4.insert("1.0", round(prices[1],2))
        elif option == 3:
            if centretapping[2] == "centre tap":
                centretapping_text = "With Centre Tapping"
            else:
                centretapping_text = "Without Centre Tapping"
            text_area1.insert("1.0", f"{ups_rating}KVA: {backup_requirement}Min Backup \n(Load: {calc_load}kW)\n(Cell type:)\n({centretapping_text})")
            text_area4.insert("1.0", round(prices[2],2))

mainscreen()
root.mainloop()