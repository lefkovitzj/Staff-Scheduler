"""
    Author: AaronTook (https://AaronTook.github.io)
    File Last Modified: 3/31/2024
    Project Name: Staff Scheduler
    File Name: cmd_app.py
"""

# Python Standard Library module imports.
import traceback

# Project module imports.
from utils import gui_get_file
from wrapper import SchedulerApplication

try:
    print("--- Staff Scheduler ---")
    print("Please select an input (staff availability) .xlsx file: ", end="")
    in_fp = gui_get_file(limit_filetypes=[("Excel files", ".xlsx")])[0]
    print(f"\"{in_fp}\"")
    num_guards = int(input("Set a number of Lifeguards per shift: "))
    num_managers = int(input("Set a number of Managers per shift: "))
    num_attendants = int(input("Set a number of Pool Attendants per shift: "))
    max_consecutive = int(input("Set a maximum ideal consecutive number of shifts: "))
    out_fp = input("Enter the output (suggested schedule) filename: ").replace(".xlsx", "") + ".xlsx" # Get the filename, ensuring it ends with .xlsx regarless of if it was included by the user.
    scheduler = SchedulerApplication(in_fp, out_fp)
    scheduler.set_per_shift(guards=num_guards, attendants=num_attendants, managers=num_managers)
    scheduler.set_max_consecutive_shifts(max_consecutive)
    scheduler.run()
    print("Schedule successfully saved!")
except Exception as e:
    print(f"Error: application encountered unexpected crash. Full error message: \n{traceback.format_exc()}")

    