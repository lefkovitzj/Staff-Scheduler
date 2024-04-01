"""
    Author: AaronTook (https://AaronTook.github.io)
    File Last Modified: 3/31/2024
    Project Name: Staff Scheduler
    File Name: gui_app.py
"""

# Python Standard Library module imports.
from tkinter import *
from tkinter.ttk import *

# Project module imports.
from utils import gui_get_file
from wrapper import SchedulerApplication

def clicked():
    if input_fp_textbox.get():
        in_fp = input_fp_textbox.get()
    else:
        input_fp_button.focus_set()
        return
    
    if lg_spin.get():
        num_guards = int(lg_spin.get())
    else:
        lg_spin.focus_set()
        return
    
    if mg_spin.get():
        num_managers = int(mg_spin.get())
    else:
        mg_spin.get()
        return
    
    if pa_spin.get():
        num_attendants = int(pa_spin.get())
    else:
        pa_spin.focus_set()
        return
    
    if max_consecutive_spin.get():
        max_consecutive = int(max_consecutive_spin.get())
    else:
        max_consecutive_spin.focus_set()
        return
    
    if output_fp_textbox.get():
        out_fp = output_fp_textbox.get().replace(".xlsx", "") + ".xlsx"
    else:
        output_fp_textbox.focus_set()
        return
    
    debug_mode = chk_state.get()
    
    scheduler = SchedulerApplication(in_fp, out_fp, debug_mode)
    scheduler.set_per_shift(guards=num_guards, attendants=num_attendants, managers=num_managers)
    scheduler.set_max_consecutive_shifts(max_consecutive)
    scheduler.run()
    window.destroy()
    
def pick_input_fp():
    
    input_fp_textbox.configure(state="normal")
    input_fp_textbox.delete(0, END)
    input_fp_textbox.insert(0, gui_get_file(limit_filetypes=[("Excel files", ".xlsx")])[0])
    input_fp_textbox.configure(state="disabled")

window = Tk()
window.title("Staff Scheduler")

window.geometry('380x240')
window.resizable(width=False, height=False)

chk_state = BooleanVar()
chk_state.set(False) #set check state


# Spinboxes in a separate frame.
spin_frame = Frame(window)
spin_frame.grid(column=0, row=2, rowspan=5, columnspan=3)

lg_num = IntVar()
mg_num = IntVar()
pa_num = IntVar()
max_consecutive_num = IntVar()

lg_num.set(5)
mg_num.set(1)
pa_num.set(1)
max_consecutive_num.set(3)

lg_spin = Spinbox(spin_frame, from_=1, to=10, width=5, textvariable=lg_num)
mg_spin = Spinbox(spin_frame, from_=1, to=10, width=5, textvariable=mg_num)
pa_spin = Spinbox(spin_frame, from_=1, to=10, width=5, textvariable=pa_num)
max_consecutive_spin = Spinbox(spin_frame, from_=1, to=10, width=5, textvariable=max_consecutive_num)

section_label = Label(spin_frame, text="\nStaff members per shift:", justify="left", anchor="w")
lg_label = Label(spin_frame, text="Lifeguards", justify="left", anchor="w")
mg_label = Label(spin_frame, text="Managers", justify="left", anchor="w")
pa_label = Label(spin_frame, text="Pool Attendants", justify="left", anchor="w")
consecutive_label = Label(spin_frame, text="Consecutive Shifts ", justify="left", anchor="w")

section_label.grid(sticky=W, column=0, row=1, columnspan=3)

lg_label.grid(sticky = W, column=1, row=2)
lg_spin.grid(sticky = W, column=0,row=2)

mg_label.grid(sticky = W, column=1, row=3)
mg_spin.grid(sticky = W, column=0,row=3)

pa_label.grid(sticky = W, column=1, row=4)
pa_spin.grid(sticky = W, column=0,row=4)

consecutive_label.grid(sticky = W, column=1, row=5)
max_consecutive_spin.grid(sticky = W, column=0, row=5)

# Input file widgets in a separate frame.
input_frame = Frame(window)
input_frame.grid(column=0, row=0, rowspan=2, columnspan=3)

input_fp_textbox = Entry(input_frame, width=50, state="disabled")
input_fp_button = Button(input_frame, text="Choose File", command=pick_input_fp)

input_label = Label(input_frame, text="Input file (.xlsx): ", justify="left", anchor="w")

input_label.grid(column=0, row=0, columnspan=3)
input_fp_textbox.grid(column=0, row=1, columnspan=2)
input_fp_button.grid(column=2, row=1)

# Output file widgets in a separate frame.
output_frame = Frame(window)
output_frame.grid(column=0, row=7, rowspan=2, columnspan=3)

output_label = Label(output_frame, text="\nOutput file (.xlsx): ", justify="left", anchor="w")
output_fp_textbox = Entry(output_frame,width=62)

output_label.grid(column=0, row=6, columnspan=3)
output_fp_textbox.grid(column=0, row=7, columnspan=3)

# Launch widgets in a separate frame.
launch_frame = Frame(window)
launch_frame.grid(column=0, row=10, rowspan=1, columnspan=3)


debug_checkbox = Checkbutton(launch_frame, text='Debug Mode', var=chk_state)
start_button = Button(launch_frame, text="Run", command=clicked)

debug_checkbox.grid(column=0, row=10)
start_button.grid(column=2,  row=10)


window.mainloop()