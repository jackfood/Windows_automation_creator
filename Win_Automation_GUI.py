import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import uiautomation as auto
import pyautogui
import keyboard
import os
import time
import json
import re
import ast
from ttkbootstrap import Style, PRIMARY

class AutoGUIApp:
    def __init__(self, root):
        self.root = root
        self.style = Style(theme='flatly')
        self.root.title("Enhanced Windows Auto GUI v1.26")
        self.fields = ["ClassName", "Name", "AutomationId"]
        self.vars = {field: tk.StringVar() for field in self.fields}
        self.action_var = tk.StringVar(value="Click")
        self.text_input_var = tk.StringVar()
        self.special_key_var = tk.StringVar()
        self.modifier_key_var = tk.StringVar()
        self.modifier_key_combo_var = tk.StringVar()
        self.wait_time_var = tk.DoubleVar(value=1.0)
        self.program_path_var = tk.StringVar()
        self.allow_mouse_movement_var = tk.BooleanVar(value=False)
        self.insert_position_var = tk.StringVar()
        self.steps = []
        self.create_widgets()
        self.setup_global_hotkeys()
        self.update_info()

    def create_widgets(self):
        for i, field in enumerate(self.fields):
            tk.Label(self.root, text=f"{field}:").grid(row=i, column=0, sticky="w", padx=5, pady=2)
            tk.Entry(self.root, textvariable=self.vars[field], width=50).grid(row=i, column=1, padx=5, pady=2)

        tk.Label(self.root, text="Action:").grid(row=len(self.fields), column=0, sticky="w", padx=5, pady=2)
        action_dropdown = ttk.Combobox(self.root, textvariable=self.action_var, 
                                       values=["Click", "Right click", "Double click", "Special key", "Input text", "Input box", "Open Program"])
        action_dropdown.grid(row=len(self.fields), column=1, sticky="w", padx=5, pady=2)
        action_dropdown.bind("<<ComboboxSelected>>", self.on_action_selected)

        tk.Label(self.root, text="Text Input:").grid(row=len(self.fields)+1, column=0, sticky="w", padx=5, pady=2)
        self.text_input_entry = tk.Entry(self.root, textvariable=self.text_input_var, width=50)
        self.text_input_entry.grid(row=len(self.fields)+1, column=1, padx=5, pady=2)

        tk.Button(self.root, text="Import Script", command=self.import_script).grid(row=len(self.fields)+14, column=0, columnspan=2, pady=5)

        tk.Label(self.root, text="Special Key:").grid(row=len(self.fields)+2, column=0, sticky="w", padx=5, pady=2)
        self.special_key_dropdown = ttk.Combobox(self.root, textvariable=self.special_key_var, 
                                                 values=["F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12",
                                                         "Enter", "Esc", "Tab", "Backspace", "Delete", "Insert",
                                                         "Home", "End", "PageUp", "PageDown",
                                                         "Left", "Right", "Up", "Down",
                                                         "Windows", "Menu"])
        self.special_key_dropdown.grid(row=len(self.fields)+2, column=1, sticky="w", padx=5, pady=2)

        tk.Label(self.root, text="Modifier Key:").grid(row=len(self.fields)+3, column=0, sticky="w", padx=5, pady=2)
        self.modifier_key_dropdown = ttk.Combobox(self.root, textvariable=self.modifier_key_var, 
                                                  values=["", "Ctrl", "Alt", "Shift"])
        self.modifier_key_dropdown.grid(row=len(self.fields)+3, column=1, sticky="w", padx=5, pady=2)
        self.modifier_key_dropdown.bind("<<ComboboxSelected>>", self.on_modifier_key_selected)

        tk.Label(self.root, text="+").grid(row=len(self.fields)+3, column=1, padx=(120, 0))
        self.modifier_key_combo_entry = tk.Entry(self.root, textvariable=self.modifier_key_combo_var, width=5)
        self.modifier_key_combo_entry.grid(row=len(self.fields)+3, column=1, padx=(140, 0))
        self.modifier_key_combo_entry.config(state='disabled')

        # Input box (multi-line text input)
        self.input_box_label = tk.Label(self.root, text="Input Box:")
        self.input_box_label.grid(row=len(self.fields)+4, column=0, sticky="w", padx=5, pady=2)
        self.input_box_label.grid_remove()

        self.input_box = tk.Text(self.root, height=5, width=50)
        self.input_box.grid(row=len(self.fields)+4, column=1, padx=5, pady=2)
        self.input_box.grid_remove()

        # Program path entry and browse button
        self.program_path_label = tk.Label(self.root, text="Program Path:")
        self.program_path_label.grid(row=len(self.fields)+5, column=0, sticky="w", padx=5, pady=2)
        self.program_path_label.grid_remove()

        self.program_path_entry = tk.Entry(self.root, textvariable=self.program_path_var, width=40)
        self.program_path_entry.grid(row=len(self.fields)+5, column=1, sticky="w", padx=5, pady=2)
        self.program_path_entry.grid_remove()

        self.browse_button = ttk.Button(self.root, text="Browse", command=self.browse_program)
        self.browse_button.grid(row=len(self.fields)+5, column=1, sticky="e", padx=5, pady=2)
        self.browse_button.grid_remove()

        tk.Label(self.root, text="Wait Time (s):").grid(row=len(self.fields)+6, column=0, sticky="w", padx=5, pady=2)
        tk.Entry(self.root, textvariable=self.wait_time_var, width=10).grid(row=len(self.fields)+6, column=1, sticky="w", padx=5, pady=2)

        # Insert position input
        tk.Label(self.root, text="Insert Position:").grid(row=len(self.fields)+7, column=0, sticky="w", padx=5, pady=2)
        tk.Entry(self.root, textvariable=self.insert_position_var, width=10).grid(row=len(self.fields)+7, column=1, sticky="w", padx=5, pady=2)

        tk.Button(self.root, text="Add Step", command=self.add_step).grid(row=len(self.fields)+8, column=0, columnspan=2, pady=5)
        tk.Button(self.root, text="Remove Step", command=self.remove_step).grid(row=len(self.fields)+9, column=0, columnspan=2, pady=5)
        tk.Button(self.root, text="Generate Code", command=self.code_generation).grid(row=len(self.fields)+10, column=0, columnspan=2, pady=5)

        self.steps_listbox = tk.Listbox(self.root, width=80, height=10)
        self.steps_listbox.grid(row=len(self.fields)+11, column=0, columnspan=2, pady=10)

        # Add checkbox for mouse movement control
        self.allow_mouse_movement_checkbox = ttk.Checkbutton(
            self.root, 
            text="Allow mouse movement and highlighter",
            variable=self.allow_mouse_movement_var,
            style='primary.TCheckbutton'
        )
        self.allow_mouse_movement_checkbox.grid(row=len(self.fields)+12, column=0, columnspan=2, pady=5, sticky="w")

        self.status_label = tk.Label(self.root, text="", fg="black")
        self.status_label.grid(row=len(self.fields)+13, column=0, columnspan=2, pady=10)

    def on_action_selected(self, event):
        action = self.action_var.get()
        if action == "Input box":
            self.input_box_label.grid()
            self.input_box.grid()
            self.text_input_entry.grid_remove()
            self.program_path_label.grid_remove()
            self.program_path_entry.grid_remove()
            self.browse_button.grid_remove()
            self.special_key_dropdown.grid_remove()
            self.modifier_key_dropdown.grid_remove()
            self.modifier_key_combo_entry.grid_remove()
        elif action == "Open Program":
            self.input_box_label.grid_remove()
            self.input_box.grid_remove()
            self.text_input_entry.grid_remove()
            self.program_path_label.grid()
            self.program_path_entry.grid()
            self.browse_button.grid()
            self.special_key_dropdown.grid_remove()
            self.modifier_key_dropdown.grid_remove()
            self.modifier_key_combo_entry.grid_remove()
        elif action == "Special key":
            self.input_box_label.grid_remove()
            self.input_box.grid_remove()
            self.text_input_entry.grid_remove()
            self.program_path_label.grid_remove()
            self.program_path_entry.grid_remove()
            self.browse_button.grid_remove()
            self.special_key_dropdown.grid()
            self.modifier_key_dropdown.grid()
            self.modifier_key_combo_entry.grid()
        else:
            self.input_box_label.grid_remove()
            self.input_box.grid_remove()
            self.text_input_entry.grid()
            self.program_path_label.grid_remove()
            self.program_path_entry.grid_remove()
            self.browse_button.grid_remove()
            self.special_key_dropdown.grid_remove()
            self.modifier_key_dropdown.grid_remove()
            self.modifier_key_combo_entry.grid_remove()

        if action == "Input text":
            self.text_input_entry.config(state='normal')
        else:
            self.text_input_entry.config(state='disabled')

    def on_modifier_key_selected(self, event):
        if self.modifier_key_var.get() == "":
            self.modifier_key_combo_entry.config(state='disabled')
        else:
            self.modifier_key_combo_entry.config(state='normal')

    def browse_program(self):
        file_path = filedialog.askopenfilename(filetypes=[("All files", "*.*")])
        if file_path:
            self.program_path_var.set(file_path)

    def setup_global_hotkeys(self):
        keyboard.add_hotkey('shift', self.add_step)

    def import_script(self):
        file_path = filedialog.askopenfilename(filetypes=[("Python files", "*.py")])
        if file_path:
            try:
                with open(file_path, "r") as file:
                    content = file.read()
                
                # Find the steps data in the script
                match = re.search(r'steps = (\[[\s\S]*?\n\])', content, re.MULTILINE)
                if match:
                    steps_data_str = match.group(1)
                    try:
                        steps_data = ast.literal_eval(steps_data_str)
                        
                        # Clear existing steps
                        self.steps.clear()
                        
                        # Restore steps
                        for step_info, wait_time in steps_data:
                            self.steps.append((step_info, wait_time))
                        
                        self.update_steps_listbox()
                        self.update_status("Script imported successfully", "green")
                    except (SyntaxError, ValueError) as parse_error:
                        self.update_status(f"Error parsing steps data: {str(parse_error)}", "red")
                else:
                    self.update_status("Unable to find steps data in the script", "red")
            except Exception as e:
                self.update_status(f"Error importing script: {str(e)}", "red")

    def update_info(self):
        try:
            x, y = pyautogui.position()
            element = auto.ControlFromPoint(x, y)
            if element:
                self.vars["ClassName"].set(element.ClassName or "Unknown")
                self.vars["Name"].set(element.Name or "")
                self.vars["AutomationId"].set(element.AutomationId or "")
                
                if not element.ClassName or element.ClassName == "Unknown":
                    parent = element.GetParentControl()
                    if parent:
                        self.vars["ClassName"].set(f"{parent.ClassName} > {element.ClassName or 'Unknown'}")
            else:
                self.vars["ClassName"].set("Unknown")
                self.vars["Name"].set("")
                self.vars["AutomationId"].set("")
        except Exception as e:
            self.update_status(f"Error: {e}", "red")
        self.root.after(100, self.update_info)

    def add_step(self):
        action = self.action_var.get()
        step_info = {field: self.vars[field].get() for field in self.fields}
        step_info['action'] = action
        
        if action == "Input text":
            step_info['text'] = self.text_input_var.get()
        elif action == "Input box":
            step_info['text'] = self.input_box.get("1.0", tk.END).strip()
        elif action == "Special key":
            step_info['special_key'] = self.special_key_var.get()
            step_info['modifier_key'] = self.modifier_key_var.get()
            step_info['modifier_key_combo'] = self.modifier_key_combo_var.get()
        elif action == "Open Program":
            step_info['program_path'] = self.program_path_var.get()
    
        if action == "Open Program" or ((step_info['ClassName'] and step_info['ClassName'] != "Unknown") or step_info['Name']):
            insert_position = self.insert_position_var.get()
            if insert_position and insert_position.isdigit():
                position = int(insert_position) - 1
                self.steps.insert(position, (step_info, self.wait_time_var.get()))
            else:
                self.steps.append((step_info, self.wait_time_var.get()))
            self.update_steps_listbox()
            self.update_status("Step added successfully", "green")
        else:
            self.update_status("Incomplete information. Ensure either ClassName or Name is provided.", "red")

    def remove_step(self):
        selected = self.steps_listbox.curselection()
        if selected:
            index = selected[0]
            self.steps_listbox.delete(index)
            self.steps.pop(index)
            self.update_status("Step removed", "green")
        else:
            self.update_status("No step selected", "red")

    def update_steps_listbox(self):
        self.steps_listbox.delete(0, tk.END)
        for i, (step_info, wait_time) in enumerate(self.steps, start=1):
            action = step_info.get('action', '')
            class_name = step_info.get('ClassName', '')
            name = step_info.get('Name', '')
            self.steps_listbox.insert(tk.END, f"{i}. {action}: {class_name}, {name}, Wait: {wait_time}s")

    def code_generation(self):
        if not self.steps:
            self.update_status("No steps to save", "red")
            return
    
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "Win_automation.py")
        steps_json = json.dumps(self.steps, indent=4)
        allow_mouse_movement = self.allow_mouse_movement_var.get()
    
        script_content = f"""# -*- coding: utf-8 -*-
import os
import pyautogui
import uiautomation as auto
from uiautomation import Keys
import time
import re
import subprocess
import win32gui
import win32api
import win32con

ALLOW_MOUSE_MOVEMENT = {allow_mouse_movement}

def highlight_element(element):
    if ALLOW_MOUSE_MOVEMENT:
        rect = element.BoundingRectangle
        width = rect.width() if callable(rect.width) else rect.width
        height = rect.height() if callable(rect.height) else rect.height
        left, top = rect.left, rect.top
        right, bottom = left + width, top + height
        hdc = win32gui.GetDC(0)
        red_pen = win32gui.CreatePen(win32con.PS_SOLID, 2, win32api.RGB(255, 0, 0))
        old_pen = win32gui.SelectObject(hdc, red_pen)
        try:
            # Draw the rectangle
            win32gui.MoveToEx(hdc, left, top)
            win32gui.LineTo(hdc, right, top)
            win32gui.LineTo(hdc, right, bottom)
            win32gui.LineTo(hdc, left, bottom)
            win32gui.LineTo(hdc, left, top)
            time.sleep(0.3)
        finally:
            # Clean up
            win32gui.SelectObject(hdc, old_pen)
            win32gui.DeleteObject(red_pen)
            win32gui.ReleaseDC(0, hdc)
    else:
        print(f"Element bounds: Left={{element.BoundingRectangle.left}}, Top={{element.BoundingRectangle.top}}, "
              f"Width={{element.BoundingRectangle.width}}, Height={{element.BoundingRectangle.height}}")

def find_control(root_control, class_name, name=None, automation_id=None, max_depth=10):
    def search(control, depth=0):
        if depth > max_depth:
            return None

        if class_name == "Unknown" or not class_name:
            if (name and control.Name == name) or (automation_id and control.AutomationId == automation_id):
                return control
        elif ">" in class_name:
            parent_class, child_class = class_name.split(" > ")
            if control.ClassName == parent_class:
                for child in control.GetChildren():
                    if child.ClassName == child_class or child_class == "Unknown":
                        if (name is None or child.Name == name) and (automation_id is None or child.AutomationId == automation_id):
                            return child
        elif (class_name.startswith("^") and re.match(class_name[1:], control.ClassName)) or (control.ClassName == class_name):
            if (name is None or control.Name == name) and (automation_id is None or control.AutomationId == automation_id):
                return control

        for child in control.GetChildren():
            result = search(child, depth + 1)
            if result:
                return result

        return None

    return search(root_control)

def run_automation():
    steps = {steps_json}

    for step_info, wait_time in steps:
        time.sleep(wait_time)

        action = step_info.get("action")

        if action == "Open Program":
            program_path = step_info.get("program_path")
            try:
                if program_path.lower().endswith(('.docx', '.xlsx', '.pptx', '.pdf', '.txt')):
                    os.startfile(program_path)
                else:
                    subprocess.Popen(program_path)
                print(f"Opened program/file: {{program_path}}")
            except Exception as e:
                print(f"Failed to open program/file: {{program_path}}. Error: {{e}}")
        else:
            class_name = step_info.get("ClassName")
            name = step_info.get("Name")
            automation_id = step_info.get("AutomationId")

            print(f"Searching for element: ClassName={{class_name}}, Name={{name}}, AutomationId={{automation_id}}")

            element = find_control(auto.GetRootControl(), class_name, name, automation_id)
            if element:
                print(f"Element found: ClassName={{element.ClassName}}, Name={{element.Name}}, AutomationId={{element.AutomationId}}")
                element.SetFocus()
                highlight_element(element)
                if action == "Click":
                    element.Click(simulateMove=ALLOW_MOUSE_MOVEMENT)
                elif action == "Right click":
                    element.RightClick(simulateMove=ALLOW_MOUSE_MOVEMENT)
                elif action == "Double click":
                    element.DoubleClick(simulateMove=ALLOW_MOUSE_MOVEMENT)
                elif action == "Input text":
                    element.SendKeys(step_info.get("text"), interval=0)
                elif action == "Input box":
                    lines = step_info.get("text").splitlines()
                    for i, line in enumerate(lines):
                        element.SendKeys(line, interval=0)
                        if i < len(lines) - 1:  # Don't press Enter after the last line
                            element.SendKeys(Keys.ENTER)
                elif action == "Special key":
                    special_key = step_info.get("special_key")
                    modifier_key = step_info.get("modifier_key")
                    modifier_key_combo = step_info.get("modifier_key_combo")
                    if modifier_key and modifier_key_combo:
                        element.SendKeys(f'{{{{{{modifier_key}}}}}}{{{{{{modifier_key_combo}}}}}}', interval=0)
                    elif special_key:
                        element.SendKeys(f'{{{{{{special_key.upper()}}}}}}', interval=0)
                print(f"Action '{{action}}' performed on the element.")
            else:
                print(f"Element not found: {{step_info}}")

if __name__ == "__main__":
    run_automation()
"""

        with open(desktop_path, "w") as f:
            f.write(script_content)
    
        self.update_status(f"Script saved to {desktop_path}", "green")

    def update_status(self, message, color):
        self.status_label.config(text=message, fg=color)
        self.root.after(3000, lambda: self.status_label.config(text="", fg="black"))

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoGUIApp(root)
    app.run()
