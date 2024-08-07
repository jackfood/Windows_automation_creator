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
        self.root.title("Enhanced Windows Auto GUI v1.4 - Robust System improved detection")
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
        self.window_name_var = tk.StringVar()
        self.create_widgets()
        self.setup_global_hotkeys()
        self.update_info()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Import Script button at the top
        ttk.Button(main_frame, text="Import Script", command=self.import_script).grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="ew")

        # Window Name input
        ttk.Label(main_frame, text="Window Name:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(main_frame, textvariable=self.window_name_var, width=50).grid(row=1, column=1, padx=5, pady=2, sticky="ew")

        # ClassName, Name, AutomationId fields
        for i, field in enumerate(self.fields, start=2):
            ttk.Label(main_frame, text=f"{field}:").grid(row=i, column=0, sticky="w", padx=5, pady=2)
            ttk.Entry(main_frame, textvariable=self.vars[field], width=50).grid(row=i, column=1, padx=5, pady=2, sticky="ew")

        # Action dropdown
        ttk.Label(main_frame, text="Action:").grid(row=5, column=0, sticky="w", padx=5, pady=2)
        action_dropdown = ttk.Combobox(main_frame, textvariable=self.action_var, 
                                       values=["Click", "Right click", "Double click", "Special key", "Input text", "Input box", "Open Program"])
        action_dropdown.grid(row=5, column=1, sticky="ew", padx=5, pady=2)
        action_dropdown.bind("<<ComboboxSelected>>", self.on_action_selected)

        # Text Input
        ttk.Label(main_frame, text="Text Input:").grid(row=6, column=0, sticky="w", padx=5, pady=2)
        self.text_input_entry = ttk.Entry(main_frame, textvariable=self.text_input_var, width=50)
        self.text_input_entry.grid(row=6, column=1, padx=5, pady=2, sticky="ew")

        # Special Key
        ttk.Label(main_frame, text="Special Key:").grid(row=7, column=0, sticky="w", padx=5, pady=2)
        self.special_key_dropdown = ttk.Combobox(main_frame, textvariable=self.special_key_var, 
                                                 values=["F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12",
                                                         "Enter", "Esc", "Tab", "Backspace", "Delete", "Insert",
                                                         "Home", "End", "PageUp", "PageDown",
                                                         "Left", "Right", "Up", "Down",
                                                         "Windows", "Menu"])
        self.special_key_dropdown.grid(row=7, column=1, sticky="ew", padx=5, pady=2)

        # Modifier Key
        ttk.Label(main_frame, text="Modifier Key:").grid(row=8, column=0, sticky="w", padx=5, pady=2)
        modifier_frame = ttk.Frame(main_frame)
        modifier_frame.grid(row=8, column=1, sticky="ew", padx=5, pady=2)
        self.modifier_key_dropdown = ttk.Combobox(modifier_frame, textvariable=self.modifier_key_var, 
                                                  values=["", "Ctrl", "Alt", "Shift"], width=10)
        self.modifier_key_dropdown.pack(side=tk.LEFT)
        self.modifier_key_dropdown.bind("<<ComboboxSelected>>", self.on_modifier_key_selected)
        ttk.Label(modifier_frame, text="+").pack(side=tk.LEFT, padx=5)
        self.modifier_key_combo_entry = ttk.Entry(modifier_frame, textvariable=self.modifier_key_combo_var, width=5)
        self.modifier_key_combo_entry.pack(side=tk.LEFT)

        # Input box (multi-line text input)
        self.input_box_label = ttk.Label(main_frame, text="Input Box:")
        self.input_box_label.grid(row=9, column=0, sticky="w", padx=5, pady=2)
        self.input_box_label.grid_remove()
        self.input_box = tk.Text(main_frame, height=5, width=50)
        self.input_box.grid(row=9, column=1, padx=5, pady=2, sticky="ew")
        self.input_box.grid_remove()

        # Program path entry and browse button
        self.program_path_label = ttk.Label(main_frame, text="Program Path:")
        self.program_path_label.grid(row=10, column=0, sticky="w", padx=5, pady=2)
        self.program_path_label.grid_remove()
        program_path_frame = ttk.Frame(main_frame)
        program_path_frame.grid(row=10, column=1, sticky="ew", padx=5, pady=2)
        self.program_path_entry = ttk.Entry(program_path_frame, textvariable=self.program_path_var)
        self.program_path_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.browse_button = ttk.Button(program_path_frame, text="Browse", command=self.browse_program)
        self.browse_button.pack(side=tk.RIGHT)
        program_path_frame.grid_remove()

        # Wait Time
        ttk.Label(main_frame, text="Wait Time (s):").grid(row=11, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(main_frame, textvariable=self.wait_time_var, width=10).grid(row=11, column=1, sticky="w", padx=5, pady=2)

        # Insert position input
        ttk.Label(main_frame, text="Insert Position:").grid(row=12, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(main_frame, textvariable=self.insert_position_var, width=10).grid(row=12, column=1, sticky="w", padx=5, pady=2)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=13, column=0, columnspan=2, pady=10)
        ttk.Button(button_frame, text="Add Step", command=self.add_step).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Remove Step", command=self.remove_step).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Generate Code", command=self.code_generation).pack(side=tk.LEFT, padx=5)

        # Steps listbox
        self.steps_listbox = tk.Listbox(main_frame, width=80, height=10)
        self.steps_listbox.grid(row=14, column=0, columnspan=2, pady=10, sticky="ew")
        self.steps_listbox.bind('<<ListboxSelect>>', self.on_step_select)

        # Mouse movement checkbox
        self.allow_mouse_movement_checkbox = ttk.Checkbutton(
            main_frame, 
            text="Allow mouse movement and highlighter",
            variable=self.allow_mouse_movement_var,
            style='primary.TCheckbutton'
        )
        self.allow_mouse_movement_checkbox.grid(row=15, column=0, columnspan=2, pady=5, sticky="w")

        # Status label
        self.status_label = ttk.Label(main_frame, text="", foreground="black")
        self.status_label.grid(row=16, column=0, columnspan=2, pady=10, sticky="ew")

        # Configure column weights
        main_frame.columnconfigure(1, weight=1)

    def on_action_selected(self, event):
        action = self.action_var.get()
        if action == "Input box":
            self.input_box_label.grid()
            self.input_box.grid()
            self.text_input_entry.grid_remove()
            self.program_path_label.grid_remove()
            self.program_path_entry.master.grid_remove()
            self.special_key_dropdown.grid_remove()
            self.modifier_key_dropdown.master.grid_remove()
        elif action == "Open Program":
            self.input_box_label.grid_remove()
            self.input_box.grid_remove()
            self.text_input_entry.grid_remove()
            self.program_path_label.grid()
            self.program_path_entry.master.grid()
            self.special_key_dropdown.grid_remove()
            self.modifier_key_dropdown.master.grid_remove()
        elif action == "Special key":
            self.input_box_label.grid_remove()
            self.input_box.grid_remove()
            self.text_input_entry.grid_remove()
            self.program_path_label.grid_remove()
            self.program_path_entry.master.grid_remove()
            self.special_key_dropdown.grid()
            self.modifier_key_dropdown.master.grid()
        else:
            self.input_box_label.grid_remove()
            self.input_box.grid_remove()
            self.text_input_entry.grid()
            self.program_path_label.grid_remove()
            self.program_path_entry.master.grid_remove()
            self.special_key_dropdown.grid_remove()
            self.modifier_key_dropdown.master.grid_remove()

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
                
                # Get the top-level window
                window = element.GetTopLevelControl()
                self.window_name_var.set(window.Name if window else "")

                if not element.ClassName or element.ClassName == "Unknown":
                    parent = element.GetParentControl()
                    if parent:
                        self.vars["ClassName"].set(f"{parent.ClassName} > {element.ClassName or 'Unknown'}")
            else:
                self.vars["ClassName"].set("Unknown")
                self.vars["Name"].set("")
                self.vars["AutomationId"].set("")
                self.window_name_var.set("")
        except Exception as e:
            self.update_status(f"Error: {e}", "red")
        self.root.after(100, self.update_info)

    def add_step(self):
        action = self.action_var.get()
        step_info = {field: self.vars[field].get() for field in self.fields}
        step_info['action'] = action
        step_info['WindowName'] = self.window_name_var.get()
        
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
            window_name = step_info.get('WindowName', '')
            
            # Add more details based on the action type
            details = ""
            if action == "Input text":
                details = f"Text: {step_info.get('text', '')}"
            elif action == "Special key":
                modifier = step_info.get('modifier_key', '')
                key_combo = step_info.get('modifier_key_combo', '')
                special_key = step_info.get('special_key', '')
                if modifier and key_combo:
                    details = f"Key: {modifier} + {key_combo}"
                elif special_key:
                    details = f"Key: {special_key}"
            elif action == "Open Program":
                details = f"Path: {step_info.get('program_path', '')}"
            
            self.steps_listbox.insert(tk.END, f"{i}. {action}: {class_name}, {name}, Window: {window_name}, {details}, Wait: {wait_time}s")

    def on_step_select(self, event):
        selection = event.widget.curselection()
        if selection:
            index = selection[0]
            step_info, wait_time = self.steps[index]
            
            # Update form fields with selected step information
            for field in self.fields:
                self.vars[field].set(step_info.get(field, ''))
            
            self.window_name_var.set(step_info.get('WindowName', ''))
            self.action_var.set(step_info.get('action', ''))
            self.wait_time_var.set(wait_time)
            
            action = step_info.get('action', '')
            if action == "Input text":
                self.text_input_var.set(step_info.get('text', ''))
            elif action == "Input box":
                self.input_box.delete("1.0", tk.END)
                self.input_box.insert(tk.END, step_info.get('text', ''))
            elif action == "Special key":
                self.special_key_var.set(step_info.get('special_key', ''))
                self.modifier_key_var.set(step_info.get('modifier_key', ''))
                self.modifier_key_combo_var.set(step_info.get('modifier_key_combo', ''))
            elif action == "Open Program":
                self.program_path_var.set(step_info.get('program_path', ''))
            
            self.on_action_selected(None)  # Update form layout based on selected action

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

def find_desktop():
    desktop = auto.PaneControl(ClassName="WorkerW")
    if not desktop.Exists(1):
        desktop = auto.PaneControl(ClassName="Progman")
    if desktop.Exists(1):
        listview = desktop.ListControl(ClassName="SysListView32")
        if listview.Exists(1):
            return listview
    return None

def find_window(window_name):
    windows = auto.GetRootControl().GetChildren()
    for window in windows:
        if window_name.lower() in window.Name.lower():
            return window
    return None

def find_control(root_control, class_name, name=None, automation_id=None, max_depth=10):
    def search(control, depth=0):
        if depth > max_depth:
            return None

        if control.ClassName == "SysListView32" and name:
            for item in control.GetChildren():
                if item.Name.lower() == name.lower():
                    return item

        if class_name_matches(control, class_name):
            if (name is None or control.Name.lower() == name.lower()) and (automation_id is None or control.AutomationId == automation_id):
                return control

        for child in control.GetChildren():
            result = search(child, depth + 1)
            if result:
                return result

        return None

    def class_name_matches(ctrl, target_class):
        if target_class == "Unknown" or not target_class:
            return True
        if ">" in target_class:
            parent_class, child_class = target_class.split(" > ")
            return ctrl.ClassName == parent_class or ctrl.ClassName == child_class
        return (target_class.startswith("^") and re.match(target_class[1:], ctrl.ClassName)) or \
               (ctrl.ClassName == target_class) or \
               (target_class.lower() in ctrl.ClassName.lower())

    result = search(root_control)
    if result:
        print(f"Found matching element: ClassName={{result.ClassName}}, Name={{result.Name}}, AutomationId={{result.AutomationId}}")
    else:
        print("No matching element found")
    return result

def perform_action(element, action, step_info, window_name=None):
    if action == "Click":
        element.Click(simulateMove=ALLOW_MOUSE_MOVEMENT)
    elif action == "Right click":
        element.RightClick(simulateMove=ALLOW_MOUSE_MOVEMENT)
    elif action == "Double click":
        element.DoubleClick(simulateMove=ALLOW_MOUSE_MOVEMENT)
    elif action == "Input text":
        element.SendKeys(step_info.get("text", ""), interval=0)
    elif action == "Input box":
        lines = step_info.get("text", "").splitlines()
        for i, line in enumerate(lines):
            element.SendKeys(line, interval=0)
            if i < len(lines) - 1:
                element.SendKeys(uiautomation.Keys.VK_RETURN)
    elif action == "Special key":
        special_key = step_info.get("special_key", "")
        modifier_key = step_info.get("modifier_key", "")
        modifier_key_combo = step_info.get("modifier_key_combo", "")

        # Construct the key combination string
        key_combo = ""
        if modifier_key:
            key_combo += modifier_key.lower() + "+"
        if modifier_key_combo:
            key_combo += modifier_key_combo.lower() + "+"
        key_combo += special_key.lower()

        # Send the key combination using pyautogui
        pyautogui.hotkey(*key_combo.split("+"))
  
    elif action == "Open Program":
        program_path = step_info.get("program_path", "")
        if program_path:
            subprocess.Popen(program_path)
        else:
            print("No program path provided for 'Open Program' action")
    else:
        print(f"Unknown action: {{action}}")
    print(f"Action '{{action}}' performed on the element.")

def run_automation():
    steps = {steps_json}

    current_window = None

    for step_info, wait_time in steps:
        time.sleep(wait_time)

        action = step_info.get("action")
        class_name = step_info.get("ClassName")
        name = step_info.get("Name")
        automation_id = step_info.get("AutomationId")
        window_name = step_info.get("WindowName", "")

        print(f"Step: Action={{action}}, ClassName={{class_name}}, Name={{name}}, AutomationId={{automation_id}}, WindowName={{window_name}}")

        element = None

        if action == "Open Program":
            perform_action(None, action, step_info)
            continue

        if not window_name or window_name == "Program Manager":
            print("Searching for desktop items")
            desktop = find_desktop()
            if desktop:
                element = find_control(desktop, class_name, name, automation_id)
            else:
                print("Desktop not found")
        else:
            try:
                print(f"Searching for window: {{window_name}}")
                new_window = None
                for w in auto.WindowControl(searchDepth=1).GetChildren():
                    if window_name.lower() in w.Name.lower():
                        new_window = w
                        break

                if new_window and new_window.Exists(1):
                    if new_window != current_window:
                        current_window = new_window
                        print(f"Switching to window: {{current_window.Name}}")
                        current_window.SetFocus()
                        current_window.Maximize()
                    print(f"Searching for element within window: {{current_window.Name}}")
                    element = find_control(current_window, class_name, name, automation_id)
                else:
                    print(f"Window not found: {{window_name}}")

            except Exception as e:
                print(f"An error occurred while searching for the window: {{e}}")

        if not element:
            print("Element not found in specific context. Searching in the entire UI tree.")
            element = find_control(auto.GetRootControl(), class_name, name, automation_id)

        if element:
            highlight_element(element)
            perform_action(element, action, step_info, window_name)
        else:
            print(f"Element not found for step: {{step_info}}")


if __name__ == "__main__":
    run_automation()
"""

        try:
            with open(desktop_path, "w", encoding="utf-8") as f:
                f.write(script_content)
            self.update_status(f"Script saved to {desktop_path}", "green")
        except Exception as e:
            self.update_status(f"Failed to save script: {e}", "red")
    
    def update_status(self, message, color):
        self.status_label.config(text=message, foreground=color)
        self.root.after(3000, lambda: self.status_label.config(text="", foreground="black"))
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoGUIApp(root)
    app.run()
