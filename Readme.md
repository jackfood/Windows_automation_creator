## Windows Automation GUI

Windows Automation GUI is a powerful, user-friendly application for creating and executing automated tasks on Windows systems. Built with Python, it offers an intuitive interface for recording UI actions, generating scripts, and streamlining repetitive workflows.

## 🌟 Features

- **Smart UI Element Detection:** Automatically captures properties of UI elements under the cursor.
- **Versatile Action Support:** 
    - Left/Right/Double clicks
    - Text input and multi-line text entry
    - Special key commands and modifier key combinations
    - Program and file launching
- **Precise Timing Control:** Set custom wait times between actions.
- **One-Click Script Generation:** Instantly create executable Python scripts.
- **Import/Export Functionality:** Save and load automation sequences.
- **Visual Element Highlighter:** Optionally highlight targeted UI elements during execution.
- **Interactive Step Editor:** Review and modify steps in a user-friendly listbox.
- **Flexible Execution Options:** Run entire sequences or individual steps.

## 🚀 Quick Start

1. **Prerequisites:**
   - Python 3.7+
   - Required libraries: `tkinter`, `ttkbootstrap`, `pyautogui`, `uiautomation`, `keyboard`, `pywin32`

2. **Installation:**

   Copy the Win_automation_GUI.py
   ```bash
   pip install ttkbootstrap pyautogui uiautomation keyboard pywin32
   ```

3. **Launch the Application:**

   ```python
   python Win_automation_GUI.py
   ```

4. **Create Your First Automation:**
   - Hover over UI elements to capture their properties.
   - Select an action type and configure its parameters.
   - Click "Add Step" to build your sequence.
   - Generate and save your automation script (defaulted location on user's desktop as Win_automation.py)

5. **Run Your Automation:**

   ```python
   python Win_automation.py
   ```

## 📖 Detailed Usage Guide

**UI Navigation:**

- Use the top section to view and select UI element properties.
- Choose actions from the dropdown menu.
- Configure action-specific parameters (e.g., text input, special keys).

**Building Sequences:**

- Add steps using the "Add Step" button or Shift hotkey.
- Review and edit steps in the listbox.
- Adjust wait times for precise control.

**Script Management:**

- Generate scripts with the "Generate Code" button.
- Import existing scripts for modification.
- Scripts are saved as `Win_automation.py` on your desktop.

**Execution Options:**

- Toggle mouse movement and highlighting for visual feedback.
- Adjust execution speed by modifying wait times.

## 🤝 Contributing

We welcome contributions to enhance Windows Automation GUI! Here's how you can help:

1. Fork the repository.
2. Create a feature branch (`git checkout -b amazing-feature`).
3. Commit your changes (`git commit -m 'Add amazing feature'`).
4. Push to the branch (`git push origin amazing-feature`).
5. Open a Pull Request.

## 🙏 Acknowledgments

- [PyAutoGUI](https://pyautogui.readthedocs.io/en/latest/) for cross-platform GUI automation.
- [UIAutomation](https://pywinauto.readthedocs.io/en/latest/code/pywinauto.uia_defines.html) for Windows UI interaction.
- All our contributors and users who make this project possible!
- Claude.ai.

## 📌 Note

This project is created from Claude.ai Sonnet 3.5.
