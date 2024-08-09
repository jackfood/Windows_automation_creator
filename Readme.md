# Enhanced Windows Auto GUI v1.44

## Description

Enhanced Windows Auto GUI is a powerful tool for automating Windows GUI interactions. It provides a user-friendly interface for recording, editing, and executing complex sequences of GUI operations across various Windows applications.

## Features

1. **Advanced GUI Element Detection**
   - Real-time detection of GUI elements under the mouse cursor
   - Captures ClassName, Name, and AutomationId properties for precise element identification
   - Supports complex class hierarchies (e.g., "ParentClass > ChildClass")
   - Fallback mechanisms to find elements when exact matches fail

2. **Versatile Action Types**
   - Click: Standard left-click on GUI elements
   - Right-click: Simulates right-click for context menu operations
   - Double-click: Fast double-left-click for opening files or specific interactions
   - Input text: Sends text input to fields, supporting multi-line input
   - Special Key presses: Simulates keyboard shortcuts and function key presses
   - Program Execution: Launches specified programs or scripts

3. **Flexible Window Targeting**
   - Specific Window: Target actions to a particular application window
   - Taskbar: Interact with taskbar items and system tray
   - Desktop: Perform actions on desktop icons and elements
   - Auto-focus: Automatically switches to the correct window before performing actions

4. **Precise Timing Control**
   - Configurable wait times between each action
   - Helps in dealing with slow-loading applications or network delays
   - Ensures actions are performed only after the UI is ready

5. **Visual Feedback and Debugging**
   - Optional visual highlighting of targeted GUI elements during execution
   - Helps in debugging by showing which elements are being interacted with
   - Can be toggled on/off for production use

6. **Robust Error Handling and Retry Mechanism**
   - Configurable retry attempts for finding GUI elements
   - Detailed error messages for troubleshooting
   - Graceful handling of missing elements or failed actions

7. **Automatic Code Generation**
   - Generates a standalone Python script for the recorded automation sequence
   - Created script can be run independently of the GUI tool
   - Includes all necessary functions and error handling

8. **Import and Export Functionality**
   - Save recorded sequences for later use
   - Import previously generated scripts for editing or execution
   - Facilitates sharing and version control of automation scripts

9. **Intelligent Desktop and Taskbar Handling**
   - Special functions to interact with desktop icons and taskbar items
   - Handles various desktop implementations (e.g., WorkerW, Progman)
   - Scrolls desktop to find items not immediately visible

10. **Multi-step Input Operations**
    - Supports complex input scenarios like filling forms
    - Can simulate natural typing with configurable intervals between keystrokes

11. **Special Key and Modifier Support**
    - Simulates complex key combinations (e.g., Ctrl+Shift+Esc)
    - Supports all standard keyboard keys and modifiers

12. **Customizable UI Tree Exploration**
    - Configurable depth for UI tree traversal
    - Helps in finding deeply nested GUI elements

13. **Detailed Logging and Reporting**
    - Comprehensive logging of each action and result
    - Helps in debugging complex automation sequences

14. **Cross-application Automation**
    - Seamlessly automate workflows spanning multiple applications
    - Handles focus switching between different windows

15. **Safety Features**
    - Option to simulate mouse movements for more natural interaction
    - Configurable delays to prevent overwhelming target applications

16. **Extensible Architecture**
    - Modular design allows for easy addition of new action types
    - Can be extended to support additional UI frameworks or special applications

## Requirements

- Python 3.x
- Required Python packages:
  - tkinter
  - ttkbootstrap
  - pyautogui
  - uiautomation
  - keyboard
  - pywin32

## Installation

1. Clone this repository or download the source code.
2. Install the required Python packages:

```
pip install tkinter ttkbootstrap pyautogui uiautomation keyboard pywin32
```

## Usage

1. Run the main script:

```
python enhanced_windows_auto_gui.py
```

2. Use the GUI to record your automation steps:
   - The tool will automatically detect GUI elements under the mouse cursor.
   - Select the desired action (Click, Input text, etc.) for each step.
   - Add wait times between steps as needed.
   - Use the "Add Step" button to record each action.

3. Review and edit your recorded steps in the listbox.

4. Generate the automation script using the "Generate Code" button.

5. The generated script will be saved to your desktop as "Win_automation.py".

6. Run the generated script to execute your automation sequence.

## Key Components

- **AutoGUIApp**: The main application class that creates the GUI and manages the automation recording process.
- **find_desktop()**: Function to locate the Windows desktop, handling various desktop configurations.
- **find_control()**: Recursive function to locate specific GUI controls within a window or the entire desktop.
- **perform_action()**: Executes the specified action on a given GUI element.
- **run_automation()**: The main loop that executes the recorded automation steps.

## Troubleshooting

- If elements are not being detected correctly, try adjusting the `MAX_RETRIES` and `RETRY_DELAY` variables in the generated script.
- For issues with specific applications, check if they have unique security settings that might interfere with UI automation.
- Use the visual highlighting feature (when allowed) to confirm that the correct elements are being targeted.

## Contributing

Contributions to improve Enhanced Windows Auto GUI are welcome. Please feel free to submit pull requests or create issues for bugs and feature requests.

## Acknowledgements

This project uses several open-source libraries, including `uiautomation` and `pyautogui`. We thank the maintainers and contributors of these projects.

## Disclaimer

This tool is for educational and productivity purposes only. Users are responsible for ensuring they have the necessary rights and permissions to automate interactions with target applications.

This project is created by Claude.ai model Sonnet 3.5
