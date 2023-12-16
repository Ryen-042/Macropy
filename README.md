# Overview

Macropy is a versatile automation tool that serves as a keyboard hotkey manager, mouse event listener, word expander, and more! It aims to enhance efficiency by providing a simple way to define and execute custom scripts through hotkeys. This allows users to automate repetitive tasks, offering a wide range of customization options. The project is written in Python and Cython and is currently only available for Windows due to its heavy reliance on Windows API.

![Terminal Output](https://github.com/Ryen-042/Macropy/blob/main/Images/Output.png?raw=true)

## Getting Started

1. Download the source files or clone the repository:

```bash
git clone https://github.com/Ryen-042/Macropy.git
```

2. Install the required dependencies using the `make install-reqs` command or manually with:

```bash
pip install -r requirements.txt
```

You can also install the package from the source files using the `make install` command or by running `pip install .` in the project directory.

## Usage

To use this script, run `python __main__.py` file, or use the `make run` command.

## Compiling Cython Extensions

To compile the Cython extensions:

1. Install the dev-requirements with `make install-reqs-dev`.

2. Install Macropy using `make compile` or `python setup.py build_ext --inplace`.

By default, `setup.py` builds the extension from the `.pyx` files and falls back to the `.c` files if Cython is not installed. If you prefer building using the `.c` files (e.g., to avoid Cython version issues), set `USE_CYTHON=False` in `setup.py`.

## Development History

Macropy implements keyboard and mouse event hooks. Unlike conventional hooks provided by modules like `keyboard`, `pynput`, etc., Macropy offers two key benefits. Firstly, it provides the flexibility to use any key combination for triggering hotkeys, enabling easy specification of multiple key combinations to activate the same hotkey. Secondly, Macropy allows the use of any keyboard keys, including special keys like `FN`, as long as the key is reported by the operating system.

In the early stages of Macropy's development, my reliance on third-party modules presented some limitations. The `keyboard` module, in particular, fell short when it came to detecting certain keys like `FN`. Seeking an alternative, I turned to `pyWinhook`, which did address the key detection issue. However, it came at a significant cost—system resources were heavily taxed, with CPU usage hovering around 30% even during idle periods due to the inefficient continuous polling of the keyboard.

To overcome this problem, I decided to take matters into my own hands and create my own low-level hook manager from scratch. The goal was to achieve not only comprehensive key detection but also a substantial reduction in CPU usage. The results were truly satisfying, as the custom hook manager managed to bring CPU usage to less than 1% at idle—a remarkable improvement over the resource-intensive pyWinhook.

## Cython

To further improve performance, Macropy leverages Cython, a superset of Python. Python's dynamic typing, while convenient for development, introduces runtime overhead and potential errors due to dynamic type changes. Cython addresses these concerns by introducing static type annotations, allowing Python code to be compiled into efficient C/C++ extensions. It aims to combine the ease of use and high-level features of Python with the performance of compiled languages.

By explicitly annotating variables, function arguments, and return types with static types, Cython allows the compiler to generate more efficient code. The compiler can make assumptions based on the known types, optimize memory layout, and generate direct machine code instructions without the need for dynamic type checks. As a nice bonus, startup time is also reduced as the compiler can generate machine code ahead of time.

## Project Structure

The Macropy project is organized into sub-packages within the `src` directory, each serving a distinctive purpose.
```
macropy
│   scriptConfigs.py
│   __main__.py
│
└───src
    ├───cythonExtensions
    │   │
    │   ├───commonUtils
    │   │       commonUtils.c
    │   │       commonUtils.cp312-win_amd64.pyd
    │   │       commonUtils.pxd
    │   │       commonUtils.pyi
    │   │       commonUtils.pyx
    │   │
    │   ├───eventHandlers
    │   │       callbacks.py
    │   │       eventHandlers.pyx
    │   │       ...
    │   │
    │   ├───explorerHelper
    │   │       explorerHelper.pyx
    │   │       ...
    │   │
    │   ├───hookManager
    │   │       hookManager.pyx
    │   │       ...
    │   │
    │   ├───imageUtils
    │   │       imageEditor.pyx
    │   │       imageHelper.pyx
    │   │       ...
    │   │
    │   ├───keyboardHelper
    │   │       keyboardHelper.pyx
    │   │       ...
    │   │
    │   ├───mouseHelper
    │   │       mouseHelper.pyx
    │   │       ...
    │   │
    │   ├───scriptRunner
    │   │       scriptRunner.pyx
    │   │       ...
    │   │
    │   ├───systemHelper
    │   │       systemHelper.pyx
    │   │       ...
    │   │
    │   ├───trayIconHelper
    │   │   │   trayIconHelper.pyx
    │   │   │    ...
    │   │   │
    │   │   └───icons
    │   │           6617813.ico
    │   │           7466912.ico
    │   │
    │   └───windowHelper
    │           windowHelper.pyx
    │           ...
    │
    ├───Images
    │   ├───Icons
    │   │       icon-0 - (512x512).ico
    │   │       icon-1 - (512x512).ico
    │   │       ...
    │   │
    │   └───static
    │           keyboard (0.5).png
    │           keyboard.png
    │
    └───SFX
            achievement-message-tone.wav
            ...
```

### Cython Extensions (`cythonExtensions`)

The `cythonExtensions` directory contains various sub-packages, each representing a specific aspect of Macropy's functionality. Let's take a look at the `commonUtils` sub-package as an example.

- **commonUtils.c**: The C source code file containing the Cythonized code.
- **commonUtils.cp312-win_amd64.pyd**: The compiled Cython extension for Windows 64-bit and Python 3.12.
- **commonUtils.pxd**: The Cython declaration file, defining the interface to the C code. It is used by other Cython modules to access the C code.
- **commonUtils.pyi**: A Python stub file providing type hints for static analysis tools.
- **commonUtils.pyx**: The Cython source code file, an extension of Python with static type annotations.
- **__init__.py**: An empty file signaling that the directory should be considered a Python package.

### Sub-Packages Overview

1. **commonUtils**: Contains common classes and constants used by other sub-packages.
2. **eventHandlers**: Handles callbacks for various events.
3. **explorerHelper**: Assists in managing Windows Explorer-related tasks.
4. **hookManager**: Manages low-level keyboard and mouse hooks.
5. **imageUtils**: Provides image editing capabilities.
6. **keyboardHelper**: Handles keyboard-related functions.
7. **mouseHelper**: Manages mouse-related operations.
8. **scriptRunner**: Executes scripts and manages related functionality.
9. **systemHelper**: Assists in system-related tasks.
10. **trayIconHelper**: Manages the system tray icon.
11. **windowHelper**: Handles window-related operations.

## Key Features

Macropy's versatility extends beyond keyboard input; it also implements mouse event hooks. Activating both the keyboard and mouse listeners simultaneously introduces the intriguing possibility of combo hotkeys, seamlessly integrating actions from both input devices. Macropy also supports word expansion, allowing users to define custom text expansions and hotkeys for them.

Here are some of the operations and features that Macropy supports:

### Windows Explorer Operations

- **Create a new file:** `Ctrl + Shift + 'm'`.

![Creating A New File](https://github.com/Ryen-042/Macropy/blob/main/Images/New_File.gif?raw=true)

- **Copy the full path:** `Shift + F2` for selected files and directories in the active explorer/desktop.

- **Merge selected images:** `Ctrl + Shift + 'p'` in the active explorer into a PDF file.

![Merging Images Into PDF](https://github.com/Ryen-042/Macropy/blob/main/Images/Merging_Images_To_PDF.gif?raw=true)

- **Convert selected Word files to PDFs:** `Backtick + 'o'`.

- **Convert selected Powerpoint files to PDFs:** `Backtick + 'p'`.

![Converting Powerpoint Files To PDF](https://github.com/Ryen-042/Macropy/blob/main/Images/Converting_Powerpoint_To_PDF.gif?raw=true)

- **Reopen closed file explorer:** `Ctrl + [Win | FN] + 't'`. (Currently only tracks explorers closed with `Alt + F4` or `Ctrl + W`).

![Reopening Closed Explorer](https://github.com/Ryen-042/Macropy/blob/main/Images/Reopening_Closed_Explorer.gif?raw=true)

- **Convert selected images to icon files:** `Ctrl + Alt + Win + 'I'`.
- **Convert selected '.mp3' files to '.wav' files:** `Ctrl + Alt + Win + 'M'`.

### Window Management Operations

- **Move the active window around:** `Backtick + (↑/→/↓/←)`.

![Moving Window Around](https://github.com/Ryen-042/Macropy/blob/main/Images/Moving_Window.gif?raw=true)

- **Make the active window always on top:** `FN + Ctrl + 'a'`.

- **Adjust window opacity:** `Backtick + ['+' | '='] / ['-' | '_']`.

![Changing Opacity](https://github.com/Ryen-042/Macropy/blob/main/Images/Changing_Opacity.gif?raw=true)

### OS/Script Management Operations

- **Put the device to sleep:** `Ctrl + Alt + [Win | FN] + 's'`.
  (*Requires [psshutdown](https://learn.microsoft.com/en-us/sysinternals/downloads/psshutdown) downloaded to `C:\Program Files\Utilities\PSTools`.*)

- **Shutdown the system:** `Ctrl + Alt + [Win | FN] + 'q'`.

- **Adjust system volume:** `Ctrl + Shift + <__>`
    - `['+' | '=']` / `['-' | '_']`.

    - `Wheel Scroll Up` / `Down`.


- **Adjust brightness:** `Backtick + [F2 / F3]`.

- **Suspend/resume the process:** `Backtick + [Pause / Alt] + Pause` for the active window.

- **Scroll up/down (by sending mouse wheel scrolls):** `ScrLock_ON + ['w'/'a'/'s'/'d']`.

### Text Expansion

- **Expand text:** (e.g., type `:name` in a text area).

![Text Expansion](https://github.com/Ryen-042/Macropy/blob/main/Images/Expanding_Text.gif?raw=true)

- **Open a file or folder:** (e.g., type `!paint` anywhere).

![Opening Paint](https://github.com/Ryen-042/Macropy/blob/main/Images/Opening_Paint.gif?raw=true)

### Extra

- The script does not receive keyboard events when the active process is elevated and the script is not. A notification message will be printed in the console every 10 seconds with a sound in this scenario. You can disable this feature by setting `ENABLE_ELEVATED_PRIVILEGES_CHECKER = False` in `scriptConfigs.py`.

- **Reload the script with elevated privileges:** `Ctrl + Alt + [Win | FN] + ESC`.

![Elevated Process Checker](https://github.com/Ryen-042/Macropy/blob/main/Images/Elevated_Checker.png?raw=true)

- **Terminate the script:** `FN + ESC`.

- **Display a toast notification with useful options:** `FN + '/'`.

- **Clear the terminal:**
  - `Ctrl + FN + 'c'`.
    
  - Type `>cls`.

- **Toggle terminal output:** `FN + Alt + 's'`.

- **Suppressing keyboard keys:** `Ctrl + Alt + 'd'`. (Hotkeys that don't use characters that require `Shift`, e.g., `':'` and `'!'`) still work.

- The script will not allow more than one active instance at a time.

- No need to restart the script after making changes to the configuration file (`scriptConfigs.py`) or hotkeys (`callbacks.py`). You can:
  - **Reload Hotkeys:** from System Tray Menu or with `Ctrl + Shift + Win + 'r'`.
    
  - **Reload Configs:** from System Tray Menu or from the Notification Window that can be opened with `Win + '/'`.

- Open a simple image editor (on-the-fly editing/displaying): `Backtick + '\'`

- **System Tray Notification.**
