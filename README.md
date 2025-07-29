# all-seeing-eye
Creepy app that tracks your most used and last used applications.

AllSeeingEye is a Windows application that tracks the active applications you use, records time spent on each, and provides a simple GUI for viewing and managing tracked data. It supports whitelisting apps, exporting usage data, running in the system tray, and includes configurable settings like dark mode and startup options.
Features

    Tracks time spent on active windows/applications

    Detects idle time and tracks it separately

    Maintains a whitelist of apps to track (empty whitelist tracks all)

    Displays the three most recently used applications with usage time and last used timestamp

    Export tracked data to a text file

    Dark mode toggle for the interface

    Start minimized to tray option

    System tray integration with icon and menu

    Run at startup (current user or all users, requires admin)

    Thread-safe background tracking loop

Installation

    Clone or download the repository.

    Ensure Python 3.8+ is installed on Windows.

    Install dependencies:

pip install pywin32 psutil pystray pillow

Usage

Run the main script:

python all_seeing_eye.py

The GUI will open showing tracked applications and usage.
How It Works

    The app uses Windows API via pywin32 to get the currently active window's process name.

    It periodically polls (every 1 second) and records the amount of time spent on each active app.

    Idle time (no input detected for 5 minutes) is tracked separately as "Idle Time".

    Usage data and settings are saved in JSON files (data.json, settings.json, last_used.json).

    Users can whitelist specific executables to limit tracking.

    The GUI displays recent apps, allows exporting data, and adjusting settings.

    The app supports minimizing to the system tray with a context menu.

Configuration

    Whitelist apps via GUI by adding/removing .exe files.

    Enable or disable dark mode in the GUI.

    Choose to start minimized and whether the app runs on startup (per user or all users).

    All settings are persisted in settings.json.

Dependencies

    Python 3.8+

    pywin32

    psutil

    pystray

    Pillow
