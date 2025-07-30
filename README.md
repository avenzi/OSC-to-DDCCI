# OSC to DDC/CI

This is a simple windows app allows an OSC device to control basic monitor settings (brightness, contrast, power) through DDC/CI.

## Installation
1. Clone the repository
2. (optional) To run at startup, create a shortcut to `monitor_osc.exe` in `C:\Users\<user>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup`

## Usage
1. Run `monitor_osc.exe`
2. Wait a few seconds while it attempts to locate connected monitors
2. Double-click on the system tray icon to bring up the configuration window.

**Log File**: Local file path to output logs. \
**IP**: IP at which to listen for OSC data. \
**Port**: Port at which to listen for OSC data. \
**Save**: Save and apply the current configuration (also occurs when the config window is closed). \
**Refresh**: Attempts to re-locate all connected monitors, populating the dropdown in the section below. \
**Add Monitor**: Adds another monitor tab for a multi-monitor setup, allowing configuration of each monitor separately.

#### Monitor Configuration Section
**ID**: Choose a monitor from the dropdown, identified by model. \
**Interval**: The update interval in milliseconds. 10ms seems to be a typical lower bound. \
**Paths**: Specify the OSC path to use for each optional setting: luminescence, contrast, and toggle (power)

**Range**: Defines the min and max output values, typically 0 - 100. \
**Offset**: Defines where the min and max range values start relative to the input, typically 0.0 - 1.0.

Example: \
Input:   0 -------------------------------------- 1

Offset: (0.2 - 0.5) \
Output:  0 ----- 0 ---------- 1 ----------------- 1

Range:  (20 - 80) \
Output: 20 ---- 20 --------- 80 ----------------- 80

