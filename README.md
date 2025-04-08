# MES Access Switch Config Generator

This Python script automates the creation of configuration files for MES access switches based on input data provided in an Excel spreadsheet. It uses a graphical user interface (GUI) to guide users through necessary inputs and produces a set of `.txt` configuration files, one per switch, zipped together for convenience.

---

## Features

- GUI prompts for user-friendly interaction.
- Reads switch configuration parameters from an Excel file.
- Automatically detects all unique zones.
- Prompts user to enter VLAN configuration for each zone:
  - VLAN ID
  - VLAN Name
  - Subnet Mask
  - Gateway IP
- Asks the user to provide descriptions for both uplink ports:
  - TenGigabitEthernet1/1/1
  - TenGigabitEthernet1/1/2
- Automatically generates a cleaned description for the Port-Channel interface by removing numeric suffixes.
- Generates full switch configuration scripts based on input.
- Saves individual config files and zips them for easy use.

---

## Excel File Format

Your Excel file must contain the following column headers:

- `hostname`
- `ip`
- `port` (either 24 or 48)
- `zone`
- `po`

Example:

| hostname | ip        | port | zone | po  |
|----------|-----------|------|------|-----|
| sw1      | 10.0.0.1  | 24   | ZoneA| 10  |
| sw2      | 10.0.0.2  | 48   | ZoneB| 20  |

---

## How to Use

1. **Run the script** using Python 3:
    ```bash
    python your_script_name.py
    ```

2. **Follow the GUI prompts**:
   - Select the Excel file.
   - Provide VLAN ID, name, subnet mask, and gateway for each unique zone found in the spreadsheet.
   - Enter uplink descriptions for TenGigabitEthernet1/1/1 and 1/1/2.

3. **Results**:
   - A folder named `YYYY-MM-DD_MES_access_switch_config` will be created.
   - Each switch gets a `.txt` file inside this folder.
   - A zipped version of the folder will also be created.

---

## Packaging as a Standalone Executable

1. Install PyInstaller:
    ```bash
    pip install pyinstaller
    ```

2. Create the `.exe` file:
    ```bash
    pyinstaller --onefile your_script_name.py
    ```

3. Find the executable inside the `dist/` folder.

---

## Requirements

- Python 3.x
- Libraries:
  - `pandas`
  - `openpyxl`
  - `tkinter` (standard with Python)

To install required libraries:
```bash
pip install pandas openpyxl
```

---

## Output

- `.txt` configuration file for each switch.
- All config files zipped into one archive: `YYYY-MM-DD_MES_access_switch_config.zip`

