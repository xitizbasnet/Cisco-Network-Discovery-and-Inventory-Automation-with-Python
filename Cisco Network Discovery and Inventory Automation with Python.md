# Cisco Network Discovery and Inventory Automation with Python

## 📘 Overview

This Python script automates network discovery and inventory collection for Cisco CBS350 / CBS Series switches using SSH.

The script connects to multiple Cisco switches, collects system and hardware information, retrieves interface information, discovers CDP and LLDP neighbors, and generates a consolidated Microsoft Excel report.

The generated Excel workbook contains separate worksheets for:

* Summary
* Switch Inventory
* Port Inventory
* CDP Neighbors
* LLDP Neighbors
* Failed Devices

The script supports parallel connections, allowing multiple switches to be scanned simultaneously.

The Excel report is automatically saved to the Desktop of the Windows user account running the program and, when supported, is opened automatically after the scan completes.

---

## 🏗️ Script Capabilities

The script provides the following capabilities:

| Capability               | Description                                                                                  |
| ------------------------ | -------------------------------------------------------------------------------------------- |
| Cisco Discovery          | Connects to multiple Cisco switches using SSH                                                |
| Parallel Processing      | Scans multiple switches simultaneously                                                       |
| System Information       | Collects hostname, uptime, contact, location, MAC address, and system information            |
| Hardware Inventory       | Collects model, serial number, product ID, and hardware version                              |
| Firmware Information     | Collects the Cisco firmware version                                                          |
| Port Inventory           | Collects interface status, speed, duplex, negotiation, flow control, back pressure, and MDIX |
| Port Descriptions        | Collects interface descriptions                                                              |
| CDP Discovery            | Collects Cisco Discovery Protocol neighbor information                                       |
| LLDP Discovery           | Collects Link Layer Discovery Protocol neighbor information                                  |
| Pagination Handling      | Handles CBS350 command pagination                                                            |
| Data Cleaning            | Removes ANSI escape sequences, invalid Excel characters, and pagination strings              |
| Excel Reporting          | Creates a structured Excel workbook                                                          |
| Failed Device Reporting  | Records switches that could not be successfully scanned                                      |
| Desktop Detection        | Detects the Windows Desktop automatically                                                    |
| OneDrive Desktop Support | Supports redirected/OneDrive Desktop locations                                               |
| Automatic Excel Opening  | Attempts to open the generated Excel workbook automatically                                  |

---

## 📋 Prerequisites

Before running the script, ensure the following requirements are available.

### Operating System

The script is designed primarily for Windows because it:

* Detects the Windows Desktop location.
* Supports redirected/OneDrive Desktop locations.
* Uses `os.startfile()` to open the generated Excel file automatically.

### Python

Install Python on the Windows computer running the script.

Verify the installation:

```powershell
python --version
```

If the `python` command is not available, verify whether the Python launcher is installed:

```powershell
py --version
```

---

## 📦 Required Python Packages

The script imports the following external Python packages:

```text
paramiko
pandas
openpyxl
```

The following modules are included with Python and do not require separate installation:

```text
time
getpass
re
os
pathlib
datetime
concurrent.futures
```

Install the required packages with:

```powershell
pip install paramiko pandas openpyxl
```

If `pip` is not recognized, use:

```powershell
python -m pip install paramiko pandas openpyxl
```

Alternatively:

```powershell
py -m pip install paramiko pandas openpyxl
```

---

## 🔐 Access Requirements

The computer running the script must be able to reach the Cisco switches over the configured SSH port.

The default SSH port used by the script is:

```python
SSH_PORT = 22
```

The Cisco switches must allow SSH access and the supplied Cisco credentials must have sufficient permissions to execute the commands used by the script.

The script prompts for credentials when it starts:

```text
Enter Cisco Username:
Enter Cisco Password:
```

The password is collected using Python's `getpass` module and is not entered directly into the source code.

---

## ⚙️ Configuration

The primary configuration is located near the beginning of the script.

### Switch IP Addresses

The switches to be scanned are defined in:

```python
SWITCH_IPS = [
    "192.168.1..2",
    "192.168.1..3",
    "192.168.1..4",
    "192.168.1..5",
    "192.168.1..6",
    "192.168.1..7",
    "192.168.1..8",
    "192.168.1..9",
    "192.168.1..10",
    "192.168.1..11",
    "192.168.1..12",
    "192.168.1..13",
    "192.168.1..14",
    "192.168.1..15",
    "192.168.1..16",
    "192.168.1..17",
    "192.168.1..19",
    "192.168.1..20",
    "192.168.1..21",
    "192.168.1.254",
]
```

### ⚠️ Important

The IP addresses in the supplied source contain entries such as:

```text
192.168.1..2
192.168.1..3
192.168.1..4
```

These contain two consecutive periods and do not represent standard IPv4 addresses.

For example:

```text
192.168.1..2
```

would normally need to be represented as:

```text
192.168.1.2
```

The supplied source code is preserved unchanged in the **Complete Python Script** section below.

> **Note:** Correct any invalid IP addresses in the `SWITCH_IPS` list before executing the script against the production network.

---

## 🔄 Adding Additional Switches

Additional switches can be added to the `SWITCH_IPS` list.

The source code includes the following example:

```python
# Add your other switches here.
#
# Example:
#
# "192.168.1..4",
# "192.168.1..5",
# "192.168.1..6",
# "192.168.1..7",
```

For valid IPv4 addresses, replace the example values with the actual management IP addresses of the Cisco switches.

---

## ⚡ Parallel Connections

The maximum number of switches scanned simultaneously is configured using:

```python
MAX_WORKERS = 10
```

The script uses Python's `ThreadPoolExecutor` to perform parallel discovery.

The configured value means that up to 10 switch connections can be processed simultaneously.

The relevant implementation is:

```python
with ThreadPoolExecutor(
    max_workers=MAX_WORKERS
) as executor:
```

This can significantly reduce the total discovery time when multiple switches need to be scanned.

---

## 🔌 SSH Port

The SSH port is configured as:

```python
SSH_PORT = 22
```

The script establishes the Paramiko transport using:

```python
transport = paramiko.Transport(
    (
        ip,
        SSH_PORT
    )
)
```

If the Cisco environment uses a different SSH port, the configuration value must be changed accordingly.

---

# 🚀 Installation

## Step 1: Install Python

Install Python on the Windows computer.

Verify:

```powershell
python --version
```

---

## Step 2: Install Python Dependencies

Run:

```powershell
pip install paramiko pandas openpyxl
```

Or:

```powershell
python -m pip install paramiko pandas openpyxl
```

---

## Step 3: Save the Python Script

Save the provided script with a `.py` extension.

For example:

```text
cisco_network_inventory.py
```

---

## Step 4: Configure Switch IP Addresses

Open the Python file and locate:

```python
SWITCH_IPS = [
```

Add the management IP addresses of the Cisco switches that should be scanned.

---

## Step 5: Run the Script

From PowerShell or Command Prompt:

```powershell
python cisco_network_inventory.py
```

---

# 🔑 Credential Prompt

When the script starts, it requests the Cisco username:

```text
Enter Cisco Username:
```

The script then requests the password:

```text
Enter Cisco Password:
```

The password is collected using:

```python
password = getpass.getpass(
    "Enter Cisco Password: "
)
```

This prevents the password from being directly displayed while it is entered.

---

# 🔎 Discovery Process

For each configured switch, the script performs the following operations.

1. Establishes an SSH transport.
2. Starts the SSH client.
3. Performs the authentication sequence.
4. Opens an interactive SSH session.
5. Detects the Cisco switch prompt.
6. Executes `show system`.
7. Executes `show version`.
8. Executes `show inventory`.
9. Executes `show interfaces description`.
10. Executes `show interfaces status`.
11. Executes `show cdp neighbors detail`.
12. Executes `show lldp neighbors detail`.
13. Parses the command output.
14. Cleans the collected data.
15. Stores the results.
16. Generates the Excel workbook.

---

# 🖥️ Cisco Commands Used

The script executes the following Cisco commands.

## `show system`

Used to collect system information including:

* System Description
* System Up Time
* System Contact
* System Name
* System Location
* System MAC Address
* System Object ID
* System Model

---

## `show version`

Used to retrieve the firmware version.

The parser searches for:

```text
Version:
```

and extracts the associated version value.

---

## `show inventory`

Used to retrieve hardware inventory information.

The script extracts:

| Field            | Description              |
| ---------------- | ------------------------ |
| Model            | Cisco switch model       |
| Serial Number    | Device serial number     |
| Product ID       | Cisco product identifier |
| Hardware Version | Hardware/VID information |
| Description      | Device description       |

The parser expects the CBS350 inventory format containing information such as:

```text
NAME: "1"
DESCR: "CBS350..."
PID: CBS350-48FP-4X
VID: V05
SN: PSZ27271RG2
```

---

## `show interfaces description`

Used to retrieve interface descriptions.

The parser associates interface names with their descriptions.

For example:

```text
gi1/0/1   description
```

The interface and description are later combined with the interface status information.

---

## `show interfaces status`

Used to collect interface status information.

The script processes the following fields:

| Field         | Description          |
| ------------- | -------------------- |
| Port          | Interface identifier |
| Type          | Interface type       |
| Duplex        | Duplex setting       |
| Speed         | Interface speed      |
| Negotiation   | Negotiation status   |
| Flow Control  | Flow control status  |
| Link State    | Interface state      |
| Back Pressure | Back pressure status |
| MDIX          | MDIX status          |

The source code documents the CBS350 format as:

```text
Port     Type         Duplex  Speed Neg      ctrl State       Pressure Mdix
-------- ------------ ------  ----- -------- ---- ----------- -------- -------
gi1/0/1  1G-Copper    Full    100   Enabled  Off  Up          Disabled Off
```

---

## `show cdp neighbors detail`

Used to collect Cisco Discovery Protocol neighbor information.

The script attempts to collect:

* Local Hostname
* Local IP
* Neighbor Hostname
* Device ID
* Neighbor IP
* Platform
* Capabilities
* Local Interface
* Neighbor Interface
* Holdtime
* Version
* Duplex
* Native VLAN

---

## `show lldp neighbors detail`

Used to collect Link Layer Discovery Protocol neighbor information.

The script attempts to collect:

* Local Hostname
* Local IP
* Neighbor
* Neighbor IP
* Platform
* Local Interface
* Neighbor Interface
* System Description

---

# 📊 Excel Report

The generated Excel file uses the following naming convention:

```text
Cisco_Network_Inventory_YYYY-MM-DD_HH-MM-SS.xlsx
```

The filename is generated using:

```python
timestamp = datetime.now().strftime(
    "%Y-%m-%d_%H-%M-%S"
)
```

The final file path is:

```python
EXCEL_FILE = DESKTOP / (
    f"Cisco_Network_Inventory_{timestamp}.xlsx"
)
```

---

# 📁 Desktop Detection

The script dynamically determines the Desktop location of the Windows account running the program.

It first attempts to use the Windows Shell API:

```python
CSIDL_DESKTOPDIRECTORY = 0x0010
SHGFP_TYPE_CURRENT = 0
```

It then falls back to:

```text
Desktop
```

or:

```text
OneDrive\Desktop
```

if required.

This allows the script to support standard and redirected/OneDrive Desktop locations.

---

# 📑 Excel Worksheets

The generated workbook contains six worksheets.

## Summary

The `Summary` worksheet contains overall discovery statistics.

The following metrics are included:

| Metric                 |
| ---------------------- |
| Report Generated       |
| Total Switches         |
| Successful Switches    |
| Failed Switches        |
| Total Port Records     |
| Ports With Description |
| Connected Ports        |
| CDP Neighbors          |
| LLDP Neighbors         |
| Excel File             |

---

## Switch Inventory

The `Switch Inventory` worksheet contains device-level information.

The script collects:

| Field              |
| ------------------ |
| Hostname           |
| IP Address         |
| Model              |
| Serial Number      |
| Product ID         |
| Hardware Version   |
| Firmware Version   |
| System Description |
| Uptime             |
| System Contact     |
| System Location    |
| MAC Address        |
| System Object ID   |

---

## Port Inventory

The `Port Inventory` worksheet contains interface-level information.

The fields include:

| Field         |
| ------------- |
| Hostname      |
| IP Address    |
| Interface     |
| Description   |
| Type          |
| Status        |
| Speed         |
| Duplex        |
| Negotiation   |
| Flow Control  |
| Back Pressure |
| MDIX          |

---

## CDP Neighbors

The `CDP Neighbors` worksheet contains discovered Cisco Discovery Protocol neighbors.

The fields include:

| Field              |
| ------------------ |
| Local Hostname     |
| Local IP           |
| Neighbor Hostname  |
| Device ID          |
| Neighbor IP        |
| Platform           |
| Capabilities       |
| Local Interface    |
| Neighbor Interface |
| Holdtime           |
| Version            |
| Duplex             |
| Native VLAN        |

---

## LLDP Neighbors

The `LLDP Neighbors` worksheet contains discovered Link Layer Discovery Protocol neighbors.

The fields include:

| Field              |
| ------------------ |
| Local Hostname     |
| Local IP           |
| Neighbor           |
| Neighbor IP        |
| Platform           |
| Local Interface    |
| Neighbor Interface |
| System Description |

---

## Failed Devices

The `Failed Devices` worksheet records switches that could not be successfully processed.

The worksheet contains:

| Field      |
| ---------- |
| IP Address |
| Error      |

This allows administrators to identify devices that require additional investigation after the discovery process.

---

# 🧹 Data Cleaning

The script includes a `clean_text()` function to clean Cisco command output before it is written to Excel.

The function removes:

* ANSI escape sequences
* Invalid Excel control characters
* `--More--`
* `-- More --`

The script uses:

```python
re.sub(
    r"\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])",
    "",
    value
)
```

to remove ANSI escape sequences.

It also removes invalid Excel control characters using:

```python
re.sub(
    r"[\x00-\x08\x0B\x0C\x0E-\x1F]",
    "",
    value
)
```

---

# 📄 Pagination Handling

Cisco CBS350 command output can contain pagination prompts.

The script detects several pagination patterns, including:

```text
more:
```

```text
--More--
```

```text
-- More --
```

```text
<space>, quit:
```

```text
One line: <return>
```

When pagination is detected, the script sends a space character:

```python
channel.send(
    " "
)
```

This allows the command output to continue.

---

# 🔌 Interface Description Preservation

The script specifically prevents interface descriptions from being lost when an interface appears in:

```text
show interfaces description
```

but does not appear in:

```text
show interfaces status
```

It creates additional port records for description-only interfaces.

This behavior is implemented using:

```python
existing = {

    row[
        "Interface"
    ].lower()

    for row in ports
}
```

and then checking:

```python
if interface not in existing:
```

This ensures interface descriptions remain available in the final `Port Inventory` worksheet.

---

# 🔄 Parallel Discovery

The script uses:

```python
ThreadPoolExecutor
```

from:

```python
from concurrent.futures import ThreadPoolExecutor, as_completed
```

Each switch is submitted as a separate task:

```python
future = executor.submit(
    collect_switch,
    ip,
    username,
    password
)
```

The results are then collected as individual tasks complete:

```python
for future in as_completed(
    futures
):
```

This allows the script to continue processing other switches while individual network connections are completing.

---

# 📈 Discovery Statistics

The script calculates several statistics before creating the Excel report.

These include:

```text
Total switches
Successful switches
Failed switches
Total port records
Ports with descriptions
Connected ports
CDP neighbors
LLDP neighbors
```

The connected-port calculation checks for:

```python
row.get(
    "Status",
    ""
).lower()
== "connected"
```

---

# 🎨 Excel Formatting

The script formats every worksheet using the `format_excel()` function.

The following formatting is applied:

### Freeze Header Row

```python
worksheet.freeze_panes = "A2"
```

This keeps the first row visible while scrolling.

### Autofilter

The script applies an autofilter when the worksheet contains more than one row:

```python
worksheet.auto_filter.ref = (
    worksheet.dimensions
)
```

### Automatic Column Width

The script determines the maximum content length in each column and adjusts the width.

The maximum column width is limited to:

```text
60
```

---

# 🖥️ Console Output

During execution, the script displays progress information.

Examples include:

```text
[START] Connecting to 192.168.1.x
```

```text
[SSH OK] 192.168.1.x
```

```text
[LOGIN SUCCESS] 192.168.1.x
```

```text
[COMMAND] show system
```

```text
[COMMAND] show version
```

```text
[COMMAND] show inventory
```

```text
[COMMAND] show interfaces description
```

```text
[COMMAND] show interfaces status
```

```text
[COMMAND] show cdp neighbors detail
```

```text
[COMMAND] show lldp neighbors detail
```

Successful devices display:

```text
[SUCCESS] hostname (ip)
```

Failed devices display:

```text
[FAILED] ip
```

followed by the associated error.

---

# ❗ Error Handling

Each switch is processed inside an exception-handling structure.

If a switch fails, the script records the error instead of stopping the entire discovery process.

The result contains:

```python
"success": False
```

and:

```python
"error": ""
```

The failed switch is subsequently added to the `Failed Devices` worksheet.

The script also includes a top-level exception handler around `main()`:

```python
if __name__ == "__main__":

    try:
        main()
    except Exception as error:
        print()
        print("=" * 70)
        print("PROGRAM ERROR")
        print("=" * 70)
        print()
        print(error)
        print()
        input("Press ENTER to close this window...")
```

---

# 🔒 Security Considerations

## Credential Handling

The Cisco password is requested interactively using:

```python
getpass.getpass()
```

Do not hard-code production credentials into the script.

## Source Code Protection

If the script is stored in a GitHub repository:

* Do not commit usernames and passwords.
* Do not commit production credentials.
* Do not commit private SSH keys.
* Do not commit sensitive network information unless the repository is appropriately protected.

## Network Access

Ensure that SSH access to the Cisco switches is restricted according to organizational security requirements.

## Least Privilege

Use a Cisco account with only the permissions required for the commands executed by the script.

---

# ⚠️ Important Operational Notes

> **Important:** The source code supplied for this documentation contains IP addresses such as `192.168.1..2`. These are preserved exactly in the complete source code below. Correct them to valid IPv4 addresses before running the script.

> **Note:** The script is specifically documented for Cisco CBS350 / CBS Series command output. Output formats can vary between Cisco platforms, firmware versions, and command implementations.

> **Note:** CDP and LLDP results depend on the corresponding protocols being available and enabled on the Cisco switches and neighboring devices.

> **Note:** The script expects the Cisco switch to present the login prompts and command output formats handled by its parsing logic.

---

# 🔧 Troubleshooting

## Python Is Not Recognized

If:

```powershell
python --version
```

does not work, try:

```powershell
py --version
```

If neither command works, install Python and ensure it is available in the system PATH.

---

## Required Module Is Missing

If Python reports an error such as:

```text
ModuleNotFoundError
```

install the required dependencies:

```powershell
pip install paramiko pandas openpyxl
```

---

## SSH Connection Fails

Verify:

* The switch IP address is correct.
* The switch is reachable.
* SSH is enabled.
* TCP port 22 is accessible.
* The supplied username is correct.
* The supplied password is correct.
* Network access controls allow the connection.

---

## Username Prompt Not Found

The script expects one of the following prompts:

```text
User Name:
```

```text
Username:
```

```text
username:
```

If the switch presents a different login sequence, the authentication logic may need to be adapted.

---

## Password Prompt Not Found

The script expects:

```text
Password:
```

If the Cisco device uses a different authentication sequence, the login handling may need to be adjusted.

---

## Cisco Prompt Not Detected

The script searches for a prompt matching:

```text
[A-Za-z0-9_.:/()\-]+[>#]
```

If the switch prompt does not match the expected pattern, the script can report:

```text
Cisco switch prompt not detected
```

---

## CDP Neighbors Worksheet Is Empty

An empty `CDP Neighbors` worksheet can occur when the switch does not return CDP neighbor information.

Check:

```text
show cdp neighbors detail
```

directly on the switch.

Also verify that CDP is available and enabled where required.

---

## LLDP Neighbors Worksheet Is Empty

An empty `LLDP Neighbors` worksheet can occur when no LLDP neighbor information is returned.

Check:

```text
show lldp neighbors detail
```

directly on the switch.

Also verify LLDP configuration and whether neighboring devices advertise LLDP information.

---

## Port Descriptions Are Missing

The script collects descriptions separately using:

```text
show interfaces description
```

and combines them with:

```text
show interfaces status
```

If descriptions are missing, manually verify:

```text
show interfaces description
```

on the affected switch.

---

## Excel Report Is Not Opened Automatically

The script attempts to open the report using:

```python
os.startfile(str(EXCEL_FILE))
```

If this fails, the report is still saved to the configured Desktop location.

The script displays:

```text
[WARNING] Could not open Excel automatically.
```

and displays the associated error.

---

# ✅ Validation

After the scan completes, verify the following:

* [ ] The number of configured switches is correct.
* [ ] Successful switches appear in `Switch Inventory`.
* [ ] Failed switches appear in `Failed Devices`.
* [ ] Hostnames are populated.
* [ ] IP addresses are correct.
* [ ] Model information is populated.
* [ ] Serial numbers are populated.
* [ ] Product IDs are populated.
* [ ] Hardware versions are populated.
* [ ] Firmware versions are populated.
* [ ] Port descriptions are populated where configured.
* [ ] Port status information is populated.
* [ ] CDP neighbors are populated where available.
* [ ] LLDP neighbors are populated where available.
* [ ] The Summary worksheet contains the expected statistics.
* [ ] The generated Excel workbook opens successfully.

---

# 📊 Expected Excel Workbook

After a successful execution, the workbook should contain:

```text
Cisco_Network_Inventory_YYYY-MM-DD_HH-MM-SS.xlsx
```

with the following worksheets:

```text
Summary
Switch Inventory
Port Inventory
CDP Neighbors
LLDP Neighbors
Failed Devices
```

---

# 💡 Best Practices

## Keep the Switch List Organized

Maintain the `SWITCH_IPS` list carefully and verify every IP address before running a production scan.

## Start With a Small Test

Before scanning a large environment, test the script against one or a small number of switches.

## Review Failed Devices

Always review the `Failed Devices` worksheet after every scan.

## Validate Cisco Output

If information is missing from the report, run the corresponding Cisco command manually and compare the actual output with the parser expectations.

## Protect Credentials

Never store Cisco passwords directly in the Python source code.

## Review Generated Reports

The Excel workbook should be reviewed after every production discovery operation to verify that the collected data is complete.

---

# 📎 Python Modules Used

The script imports:

```python
import paramiko
import pandas as pd
import time
import getpass
import re
import os
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
```

Their primary purposes are:

| Module               | Purpose                                            |
| -------------------- | -------------------------------------------------- |
| `paramiko`           | SSH communication with Cisco switches              |
| `pandas`             | Data processing and Excel data preparation         |
| `time`               | Timing, delays, and command-output handling        |
| `getpass`            | Secure interactive password input                  |
| `re`                 | Regular expressions and output parsing             |
| `os`                 | Operating system functionality and Excel launching |
| `pathlib`            | File and Desktop path handling                     |
| `datetime`           | Timestamp generation                               |
| `concurrent.futures` | Parallel switch discovery                          |

---

# 🐍 Complete Python Script

The following is the complete Python source provided for this documentation. The code is preserved as supplied.

```python
import paramiko
import pandas as pd
import time
import getpass
import re
import os
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed


# ============================================================
# CISCO NETWORK DISCOVERY & INVENTORY
# Cisco CBS350 / CBS Series
# ============================================================


# ============================================================
# CONFIGURATION
# ============================================================

SWITCH_IPS = [
    "192.168.1..2",
    "192.168.1..3",
    "192.168.1..4",
    "192.168.1..5",
    "192.168.1..6",
    "192.168.1..7",
    "192.168.1..8",
    "192.168.1..9",
    "192.168.1..10",
    "192.168.1..11",
    "192.168.1..12",
    "192.168.1..13",
    "192.168.1..14",
    "192.168.1..15",
    "192.168.1..16",
    "192.168.1..17",
    "192.168.1..19",
    "192.168.1..20",
    "192.168.1..21",
    "192.168.1.254",


    # Add your other switches here.
    #
    # Example:
    #
    # "192.168.1..4",
    # "192.168.1..5",
    # "192.168.1..6",
    # "192.168.1..7",
]


# Maximum number of switches scanned simultaneously
MAX_WORKERS = 10


# SSH port
SSH_PORT = 22


# ============================================================
# DESKTOP LOCATION
# ============================================================

def get_desktop():

    # Resolve the Desktop of the Windows account currently running
    # this program. No username is hard-coded.
    # Handles normal Desktop and redirected/OneDrive Desktop.

    try:
        import ctypes
        from ctypes import wintypes

        CSIDL_DESKTOPDIRECTORY = 0x0010
        SHGFP_TYPE_CURRENT = 0

        buffer = ctypes.create_unicode_buffer(wintypes.MAX_PATH)

        result = ctypes.windll.shell32.SHGetFolderPathW(
            None,
            CSIDL_DESKTOPDIRECTORY,
            None,
            SHGFP_TYPE_CURRENT,
            buffer
        )

        if result == 0 and buffer.value:
            desktop = Path(buffer.value)
            desktop.mkdir(parents=True, exist_ok=True)
            return desktop

    except Exception:
        pass

    home = Path.home()

    desktop = home / "Desktop"
    if desktop.exists():
        return desktop

    onedrive_desktop = home / "OneDrive" / "Desktop"
    if onedrive_desktop.exists():
        return onedrive_desktop

    desktop.mkdir(parents=True, exist_ok=True)
    return desktop


DESKTOP = get_desktop()


# ============================================================
# EXCEL FILE NAME
# ============================================================

timestamp = datetime.now().strftime(
    "%Y-%m-%d_%H-%M-%S"
)

EXCEL_FILE = DESKTOP / (
    f"Cisco_Network_Inventory_{timestamp}.xlsx"
)


# ============================================================
# CLEAN TEXT
# ============================================================

def clean_text(value):

    if value is None:
        return ""

    value = str(value)

    # Remove ANSI escape sequences
    value = re.sub(
        r"\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])",
        "",
        value
    )

    # Remove invalid Excel control characters
    value = re.sub(
        r"[\x00-\x08\x0B\x0C\x0E-\x1F]",
        "",
        value
    )

    # Remove pagination strings
    value = value.replace(
        "--More--",
        ""
    )

    value = value.replace(
        "-- More --",
        ""
    )

    return value.strip()


# ============================================================
# CLEAN DATAFRAME
# ============================================================

def clean_dataframe(df):

    for column in df.columns:

        df[column] = df[column].apply(
            clean_text
        )

    return df


# ============================================================
# READ FROM SSH CHANNEL
# ============================================================

def read_channel(
    channel,
    timeout=5
):

    output = ""

    start_time = time.time()

    last_data_time = time.time()

    while (
        time.time() - start_time
        < timeout
    ):

        received = False

        while channel.recv_ready():

            data = channel.recv(
                65535
            )

            if not data:
                break

            text = data.decode(
                "utf-8",
                errors="ignore"
            )

            output += text

            received = True

            last_data_time = time.time()

        if received:
            start_time = time.time()

        else:

            if (
                time.time()
                - last_data_time
                > 1
            ):
                break

            time.sleep(
                0.1
            )

    return output


# ============================================================
# WAIT FOR LOGIN OUTPUT
# ============================================================

def read_login_output(
    channel,
    timeout=5
):

    return read_channel(
        channel,
        timeout
    )


# ============================================================
# SEND COMMAND
# ============================================================

def send_command(
    channel,
    command,
    wait_time=1.5
):

    channel.send(
        command + "\n"
    )

    time.sleep(
        wait_time
    )

    output = ""

    last_data_time = time.time()

    while True:

        received = False

        while channel.recv_ready():

            data = channel.recv(
                65535
            )

            if not data:
                break

            text = data.decode(
                "utf-8",
                errors="ignore"
            )

            output += text

            received = True

            last_data_time = time.time()

        lower_output = output.lower()

        # ====================================================
        # CBS350 PAGINATION
        #
        # more:
        # <space>, Quit: q or CTRL+Z,
        # One line: <return>
        # ====================================================

        pagination = (
            "more:" in lower_output
            or
            "--more--" in lower_output
            or
            "-- more --" in lower_output
            or
            "<space>, quit:" in lower_output
            or
            "one line: <return>" in lower_output
        )

        if pagination:

            # Remove visible pagination prompt
            output = re.sub(
                r"more:\s*<space>,\s*quit:.*",
                "",
                output,
                flags=re.IGNORECASE
            )

            output = re.sub(
                r"--more--",
                "",
                output,
                flags=re.IGNORECASE
            )

            output = re.sub(
                r"--\s*more\s*--",
                "",
                output,
                flags=re.IGNORECASE
            )

            # Press SPACE to continue
            channel.send(
                " "
            )

            time.sleep(
                0.7
            )

            continue

        # ====================================================
        # PROMPT DETECTION
        # ====================================================

        prompt_found = re.search(
            r"[A-Za-z0-9_.:/()\-]+[>#]\s*$",
            output,
            re.MULTILINE
        )

        if prompt_found:

            time.sleep(
                0.3
            )

            if channel.recv_ready():
                continue

            break

        # ====================================================
        # TIMEOUT
        # ====================================================

        if (
            not received
            and
            time.time()
            - last_data_time
            > 3
        ):

            break

        time.sleep(
            0.1
        )

    return output


# ============================================================
# PARSE SHOW SYSTEM
# ============================================================

def parse_system(output):

    data = {

        "System Description": "",
        "System Up Time": "",
        "System Contact": "",
        "System Name": "",
        "System Location": "",
        "System MAC Address": "",
        "System Object ID": "",
        "System Model": "",
    }

    for raw_line in output.splitlines():

        line = clean_text(
            raw_line
        )

        if ":" not in line:
            continue

        key, value = line.split(
            ":",
            1
        )

        key = key.strip()
        value = value.strip()

        if key in data:

            data[key] = value

    # ========================================================
    # CBS350 UNIT MODEL
    # ========================================================

    match = re.search(
        r"^\s*1\s+(\S+)",
        output,
        re.MULTILINE
    )

    if match:

        data["System Model"] = (
            match.group(1)
        )

    return data


# ============================================================
# PARSE SHOW VERSION
# ============================================================

def parse_version(output):

    version = ""

    match = re.search(
        r"\bVersion:\s*([^\s]+)",
        output,
        re.IGNORECASE
    )

    if match:

        version = match.group(
            1
        )

    return version


# ============================================================
# PARSE SHOW INVENTORY
# ============================================================

def parse_inventory(output):

    data = {

        "Model": "",
        "Serial Number": "",
        "Product ID": "",
        "Hardware Version": "",
        "Description": "",
    }

    # ========================================================
    # CBS350 MAIN UNIT
    #
    # NAME: "1"
    # DESCR: "CBS350..."
    # PID: CBS350-48FP-4X
    # VID: V05
    # SN: PSZ27271RG2
    # ========================================================

    match = re.search(
        r'NAME:\s*"1"\s+'
        r'DESCR:\s*"([^"]+)"\s+'
        r'PID:\s*(\S+)\s+'
        r'VID:\s*(\S+)\s+'
        r'SN:\s*(\S+)',
        output,
        re.IGNORECASE
    )

    if match:

        data["Description"] = (
            match.group(1)
        )

        data["Model"] = (
            match.group(2)
        )

        data["Hardware Version"] = (
            match.group(3)
        )

        data["Serial Number"] = (
            match.group(4)
        )

        data["Product ID"] = (
            match.group(2)
        )

    return data


# ============================================================
# PARSE INTERFACE DESCRIPTIONS
# ============================================================

def parse_interface_descriptions(output):

    descriptions = {}

    for raw_line in output.splitlines():

        line = clean_text(
            raw_line
        )

        if not line:
            continue

        lower = line.lower()

        # ====================================================
        # Ignore pagination
        # ====================================================

        if (
            lower.startswith("more:")
            or
            "<space>" in lower
            or
            "quit:" in lower
            or
            "one line:" in lower
        ):
            continue

        # ====================================================
        # Ignore command itself
        # ====================================================

        if (
            "show interfaces description"
            in lower
        ):
            continue

        # ====================================================
        # Header
        # ====================================================

        if lower.startswith(
            "port"
        ):
            continue

        if line.startswith("-"):
            continue

        # ====================================================
        # Interface + description
        #
        # gi1/0/1   description
        # ====================================================

        match = re.match(
            r"^((?:Gi|Fa|Te|Tw|Eth|Fo|Hu|Po)[A-Za-z0-9/.-]+)\s+(.*)$",
            line,
            re.IGNORECASE
        )

        if not match:
            continue

        interface = match.group(
            1
        ).strip()

        description = match.group(
            2
        ).strip()

        if interface.lower() == "more":
            continue

        descriptions[
            interface.lower()
        ] = description

    return descriptions


# ============================================================
# PARSE INTERFACE STATUS
# ============================================================

def parse_interface_status(
    output,
    hostname,
    ip,
    descriptions
):

    ports = []

    valid_prefixes = (
        "gi",
        "fa",
        "te",
        "tw",
        "eth",
        "fo",
        "hu",
        "po",
    )

    for raw_line in output.splitlines():

        line = clean_text(raw_line)

        if not line:
            continue

        lower = line.lower()

        # Ignore pagination and command/header text.
        if (
            lower.startswith("more:")
            or "<space>" in lower
            or "quit:" in lower
            or "one line:" in lower
        ):
            continue

        if "show interfaces status" in lower:
            continue

        if lower.startswith("port"):
            continue

        if line.startswith("-"):
            continue

        parts = line.split()

        if not parts:
            continue

        interface = parts[0]

        if not interface.lower().startswith(valid_prefixes):
            continue

        # Cisco CBS350 "show interfaces status" format:
        #
        # Port     Type         Duplex  Speed Neg      ctrl State       Pressure Mdix
        # -------- ------------ ------  ----- -------- ---- ----------- -------- -------
        # gi1/0/1  1G-Copper    Full    100   Enabled  Off  Up          Disabled Off
        #
        # Positional fields:
        # 0 = Port
        # 1 = Type
        # 2 = Duplex
        # 3 = Speed
        # 4 = Neg
        # 5 = Flow Control
        # 6 = Link State
        # 7 = Back Pressure
        # 8 = MDIX

        port_type = parts[1] if len(parts) > 1 else ""
        duplex = parts[2] if len(parts) > 2 else ""
        speed = parts[3] if len(parts) > 3 else ""
        negotiation = parts[4] if len(parts) > 4 else ""
        flow_control = parts[5] if len(parts) > 5 else ""
        link_state = parts[6] if len(parts) > 6 else ""
        back_pressure = parts[7] if len(parts) > 7 else ""
        mdix = parts[8] if len(parts) > 8 else ""

        description = descriptions.get(
            interface.lower(),
            ""
        )

        ports.append({
            "Hostname": hostname,
            "IP Address": ip,
            "Interface": interface,
            "Description": description,
            "Type": port_type,
            "Status": link_state,
            "Speed": speed,
            "Duplex": duplex,
            "Negotiation": negotiation,
            "Flow Control": flow_control,
            "Back Pressure": back_pressure,
            "MDIX": mdix,
        })

    return ports


# ============================================================
# PARSE CBS350 CDP
# ============================================================

def parse_cdp(
    output,
    hostname,
    ip
):

    neighbors = []

    current = {

        "Device ID": "",
        "Neighbor IP": "",
        "Platform": "",
        "Capabilities": "",
        "Local Interface": "",
        "Neighbor Interface": "",
        "Holdtime": "",
        "Version": "",
        "Duplex": "",
        "Native VLAN": "",
        "Neighbor Hostname": "",
    }

    def save_neighbor():

        if not (
            current["Device ID"]
            or
            current["Neighbor Hostname"]
        ):
            return

        neighbors.append({

            "Local Hostname":
                hostname,

            "Local IP":
                ip,

            "Neighbor Hostname":
                current[
                    "Neighbor Hostname"
                ],

            "Device ID":
                current[
                    "Device ID"
                ],

            "Neighbor IP":
                current[
                    "Neighbor IP"
                ],

            "Platform":
                current[
                    "Platform"
                ],

            "Capabilities":
                current[
                    "Capabilities"
                ],

            "Local Interface":
                current[
                    "Local Interface"
                ],

            "Neighbor Interface":
                current[
                    "Neighbor Interface"
                ],

            "Holdtime":
                current[
                    "Holdtime"
                ],

            "Version":
                current[
                    "Version"
                ],

            "Duplex":
                current[
                    "Duplex"
                ],

            "Native VLAN":
                current[
                    "Native VLAN"
                ],
        })

    for raw_line in output.splitlines():

        line = clean_text(
            raw_line
        )

        if not line:
            continue

        # ====================================================
        # New neighbor
        # ====================================================

        if line.startswith(
            "--------------------------------"
        ):

            save_neighbor()

            current = {

                "Device ID": "",
                "Neighbor IP": "",
                "Platform": "",
                "Capabilities": "",
                "Local Interface": "",
                "Neighbor Interface": "",
                "Holdtime": "",
                "Version": "",
                "Duplex": "",
                "Native VLAN": "",
                "Neighbor Hostname": "",
            }

            continue

        # ====================================================
        # Device ID
        # ====================================================

        if line.startswith(
            "Device-ID:"
        ):

            current["Device ID"] = (
                line.split(
                    "Device-ID:",
                    1
                )[1].strip()
            )

            continue

        # ====================================================
        # Platform
        # ====================================================

        if line.startswith(
            "Platform:"
        ):

            current["Platform"] = (
                line.split(
                    "Platform:",
                    1
                )[1].strip()
            )

            continue

        # ====================================================
        # Capabilities
        # ====================================================

        if line.startswith(
            "Capabilities:"
        ):

            current["Capabilities"] = (
                line.split(
                    "Capabilities:",
                    1
                )[1].strip()
            )

            continue

        # ====================================================
        # Interface
        # ====================================================

        if line.startswith(
            "Interface:"
        ):

            value = line.split(
                "Interface:",
                1
            )[1].strip()

            match = re.match(
                r"([^,]+),\s*"
                r"Port ID \(outgoing port\):\s*(.+)",
                value,
                re.IGNORECASE
            )

            if match:

                current[
                    "Local Interface"
                ] = match.group(
                    1
                ).strip()

                current[
                    "Neighbor Interface"
                ] = match.group(
                    2
                ).strip()

            else:

                current[
                    "Local Interface"
                ] = value

            continue

        # ====================================================
        # Holdtime
        # ====================================================

        if line.startswith(
            "Holdtime:"
        ):

            current["Holdtime"] = (
                line.split(
                    "Holdtime:",
                    1
                )[1].strip()
            )

            continue

        # ====================================================
        # Version
        # ====================================================

        if line.startswith(
            "Version:"
        ):

            current["Version"] = (
                line.split(
                    "Version:",
                    1
                )[1].strip()
            )

            continue

        # ====================================================
        # Duplex
        # ====================================================

        if line.startswith(
            "Duplex:"
        ):

            current["Duplex"] = (
                line.split(
                    "Duplex:",
                    1
                )[1].strip()
            )

            continue

        # ====================================================
        # Native VLAN
        # ====================================================

        if line.startswith(
            "Native VLAN:"
        ):

            current["Native VLAN"] = (
                line.split(
                    "Native VLAN:",
                    1
                )[1].strip()
            )

            continue

        # ====================================================
        # Neighbor system name
        # ====================================================

        if line.startswith(
            "SysName:"
        ):

            current[
                "Neighbor Hostname"
            ] = (
                line.split(
                    "SysName:",
                    1
                )[1].strip()
            )

            continue

        # ====================================================
        # IPv4 address
        #
        # CBS350:
        #
        # Addresses:
        #          IP 192.168.1..5
        # ====================================================

        ip_match = re.match(
            r"^\s*IP\s+"
            r"(\d+\.\d+\.\d+\.\d+)",
            line,
            re.IGNORECASE
        )

        if ip_match:

            current[
                "Neighbor IP"
            ] = ip_match.group(1)

            continue

    # ========================================================
    # Save final neighbor
    # ========================================================

    save_neighbor()

    return neighbors


# ============================================================
# PARSE LLDP
# ============================================================

def parse_lldp(
    output,
    hostname,
    ip
):

    neighbors = []

    current = {

        "Neighbor": "",
        "Neighbor IP": "",
        "Platform": "",
        "Local Interface": "",
        "Neighbor Interface": "",
        "System Description": "",
    }

    def save_neighbor():

        if not (
            current["Neighbor"]
            or
            current["Neighbor Interface"]
        ):
            return

        neighbors.append({

            "Local Hostname":
                hostname,

            "Local IP":
                ip,

            "Neighbor":
                current["Neighbor"],

            "Neighbor IP":
                current["Neighbor IP"],

            "Platform":
                current["Platform"],

            "Local Interface":
                current["Local Interface"],

            "Neighbor Interface":
                current[
                    "Neighbor Interface"
                ],

            "System Description":
                current[
                    "System Description"
                ],
        })

    for raw_line in output.splitlines():

        line = clean_text(
            raw_line
        )

        if not line:
            continue

        lower = line.lower()

        # ====================================================
        # Ignore pagination
        # ====================================================

        if (
            "more:" in lower
            or
            "<space>" in lower
            or
            "quit:" in lower
        ):
            continue

        # ====================================================
        # Common LLDP formats
        # ====================================================

        if (
            line.startswith(
                "Device ID:"
            )
            or
            line.startswith(
                "Device-ID:"
            )
            or
            line.startswith(
                "System Name:"
            )
        ):

            save_neighbor()

            if ":" in line:

                current["Neighbor"] = (
                    line.split(
                        ":",
                        1
                    )[1].strip()
                )

            continue

        if line.startswith(
            "Local Interface:"
        ):

            current[
                "Local Interface"
            ] = line.split(
                ":",
                1
            )[1].strip()

            continue

        if line.startswith(
            "Port ID:"
        ):

            current[
                "Neighbor Interface"
            ] = line.split(
                ":",
                1
            )[1].strip()

            continue

        if line.startswith(
            "Port Description:"
        ):

            current[
                "System Description"
            ] = line.split(
                ":",
                1
            )[1].strip()

            continue

        if line.startswith(
            "System Description:"
        ):

            current[
                "System Description"
            ] = line.split(
                ":",
                1
            )[1].strip()

            continue

        if line.startswith(
            "Management Address:"
        ):

            current[
                "Neighbor IP"
            ] = line.split(
                ":",
                1
            )[1].strip()

            continue

        # ====================================================
        # IPv4 address
        # ====================================================

        ip_match = re.search(
            r"\b(\d+\.\d+\.\d+\.\d+)\b",
            line
        )

        if (
            ip_match
            and
            not current["Neighbor IP"]
        ):

            current[
                "Neighbor IP"
            ] = ip_match.group(1)

    save_neighbor()

    return neighbors


# ============================================================
# COLLECT ONE SWITCH
# ============================================================

def collect_switch(
    ip,
    username,
    password
):

    transport = None
    channel = None

    result = {

        "success": False,

        "ip":
            ip,

        "switch":
            {},

        "ports":
            [],

        "cdp":
            [],

        "lldp":
            [],

        "error":
            "",
    }

    try:

        print(
            f"[START] Connecting to {ip}"
        )

        # ====================================================
        # SSH TRANSPORT
        # ====================================================

        transport = paramiko.Transport(
            (
                ip,
                SSH_PORT
            )
        )

        transport.start_client(
            timeout=15
        )

        print(
            f"[SSH OK] {ip}"
        )

        # ====================================================
        # CBS350 ACCEPTS NONE AUTHENTICATION
        # THEN PROMPTS FOR USERNAME/PASSWORD
        # ====================================================

        transport.auth_none(
            username
        )

        # ====================================================
        # OPEN INTERACTIVE SESSION
        # ====================================================

        channel = transport.open_session()

        channel.get_pty(
            term="vt100",
            width=200,
            height=50
        )

        channel.invoke_shell()

        login_output = read_login_output(
            channel,
            5
        )

        # ====================================================
        # USERNAME
        # ====================================================

        if (
            "User Name:" in login_output
            or
            "Username:" in login_output
            or
            "username:" in login_output
        ):

            channel.send(
                username + "\n"
            )

        else:

            raise Exception(
                "Username prompt not found"
            )

        # ====================================================
        # PASSWORD
        # ====================================================

        password_output = read_login_output(
            channel,
            5
        )

        if (
            "Password:" in
            password_output
        ):

            channel.send(
                password + "\n"
            )

        else:

            raise Exception(
                "Password prompt not found"
            )

        # ====================================================
        # WAIT FOR SWITCH PROMPT
        # ====================================================

        final_login = read_login_output(
            channel,
            5
        )

        if not re.search(
            r"[A-Za-z0-9_.:/()\-]+[>#]\s*$",
            final_login,
            re.MULTILINE
        ):

            raise Exception(
                "Cisco switch prompt not detected"
            )

        print(
            f"[LOGIN SUCCESS] {ip}"
        )

        # ====================================================
        # SHOW SYSTEM
        # ====================================================

        print(
            "[COMMAND] show system"
        )

        system_output = send_command(
            channel,
            "show system",
            2
        )

        system = parse_system(
            system_output
        )

        hostname = (
            system[
                "System Name"
            ]
            or
            "Unknown"
        )

        print(
            f"[HOSTNAME] {hostname}"
        )

        # ====================================================
        # SHOW VERSION
        # ====================================================

        print(
            "[COMMAND] show version"
        )

        version_output = send_command(
            channel,
            "show version",
            2
        )

        firmware = parse_version(
            version_output
        )

        # ====================================================
        # SHOW INVENTORY
        # ====================================================

        print(
            "[COMMAND] show inventory"
        )

        inventory_output = send_command(
            channel,
            "show inventory",
            2
        )

        inventory = parse_inventory(
            inventory_output
        )

        # ====================================================
        # PORT DESCRIPTIONS
        # ====================================================

        print(
            "[COMMAND] show interfaces description"
        )

        description_output = send_command(
            channel,
            "show interfaces description",
            2
        )

        descriptions = (
            parse_interface_descriptions(
                description_output
            )
        )

        print(
            f"[DESCRIPTIONS] "
            f"{len(descriptions)}"
        )

        # ====================================================
        # INTERFACE STATUS
        # ====================================================

        print(
            "[COMMAND] show interfaces status"
        )

        status_output = send_command(
            channel,
            "show interfaces status",
            2
        )

        ports = parse_interface_status(
            status_output,
            hostname,
            ip,
            descriptions
        )

        # ====================================================
        # ADD DESCRIPTION-ONLY PORTS
        #
        # This ensures ports that appear in
        # show interfaces description but not
        # in show interfaces status are not lost.
        # ====================================================

        existing = {

            row[
                "Interface"
            ].lower()

            for row in ports
        }

        for interface, description in (
            descriptions.items()
        ):

            if interface not in existing:

                ports.append({

                    "Hostname":
                        hostname,

                    "IP Address":
                        ip,

                    "Interface":
                        interface,

                    "Description":
                        description,

                    "Type":
                        "",

                    "Status":
                        "",

                    "Speed":
                        "",

                    "Duplex":
                        "",

                    "Negotiation":
                        "",

                    "Flow Control":
                        "",

                    "Back Pressure":
                        "",

                    "MDIX":
                        "",
                })

        # ====================================================
        # CDP
        # ====================================================

        print(
            "[COMMAND] "
            "show cdp neighbors detail"
        )

        cdp_output = send_command(
            channel,
            "show cdp neighbors detail",
            2
        )

        cdp = parse_cdp(
            cdp_output,
            hostname,
            ip
        )

        print(
            f"[CDP] "
            f"{len(cdp)} neighbors"
        )

        # ====================================================
        # LLDP
        # ====================================================

        print(
            "[COMMAND] "
            "show lldp neighbors detail"
        )

        lldp_output = send_command(
            channel,
            "show lldp neighbors detail",
            2
        )

        lldp = parse_lldp(
            lldp_output,
            hostname,
            ip
        )

        print(
            f"[LLDP] "
            f"{len(lldp)} neighbors"
        )

        # ====================================================
        # SWITCH RECORD
        # ====================================================

        switch_record = {

            "Hostname":
                hostname,

            "IP Address":
                ip,

            "Model":
                (
                    inventory["Model"]
                    or
                    system["System Model"]
                ),

            "Serial Number":
                inventory[
                    "Serial Number"
                ],

            "Product ID":
                inventory[
                    "Product ID"
                ],

            "Hardware Version":
                inventory[
                    "Hardware Version"
                ],

            "Firmware Version":
                firmware,

            "System Description":
                system[
                    "System Description"
                ],

            "Uptime":
                system[
                    "System Up Time"
                ],

            "System Contact":
                system[
                    "System Contact"
                ],

            "System Location":
                system[
                    "System Location"
                ],

            "MAC Address":
                system[
                    "System MAC Address"
                ],

            "System Object ID":
                system[
                    "System Object ID"
                ],
        }

        # ====================================================
        # SAVE RESULTS
        # ====================================================

        result["success"] = True

        result["switch"] = (
            switch_record
        )

        result["ports"] = ports

        result["cdp"] = cdp

        result["lldp"] = lldp

        print()

        print(
            f"[SUCCESS] "
            f"{hostname} ({ip})"
        )

        print(
            f"          Model: "
            f"{switch_record['Model']}"
        )

        print(
            f"          Serial: "
            f"{switch_record['Serial Number']}"
        )

        print(
            f"          Firmware: "
            f"{switch_record['Firmware Version']}"
        )

        print(
            f"          Port records: "
            f"{len(ports)}"
        )

        print(
            f"          Descriptions: "
            f"{len(descriptions)}"
        )

        print(
            f"          CDP neighbors: "
            f"{len(cdp)}"
        )

        print(
            f"          LLDP neighbors: "
            f"{len(lldp)}"
        )

    except Exception as error:

        result["error"] = clean_text(
            str(error)
        )

        print()

        print(
            f"[FAILED] {ip}"
        )

        print(
            f"         {error}"
        )

    finally:

        if channel:

            try:
                channel.close()
            except:
                pass

        if transport:

            try:
                transport.close()
            except:
                pass

    return result


# ============================================================
# FORMAT EXCEL WORKSHEETS
# ============================================================

def format_excel(writer):

    workbook = writer.book

    for worksheet in workbook.worksheets:

        # Freeze first row
        worksheet.freeze_panes = "A2"

        # Autofilter
        if worksheet.max_row > 1:

            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

        # Column width
        for column in worksheet.columns:

            max_length = 0

            column_letter = (
                column[0].column_letter
            )

            for cell in column:

                if cell.value is not None:

                    length = len(
                        str(cell.value)
                    )

                    if length > max_length:

                        max_length = length

            worksheet.column_dimensions[
                column_letter
            ].width = min(
                max_length + 2,
                60
            )


# ============================================================
# MAIN
# ============================================================

def main():

    print()

    print(
        "=" * 70
    )

    print(
        "       CISCO NETWORK DISCOVERY & INVENTORY"
    )

    print(
        "=" * 70
    )

    print()

    print(
        f"Switches configured: "
        f"{len(SWITCH_IPS)}"
    )

    print(
        f"Parallel connections: "
        f"{MAX_WORKERS}"
    )

    print()

    print(
        "Excel report will be saved to:"
    )

    print(
        EXCEL_FILE
    )

    print()

    # ========================================================
    # CREDENTIALS
    # ========================================================

    username = input(
        "Enter Cisco Username: "
    )

    password = getpass.getpass(
        "Enter Cisco Password: "
    )

    print()

    print(
        "Starting discovery..."
    )

    print()

    results = []

    # ========================================================
    # PARALLEL DISCOVERY
    # ========================================================

    with ThreadPoolExecutor(
        max_workers=MAX_WORKERS
    ) as executor:

        futures = {}

        for ip in SWITCH_IPS:

            future = executor.submit(
                collect_switch,
                ip,
                username,
                password
            )

            futures[future] = ip

        for future in as_completed(
            futures
        ):

            ip = futures[future]

            try:

                results.append(
                    future.result()
                )

            except Exception as error:

                results.append({

                    "success":
                        False,

                    "ip":
                        ip,

                    "switch":
                        {},

                    "ports":
                        [],

                    "cdp":
                        [],

                    "lldp":
                        [],

                    "error":
                        str(error),
                })

    # ========================================================
    # PREPARE DATA
    # ========================================================

    switch_rows = []

    port_rows = []

    cdp_rows = []

    lldp_rows = []

    failed_rows = []

    for result in results:

        if result["success"]:

            switch_rows.append(
                result["switch"]
            )

            port_rows.extend(
                result["ports"]
            )

            cdp_rows.extend(
                result["cdp"]
            )

            lldp_rows.extend(
                result["lldp"]
            )

        else:

            failed_rows.append({

                "IP Address":
                    result["ip"],

                "Error":
                    result["error"],
            })

    # ========================================================
    # DATAFRAMES
    # ========================================================

    df_switch = pd.DataFrame(
        switch_rows
    )

    df_ports = pd.DataFrame(
        port_rows
    )

    df_cdp = pd.DataFrame(
        cdp_rows
    )

    df_lldp = pd.DataFrame(
        lldp_rows
    )

    df_failed = pd.DataFrame(
        failed_rows
    )

    # ========================================================
    # CLEAN DATA
    # ========================================================

    df_switch = clean_dataframe(
        df_switch
    )

    df_ports = clean_dataframe(
        df_ports
    )

    df_cdp = clean_dataframe(
        df_cdp
    )

    df_lldp = clean_dataframe(
        df_lldp
    )

    df_failed = clean_dataframe(
        df_failed
    )

    # ========================================================
    # STATISTICS
    # ========================================================

    total_switches = len(
        SWITCH_IPS
    )

    successful = len(
        switch_rows
    )

    failed = len(
        failed_rows
    )

    total_ports = len(
        port_rows
    )

    ports_with_description = sum(

        1

        for row in port_rows

        if row.get(
            "Description",
            ""
        ).strip()
    )

    connected_ports = sum(

        1

        for row in port_rows

        if row.get(
            "Status",
            ""
        ).lower()
        == "connected"
    )

    cdp_count = len(
        cdp_rows
    )

    lldp_count = len(
        lldp_rows
    )

    # ========================================================
    # SUMMARY
    # ========================================================

    df_summary = pd.DataFrame({

        "Metric": [

            "Report Generated",

            "Total Switches",

            "Successful Switches",

            "Failed Switches",

            "Total Port Records",

            "Ports With Description",

            "Connected Ports",

            "CDP Neighbors",

            "LLDP Neighbors",

            "Excel File",
        ],

        "Value": [

            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),

            total_switches,

            successful,

            failed,

            total_ports,

            ports_with_description,

            connected_ports,

            cdp_count,

            lldp_count,

            str(EXCEL_FILE),
        ],
    })

    # ========================================================
    # CREATE EXCEL
    # ========================================================

    print()

    print(
        "Creating Excel report..."
    )

    with pd.ExcelWriter(
        EXCEL_FILE,
        engine="openpyxl"
    ) as writer:

        # ====================================================
        # SUMMARY
        # ====================================================

        df_summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        # ====================================================
        # SWITCH INVENTORY
        # ====================================================

        df_switch.to_excel(
            writer,
            sheet_name="Switch Inventory",
            index=False
        )

        # ====================================================
        # PORT INVENTORY
        # ====================================================

        df_ports.to_excel(
            writer,
            sheet_name="Port Inventory",
            index=False
        )

        # ====================================================
        # CDP
        # ====================================================

        df_cdp.to_excel(
            writer,
            sheet_name="CDP Neighbors",
            index=False
        )

        # ====================================================
        # LLDP
        # ====================================================

        df_lldp.to_excel(
            writer,
            sheet_name="LLDP Neighbors",
            index=False
        )

        # ====================================================
        # FAILED DEVICES
        # ====================================================

        df_failed.to_excel(
            writer,
            sheet_name="Failed Devices",
            index=False
        )

        # ====================================================
        # FORMAT
        # ====================================================

        format_excel(
            writer
        )

    # ========================================================
    # FINAL RESULT
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "             SCAN COMPLETED"
    )

    print(
        "=" * 70
    )

    print()

    print(
        f"Total switches        : "
        f"{total_switches}"
    )

    print(
        f"Successful             : "
        f"{successful}"
    )

    print(
        f"Failed                 : "
        f"{failed}"
    )

    print(
        f"Total port records     : "
        f"{total_ports}"
    )

    print(
        f"Ports with description : "
        f"{ports_with_description}"
    )

    print(
        f"Connected ports        : "
        f"{connected_ports}"
    )

    print(
        f"CDP neighbors          : "
        f"{cdp_count}"
    )

    print(
        f"LLDP neighbors         : "
        f"{lldp_count}"
    )

    print()

    print(
        "Excel report created successfully!"
    )

    print()

    print(
        "Saved to:"
    )

    print(
        EXCEL_FILE
    )

    print()

    print(
        "The Excel file has been saved to the Desktop of the"
    )

    print(
        "Windows user account currently running this program."
    )

    print()
    print("Opening Excel report automatically...")

    try:
        os.startfile(str(EXCEL_FILE))
        print("[OK] Excel report opened successfully.")
    except Exception as error:
        print("[WARNING] Could not open Excel automatically.")
        print(f"          {error}")

    print()

    print(
        "=" * 70
    )

    print()
    print("SCAN FINISHED.")
    print("The Excel report is saved at:")
    print(EXCEL_FILE)
    print()
    print("This window will remain open.")
    input("Press ENTER to close this window...")


# ============================================================
# START PROGRAM
# ============================================================

if __name__ == "__main__":

    try:
        main()
    except Exception as error:
        print()
        print("=" * 70)
        print("PROGRAM ERROR")
        print("=" * 70)
        print()
        print(error)
        print()
        input("Press ENTER to close this window...")
```

---

# 📎 References

The script uses the following Python packages:

```text
Paramiko
pandas
openpyxl
```

The script also relies on Cisco CBS350 / CBS Series CLI commands and their corresponding command-output formats.

---

# 📌 Document Scope

This document describes the supplied Cisco Network Discovery and Inventory Python script, including its configuration, execution process, Cisco commands, parsing logic, Excel reporting, parallel processing, error handling, and complete source code.

The Python source code has been retained in its entirety as provided.
