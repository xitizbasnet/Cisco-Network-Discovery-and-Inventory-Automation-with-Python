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
