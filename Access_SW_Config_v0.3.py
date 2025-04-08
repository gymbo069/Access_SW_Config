import pandas as pd
from datetime import datetime
import os
import zipfile
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import re
import ipaddress

# GUI setup
root = tk.Tk()
root.withdraw()

# Ask user to select Excel file
excel_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
if not excel_path:
    messagebox.showerror("Error", "No Excel file selected.")
    exit()

# Read Excel file
try:
    df = pd.read_excel(excel_path)
    df.columns = df.columns.str.strip().str.lower()  # Clean column names and convert to lowercase
except Exception as e:
    messagebox.showerror("Error Reading Excel", str(e))
    exit()

# Ensure required columns are present (case-insensitive)
required_columns = ['hostname', 'ip', 'port', 'zone', 'po', 'uplink', 'netmask']
missing_columns = [col for col in required_columns if col not in df.columns]

if missing_columns:
    messagebox.showerror("Missing Columns", f"The following required columns are missing: {', '.join(missing_columns)}")
    exit()

# Show columns in a pop-up message box
column_names = ', '.join(df.columns)
messagebox.showinfo("Columns in the Excel file", f"Columns found in the Excel file:\n{column_names}")

# Function to convert CIDR to standard subnet mask
def cidr_to_netmask(cidr):
    return str(ipaddress.IPv4Network(f"0.0.0.0/{cidr}", strict=False).netmask)



# Output folder
today = datetime.now().strftime("%Y-%m-%d")
output_dir = f"{today}_Access_switch_config"
os.makedirs(output_dir, exist_ok=True)

# Extract zone info from Excel and prompt user for VLAN settings
zone_names = df['zone'].dropna().unique()
zone_config = {}

for zone in zone_names:
    vlan = simpledialog.askinteger("Zone Config", f"Enter VLAN ID for zone '{zone}':")
    vlan_name = simpledialog.askstring("Zone Config", f"Enter VLAN Name for zone '{zone}':")

    # Get the 'netmask' column case-insensitively
    subnet_cidr = df.loc[df['zone'] == zone, 'netmask'].iloc[0]  # Column names are now all lowercase
    subnet_mask = cidr_to_netmask(subnet_cidr.split('/')[1])  # Convert CIDR to standard subnet mask
    gateway = simpledialog.askstring("Zone Config", f"Enter Gateway IP for zone '{zone}':")

    # Validate the gateway is in proper IPv4 format
    try:
        ipaddress.IPv4Address(gateway)
    except ipaddress.AddressValueError:
        messagebox.showerror("Invalid Gateway", f"The entered gateway IP '{gateway}' is not a valid IPv4 address.")
        exit()

    zone_config[zone] = {
        "vlan": vlan,
        "vlan_name": vlan_name,
        "subnet_mask": subnet_mask,
        "gateway": gateway
    }

# Function to strip numbers from a string
def strip_numbers(desc):
    return re.sub(r"\d+$", "", desc)



# Function to check if uplink descriptions are equal after stripping numbers
def check_uplink_desc_equal(uplink1_desc, uplink2_desc):
    stripped_uplink1 = strip_numbers(uplink1_desc)
    stripped_uplink2 = strip_numbers(uplink2_desc)
    if stripped_uplink1 != stripped_uplink2:
        messagebox.showerror("Uplink Description Mismatch", 
                             f"The stripped uplink descriptions do not match:\n"
                             f"Uplink 1: {stripped_uplink1}\nUplink 2: {stripped_uplink2}")
        return False
    return True

# Ask uplink descriptions for each port
uplink1_desc = simpledialog.askstring("Uplink Description", "Enter uplink description for TenGigabitEthernet1/1/1 (e.g., oabbsw1):")
uplink2_desc = simpledialog.askstring("Uplink Description", "Enter uplink description for TenGigabitEthernet1/1/2 (e.g., oabbsw2):")

# Check if the stripped uplink descriptions are equal
if not check_uplink_desc_equal(uplink1_desc, uplink2_desc):
    exit()
# Derive common base for Port-Channel by stripping numbers from the end
portchannel_desc = strip_numbers(uplink1_desc)

# Generate config for each switch
for _, row in df.iterrows():
    hostname = row["hostname"]
    ip = row["ip"]
    port = row["port"]
    zone = row["zone"]
    po = row["po"]
    uplink = row["uplink"]

    vlan_info = zone_config[zone]
    vlan = vlan_info["vlan"]
    vlan_name = vlan_info["vlan_name"]
    mask = vlan_info["subnet_mask"]
    gateway = vlan_info["gateway"]

    # Access port range
    port_range = "1-24" if port == 24 else "1-48"

    config = f"""conf t
hostname {hostname}

clock timezone EST -5 0
clock summer-time EDT recurring 2 Sun Mar 02:00 1 Sun Nov 02:00 60

service timestamps debug datetime msec localtime show-timezone
service timestamps log datetime msec localtime show-timezone
service password-encryption

logging buffer 10240
license smart transport off

transceiver type all
 monitoring

no ip domain lookup
ip domain name skcc.com
crypto key generate rsa modulus 1024

vtp mode transparent

username skcc privilege 15 password m5enHapppy1009(
enable secret m5enHapppy1009(

vlan {vlan}
 name {vlan_name}

interface Vlan{vlan}
 description ## {zone} VLAN - {vlan_name} ##
 ip address {ip} {mask}
 no shutdown
 ip default-gateway {gateway}

interface range GigabitEthernet1/0/{port_range}
 switchport mode access
 switchport access vlan {vlan}
 spanning-tree portfast
 spanning-tree bpduguard enable
 load-interval 30
 negotiation auto
 no shutdown

interface TenGigabitEthernet1/1/1
 description ## {uplink1_desc} ({uplink}) ##
interface TenGigabitEthernet1/1/2
 description ## {uplink2_desc} ({uplink}) ##

interface range TenGigabitEthernet1/1/1-2
 channel-group 10 mode active

interface Port-channel10
 description ## {portchannel_desc} (Po{po}) ##
 switchport mode trunk
 switchport trunk allowed vlan {vlan}

spanning-tree pathcost method short

line con 0
 logging syn

line vty 0 31
 login local
 transport input telnet ssh
 logging sync

interface range TenGigabitEthernet1/1/1-4
 load-interval 30
 negotiation auto
 no shutdown

end
"""

    with open(os.path.join(output_dir, f"{hostname}.txt"), "w") as f:
        f.write(config)

# Zip output
zip_filename = f"{output_dir}.zip"
with zipfile.ZipFile(zip_filename, 'w') as zipf:
    for file in os.listdir(output_dir):
        zipf.write(os.path.join(output_dir, file), arcname=file)

messagebox.showinfo("Done", f"Configs generated and zipped as: {zip_filename}")
