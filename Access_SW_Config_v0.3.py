import pandas as pd
from datetime import datetime
import os
import zipfile
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import re

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
    df.columns = df.columns.str.strip()  # Clean column names
except Exception as e:
    messagebox.showerror("Error Reading Excel", str(e))
    exit()

# Extract zone info from Excel and prompt user for VLAN settings
zone_names = df['zone'].dropna().unique()
zone_config = {}

for zone in zone_names:
    vlan = simpledialog.askinteger("Zone Config", f"Enter VLAN ID for zone '{zone}':")
    vlan_name = simpledialog.askstring("Zone Config", f"Enter VLAN Name for zone '{zone}':")
    subnet_mask = simpledialog.askstring("Zone Config", f"Enter Subnet Mask for zone '{zone}':")
    gateway = simpledialog.askstring("Zone Config", f"Enter Gateway IP for zone '{zone}':")

    zone_config[zone] = {
        "vlan": vlan,
        "vlan_name": vlan_name,
        "subnet_mask": subnet_mask,
        "gateway": gateway
    }

# Ask uplink descriptions for each port
uplink1_desc = simpledialog.askstring("Uplink Description", "Enter uplink description for TenGigabitEthernet1/1/1 (e.g., oabbsw1):")
uplink2_desc = simpledialog.askstring("Uplink Description", "Enter uplink description for TenGigabitEthernet1/1/2 (e.g., oabbsw2):")

# Derive common base for Port-Channel by stripping numbers from the end
def strip_numbers(desc):
    return re.sub(r"\d+$", "", desc)

portchannel_desc = strip_numbers(uplink1_desc)

# Output folder
today = datetime.now().strftime("%Y-%m-%d")
output_dir = f"{today}_MES_access_switch_config"
os.makedirs(output_dir, exist_ok=True)

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
 description ## {portchannel_desc} Po({po}) ##
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
