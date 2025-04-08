import pandas as pd
from datetime import datetime
import os
import zipfile

# =====================
# STEP 1: COLLECT ZONE CONFIGURATIONS
# =====================

# Prompt user to enter the names of all zones (comma-separated)
zones = input("Enter all zone names separated by commas (e.g., Elec,Assem,Form): ").split(',')

# Dictionary to store VLAN details for each zone
zone_config = {}

# Ask user to enter VLAN, VLAN Name, Subnet Mask, and Gateway for each zone
for zone in zones:
    zone = zone.strip()
    vlan = input(f"[{zone}] VLAN ID: ")
    vlan_name = input(f"[{zone}] VLAN Name: ")
    subnet_mask = input(f"[{zone}] Subnet Mask (e.g., 255.255.252.0): ")
    gateway = input(f"[{zone}] Default Gateway: ")

    zone_config[zone] = {
        "vlan": vlan,
        "vlan_name": vlan_name,
        "subnet_mask": subnet_mask,
        "gateway": gateway
    }

# =====================
# STEP 2: LOAD SWITCH DATA
# =====================

# Prompt for the Excel or CSV file that contains switch info
filename = input("\nEnter the path to your switch data Excel or CSV file: ")
if filename.endswith(".csv"):
    df = pd.read_csv(filename)
else:
    df = pd.read_excel(filename)

# =====================
# STEP 3: CREATE CONFIGS FOR EACH SWITCH
# =====================

# Set output folder name with today's date
today = datetime.now().strftime("%Y-%m-%d")
output_dir = f"{today}_MES_access_switch_config"
os.makedirs(output_dir, exist_ok=True)

# Loop through each row in the spreadsheet
for _, row in df.iterrows():
    # Extract switch details
    hostname = row["hostname"]
    ip = row["ip"]
    port = row["port"]
    zone = row["zone"]
    uplink = row["uplink"]
    po = row["po"]

    # Get VLAN info for the switch's zone
    vlan_info = zone_config[zone]
    vlan = vlan_info["vlan"]
    vlan_name = vlan_info["vlan_name"]
    mask = vlan_info["subnet_mask"]

    # Determine access port range based on switch model
    port_range = "1-24" if port == 24 else "1-48"

    # Begin writing the Cisco IOS-XE configuration
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
 description ## {zone} VLAN ##
 ip address {ip} {mask}
 no shutdown

interface range GigabitEthernet1/0/{port_range}
 switchport mode access
 switchport access vlan {vlan}
 spanning-tree portfast
 spanning-tree bpduguard enable
 load-interval 30
 negotiation auto
 no shutdown

interface TenGigabitEthernet1/1/1
 description ## HS_MESBB_SW_N9508_1 ({uplink}) ##
interface TenGigabitEthernet1/1/2
 description ## HS_MESBB_SW_N9508_2 ({uplink}) ##

interface range TenGigabitEthernet1/1/1-2
 channel-group 10 mode active

interface Port-channel10
 description ## HS_MESBB_SW_N9508 (Po{po}) ##
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

    # Save configuration to a .txt file for the current switch
    with open(os.path.join(output_dir, f"{hostname}.txt"), "w") as f:
        f.write(config)

# =====================
# STEP 4: ZIP CONFIG FILES
# =====================

# Create a ZIP archive containing all generated config files
zip_filename = f"{output_dir}.zip"
with zipfile.ZipFile(zip_filename, 'w') as zipf:
    for file in os.listdir(output_dir):
        zipf.write(os.path.join(output_dir, file), arcname=file)

# =====================
# DONE
# =====================
print(f"\n✅ Configuration files saved and zipped as: {zip_filename}")
