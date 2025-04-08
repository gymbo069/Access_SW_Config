import pandas as pd
from datetime import datetime
import os


"""
 What You Need To Do:

    Save your Excel/CSV file as your_switch_data.xlsx in the same folder.

    Run this script using Python 3.

    It will:

        Create one .txt file per switch.

        Zip them into: YYYY-MM-DD_MES_access_switch_config.zip
"""

# Define VLAN mapping per zone
zone_config = {
    "Elec": {
        "vlan": 2520,
        "vlan_name": "MES_ELEC_10.99.52.0/22",
        "subnet_mask": "255.255.252.0",
        "gateway": "10.99.52.1"
    },
    "Assem": {
        "vlan": 2560,
        "vlan_name": "MES_ASSEM_10.99.56.0/21",
        "subnet_mask": "255.255.248.0",
        "gateway": "10.99.56.1"
    },
    "Form": {
        "vlan": 2640,
        "vlan_name": "MES_FORM_10.99.64.0/21",
        "subnet_mask": "255.255.248.0",
        "gateway": "10.99.64.1"
    }
}

# Load the data
filename = "GPT 컨피그용 Access SW 정보.xlsx"  # or .csv
df = pd.read_excel(filename)  # or pd.read_csv(filename)

# Clean column names to remove leading/trailing spaces
df.columns = df.columns.str.strip()

# Print the column names to check
print(df.columns)

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
    uplink = row["uplink"]
    po = row["po"]

    vlan_info = zone_config[zone]
    vlan = vlan_info["vlan"]
    vlan_name = vlan_info["vlan_name"]
    mask = vlan_info["subnet_mask"]
    gateway = vlan_info["gateway"]  # Get the gateway address

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
 description ## {zone} VLAN ##
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

    # Save to file
    with open(os.path.join(output_dir, f"{hostname}.txt"), "w") as f:
        f.write(config)

# Optionally, compress into a zip
import zipfile

zip_filename = f"{output_dir}.zip"
with zipfile.ZipFile(zip_filename, 'w') as zipf:
    for file in os.listdir(output_dir):
        zipf.write(os.path.join(output_dir, file), arcname=file)

print(f"Configs generated and zipped as: {zip_filename}")
