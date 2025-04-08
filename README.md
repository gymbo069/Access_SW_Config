Here’s a complete README file that combines the script explanation and the example Excel table.

---

# Access Switch Configuration Script

## Overview

This Python script is designed to generate Cisco access switch configuration files based on data provided in an Excel spreadsheet. The script prompts the user for additional details and generates configuration files for each switch. It also allows customization of VLANs, subnet masks, uplinks, and port-channel descriptions.

### Features:
- Reads switch configuration data from an Excel file.
- Converts CIDR subnet masks to standard notation.
- Validates gateway IPs for correctness.
- Ensures consistency in uplink descriptions.
- Generates configuration files for each switch, including VLANs, IP addresses, and port configurations.
- Zips all generated configuration files into a single compressed file for easy distribution.

---

## Requirements

- Python 3.x
- Required Python Libraries:
  - pandas
  - tkinter
  - ipaddress
  - re
  - zipfile
  - os

You can install the necessary libraries using the following command:
```
pip install pandas
```

---

## How to Use

1. **Prepare the Excel File:**
   - The script expects an Excel file with specific columns. The column names must match the following (case-insensitive):

| **hostname**  | **ip**           | **port** | **zone**  | **po** | **uplink**  | **netmask**     |
|---------------|------------------|----------|-----------|--------|-------------|-----------------|
| Switch1       | 192.168.1.10     | 1        | Sales     | 10     | uplink1     | 192.168.1.0/24 |
| Switch1       | 192.168.1.11     | 2        | Sales     | 10     | uplink1     | 192.168.1.0/24 |
| Switch2       | 192.168.2.10     | 1        | HR        | 20     | uplink2     | 192.168.2.0/24 |
| Switch2       | 192.168.2.11     | 2        | HR        | 20     | uplink2     | 192.168.2.0/24 |
| Switch3       | 192.168.3.10     | 1        | Marketing | 30     | uplink3     | 192.168.3.0/24 |
| Switch3       | 192.168.3.11     | 2        | Marketing | 30     | uplink3     | 192.168.3.0/24 |

  **Explanation of Columns:**
  
    - **hostname**: The name of the switch.
    - **ip**: The IP address of the switch interface.
    - **port**: The port number (used for configuring port ranges, typically 1-24 or 1-48).
    - **zone**: The zone name (used for VLAN configurations).
    - **po**: The Port-Channel number.
    - **uplink**: The description of the uplink interface (used in generating port-channel configurations).
    - **netmask**: The subnet in CIDR format (e.g., 192.168.1.0/24).
      
2. **Run the Script:**
   - The script will prompt you to select the Excel file containing your switch configuration data.
   - It will then ask for VLAN IDs, VLAN names, and gateway IPs for each zone.
   - Next, you’ll be asked for uplink descriptions. The script will check if the uplink descriptions match after removing numeric suffixes (to ensure consistency).
   
3. **Configuration Files:**
   - The script will generate a configuration file for each switch based on the details in the Excel file and the user inputs.
   - These configuration files will be stored in a folder named with the current date (e.g., `2025-04-08_Access_switch_config`).
   - The folder will be zipped into a single file for easy distribution.

---

## How the Script Works

1. **Reads Excel File:**
   The script uses the `pandas` library to read the Excel file. It expects the following columns: `hostname`, `ip`, `port`, `zone`, `po`, `uplink`, and `netmask`.

2. **Column Case Insensitivity:**
   The column names are automatically converted to lowercase, ensuring that they are case-insensitive.

3. **CIDR to Standard Subnet Mask Conversion:**
   The script converts the CIDR format in the `netmask` column to a standard subnet mask.

4. **Gateway Validation:**
   The script validates that the entered gateway IP addresses are in correct IPv4 format. If they are not, the user is prompted with an error.

5. **Uplink Description Consistency Check:**
   The script checks if the uplink descriptions for the two ports (`TenGigabitEthernet1/1/1` and `TenGigabitEthernet1/1/2`) are consistent. If they are not, the script will terminate and display an error message.

6. **Configuration File Generation:**
   The script generates the configuration files for each switch, ensuring the configuration includes:
   - Hostname
   - VLAN configuration (VLAN ID, Name, IP Address, Subnet Mask)
   - Port settings (Access Ports)
   - Uplink descriptions
   - Port-Channel configuration

7. **Zipping the Output:**
   After generating the configuration files, the script will zip the folder containing all the generated files and display the location of the zipped file.

---

## Example of Generated Configuration

Here’s an example of a generated configuration file for a switch:

```
conf t
hostname Switch1

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

vlan 10
 name Sales

interface Vlan10
 description ## Sales VLAN - Sales ##
 ip address 192.168.1.10 255.255.255.0
 no shutdown
 ip default-gateway 192.168.1.1

interface range GigabitEthernet1/0/1-24
 switchport mode access
 switchport access vlan 10
 spanning-tree portfast
 spanning-tree bpduguard enable
 load-interval 30
 negotiation auto
 no shutdown

interface TenGigabitEthernet1/1/1
 description ## uplink1 (uplink1) ##
interface TenGigabitEthernet1/1/2
 description ## uplink1 (uplink1) ##

interface range TenGigabitEthernet1/1/1-2
 channel-group 10 mode active

interface Port-channel10
 description ## uplink1 (Po10) ##
 switchport mode trunk
 switchport trunk allowed vlan 10

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
```

---

## Notes

- Ensure that your Excel file is formatted correctly before running the script.
- The script will validate the data entered for gateways and uplinks, ensuring consistency across configurations.
- The final output will be a zip file containing configuration files for each switch.

---

## License

This script is provided under the MIT License. Feel free to modify and adapt it to your needs.

---

Let me know if you need further adjustments or clarifications for the README!
