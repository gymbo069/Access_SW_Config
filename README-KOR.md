---

# 액세스 스위치 구성 스크립트

## 개요

이 Python 스크립트는 Excel 스프레드시트에 제공된 데이터를 기반으로 Cisco 액세스 스위치 구성 파일을 생성하도록 설계되었습니다. 사용자는 추가 정보를 입력하도록 안내받으며, 각 스위치에 대한 구성 파일이 생성됩니다. VLAN, 서브넷 마스크, 업링크 및 포트채널 설명도 사용자 지정이 가능합니다.

---

## ⚠️ 주의사항 / Caution

- 해당 스크립트는 **VPC 양팔 구조(dual-homed)** Access 스위치 설정 시에만 유용합니다.
- 단일 업링크 또는 이중화 구성 없이 단독 포트채널을 사용하는 경우에는 사용하지 않는 것이 좋습니다.
- Excel 파일은 반드시 올바른 형식과 데이터로 작성되어야 하며, 오타나 누락된 정보가 있는 경우 구성 생성이 실패할 수 있습니다.
- 게이트웨이, VLAN 정보, 업링크 설명 등 사용자 입력값은 구성의 핵심 요소이므로 정확하게 입력해야 합니다.
- 생성된 구성 파일은 최종 적용 전에 반드시 검토하시기 바랍니다.

---

- This script is designed **specifically for dual-homed (VPC-style)** access switch configurations.
- It is **not recommended** for setups with a single uplink or without a redundant Port-Channel architecture.
- The Excel file must be formatted properly and contain all required data. Missing or incorrect information can lead to failure during config generation.
- User inputs such as gateway IPs, VLAN info, and uplink descriptions are critical and must be entered correctly.
- Always review the generated configuration files before applying them to live network equipment.

---

### 주요 기능:
- Excel 파일에서 스위치 구성 데이터를 읽어옵니다.
- CIDR 서브넷 마스크를 표준 형식으로 변환합니다.
- 게이트웨이 IP의 유효성을 검사합니다.
- 업링크 설명의 일관성을 확인합니다.
- VLAN, IP 주소, 포트 설정 등을 포함한 각 스위치 구성 파일을 생성합니다.
- 생성된 모든 구성 파일을 하나의 압축 파일로 묶어 배포를 간편하게 합니다.

---

## 요구 사항

- Python 3.x
- 필수 Python 라이브러리:
  - pandas
  - tkinter
  - ipaddress
  - re
  - zipfile
  - os

다음 명령어로 필요한 라이브러리를 설치할 수 있습니다:
```
pip install pandas
```

---

## 사용 방법

1. **Excel 파일 준비:**
   - 스크립트는 특정 열 이름을 가진 Excel 파일을 기대합니다. 열 이름은 아래 표와 같아야 하며, 대소문자는 구분하지 않습니다:

| **hostname**  | **ip**           | **port** | **zone**  | **po** | **uplink**  | **netmask**     |
|---------------|------------------|----------|-----------|--------|-------------|-----------------|
| Switch1       | 192.168.1.10     | 1        | Sales     | 10     | uplink1     | 192.168.1.0/24 |
| Switch1       | 192.168.1.11     | 2        | Sales     | 10     | uplink1     | 192.168.1.0/24 |
| Switch2       | 192.168.2.10     | 1        | HR        | 20     | uplink2     | 192.168.2.0/24 |
| Switch2       | 192.168.2.11     | 2        | HR        | 20     | uplink2     | 192.168.2.0/24 |
| Switch3       | 192.168.3.10     | 1        | Marketing | 30     | uplink3     | 192.168.3.0/24 |
| Switch3       | 192.168.3.11     | 2        | Marketing | 30     | uplink3     | 192.168.3.0/24 |

  **열 설명:**
  
    - **hostname**: 스위치의 이름
    - **ip**: 스위치 인터페이스의 IP 주소
    - **port**: 포트 번호 (일반적으로 1-24 또는 1-48 범위)
    - **zone**: VLAN 구성 시 사용할 존 이름
    - **po**: 포트 채널 번호
    - **uplink**: 업링크 인터페이스 설명 (포트채널 구성에 사용)
    - **netmask**: CIDR 형식의 서브넷 (예: 192.168.1.0/24)
      
2. **스크립트 실행:**
   - 스크립트 실행 시 스위치 구성 데이터가 포함된 Excel 파일을 선택하라는 창이 뜹니다.
   - 이후 각 존에 대한 VLAN ID, VLAN 이름, 게이트웨이 IP를 입력하도록 안내합니다.
   - 마지막으로 업링크 설명을 입력하며, 숫자 접미사를 제거한 후 서로 일치하는지 확인합니다.

3. **구성 파일 생성:**
   - 입력된 정보와 Excel 데이터를 기반으로 각 스위치에 대한 구성 파일이 생성됩니다.
   - 구성 파일은 현재 날짜가 포함된 폴더(예: `2025-04-08_Access_switch_config`)에 저장됩니다.
   - 해당 폴더는 자동으로 압축(zip)되어 하나의 파일로 생성됩니다.

---

## 스크립트 작동 방식

1. **Excel 파일 읽기:**
   `pandas` 라이브러리를 사용하여 Excel 파일을 읽습니다. 필수 열은 `hostname`, `ip`, `port`, `zone`, `po`, `uplink`, `netmask`입니다.

2. **열 이름 대소문자 무시:**
   열 이름은 자동으로 소문자로 변환되어 대소문자 구분 없이 처리됩니다.

3. **CIDR → 서브넷 마스크 변환:**
   `netmask` 열의 CIDR 형식을 일반적인 서브넷 마스크 형식으로 변환합니다.

4. **게이트웨이 유효성 검사:**
   사용자가 입력한 게이트웨이 IP가 올바른 IPv4 형식인지 확인하며, 유효하지 않으면 오류 메시지를 표시합니다.

5. **업링크 설명 일관성 검사:**
   두 포트(`TenGigabitEthernet1/1/1`, `TenGigabitEthernet1/1/2`)의 업링크 설명이 일치하는지 확인합니다. 불일치 시 스크립트가 종료됩니다.

6. **구성 파일 생성:**
   각 스위치에 대한 구성 파일은 다음을 포함합니다:
   - 호스트네임
   - VLAN 구성 (VLAN ID, 이름, IP 주소, 서브넷 마스크)
   - 액세스 포트 설정
   - 업링크 설명
   - 포트채널 구성

7. **압축 파일 생성:**
   모든 구성 파일이 포함된 폴더를 zip 파일로 압축하고, 생성 위치를 표시합니다.

---

## 구성 예시

다음은 스크립트로 생성된 스위치 구성 예시입니다:

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

## 참고 사항

- Excel 파일이 올바른 형식으로 작성되었는지 사전에 반드시 확인하세요.
- 게이트웨이 및 업링크 입력 값은 스크립트에서 유효성 및 일관성을 검사합니다.
- 최종 출력은 모든 스위치 구성 파일이 포함된 zip 파일입니다.

---

## 라이선스

이 스크립트는 MIT 라이선스를 따릅니다. 자유롭게 수정 및 활용하실 수 있습니다.

---
