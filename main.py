import os
from lxml import etree
import wmi
import mysql.connector
from pprint import pprint
import json

cim = wmi.WMI()
data = {}


def upload_to_db(data):
    pass


def show_software_installed():

    software_data = {}
    for product in cim.Win32_Product():

        software_data[product.Caption] = {}
        # print("Package name: " + product.Caption)
        software_data[product.Caption]["package_name"] = product.Caption
        # print("Vendor: " + product.Vendor)
        software_data[product.Caption]["vendor"] = product.Vendor
        # print("Version" + product.Version)
        software_data[product.Caption]["version"] = product.Version
        # print("Installation date: " + product.InstallDate)
        software_data[product.Caption]["install_date"] = product.InstallDate
        # print()

    data[os.environ["COMPUTERNAME"] + "_software_installed"] = software_data


def show_netconfigs():

    network_data = {}
    for config in cim.Win32_NetworkAdapterConfiguration():

        network_data[config.Caption] = {}
        # print("Network Adapter name: " + config.Caption)
        network_data[config.Caption]["network_adapter_name"] = config.Caption

        # print("DHCP: " + str(config.DHCPEnabled))
        network_data[config.Caption]["dhcp"] = config.DHCPEnabled

        # print("IPEnabled: " + str(config.IPEnabled))
        network_data[config.Caption]["ip_enabled"] = config.IPEnabled

        # print("Service name: " + config.ServiceName)
        network_data[config.Caption]["service_name"] = config.ServiceName
        # print()

    data[os.environ["COMPUTERNAME"] + "_network_adapter_configs"] = network_data


def show_os_info():

    os_info = {}
    for info in cim.Win32_OperatingSystem():

        os_info[info.Caption] = {}
        # print("Caption: " + info.Caption)
        os_info[info.Caption]["os_name"] = info.Caption

        #print("Build Number: " + info.BuildNumber)
        os_info[info.Caption]["build_number"] = info.Caption

        # print("Computer Name: " + info.CSName)
        os_info[info.Caption]["computer_name"] = info.CSName

        # print("Distributed: " + str(info.Distributed))
        os_info[info.Caption]["distributed"] = info.Distributed

        # print("Install Date: " + info.InstallDate)
        os_info[info.Caption]["install_date"] = info.InstallDate

        # print("Number Of Users: " + str(info.NumberOfUsers))
        os_info[info.Caption]["number_of_users"] = info.NumberOfUsers

        # print("OS Type: " + str(info.OSType))
        os_info[info.Caption]["os_type"] = info.OSType

        # print("Service Pack: " + str(info.ServicePackMajorVersion))
        os_info[info.Caption]["service_pack"] = info.ServicePackMajorVersion

        # print("Windows Directory: " + info.WindowsDirectory)
        os_info[info.Caption]["windows_directory"] = info.WindowsDirectory

        # print("System Directory: " + info.SystemDirectory)
        os_info[info.Caption]["system_directory"] = info.SystemDirectory

        # print("Version: " + info.Version)
        os_info[info.Caption]["version"] = info.Version

    data[os.environ["COMPUTERNAME"] + "_os_info"] = os_info


def show_bios_info():

    bios_info = {}
    for info in cim.Win32_BIOS():

        bios_info[info.Caption] = {}
        # print("Caption:", info.Caption)
        bios_info[info.Caption]["bios_name"] = info.Caption

        # print("Manufacturer:", info.Manufacturer)
        bios_info[info.Caption]["manufacturer"] = info.Manufacturer

        # print("BIOS Version:", info.BIOSVersion)
        bios_info[info.Caption]["bios_version"] = info.BIOSVersion

        # print("Release Date:", info.ReleaseDate)
        bios_info[info.Caption]["release_date"] = info.ReleaseDate

        # print("Serial Number:", info.SerialNumber)
        bios_info[info.Caption]["serial_number"] = info.SerialNumber

        # print("Status:", info.Status)
        bios_info[info.Caption]["status"] = info.Status

    data[os.environ["COMPUTERNAME"] + "_bios_info"] = bios_info


def show_cpu_info():

    cpu_info = {}
    for info in cim.Win32_Processor():

        cpu_info[info.Caption] = {}
        # print("Caption:", info.Caption)
        cpu_info[info.Caption]["cpu_model"] = info.Caption

        # print("Name:", info.Name)
        cpu_info[info.Caption]["cpu_name"] = info.Name

        # print("Description:", info.Description)
        cpu_info[info.Caption]["description"] = info.Description

        # print("Address Width:", info.AddressWidth)
        cpu_info[info.Caption]["address_width"] = info.AddressWidth

        # print("Manufacturer:", info.Manufacturer)
        cpu_info[info.Caption]["manufacturer"] = info.Manufacturer

        # print("Max Clock Speed: ", info.MaxClockSpeed)
        cpu_info[info.Caption]["max_clock_speed"] = info.MaxClockSpeed

        # print("Number of Cores: ", info.NumberOfCores)
        cpu_info[info.Caption]["number_of_cores"] = info.NumberOfCores

        # print("Number of Logical Processors:", info.NumberOfLogicalProcessors)
        cpu_info[info.Caption]["number_of_logical_processors"] = info.NumberOfLogicalProcessors

        # print("Revision:", info.Revision)
        cpu_info[info.Caption]["revision"] = info.Revision

        # print("Status:", info.Status)
        cpu_info[info.Caption]["status"] = info.Status

    data[os.environ["COMPUTERNAME"] + "_cpu_info"] = cpu_info


def scan_usb():
    # directory_to_upload = "X:\\"

    # print(os.path.dirname(os.path.abspath(__file__)))

    if not os.path.isdir(os.getcwd() + "/output/"):
        os.mkdir("output")

    # Export of csv, tsv & xml using usbdeview.exe
    args = ["xml", "csv", "tab"]
    for arg in args:
        if arg == "csv":
            os.system("USBDeview.exe /scomma output/" + os.environ["COMPUTERNAME"] + ".csv")
        elif arg == "xml":
            os.system("USBDeview.exe /sxml output/" + os.environ["COMPUTERNAME"] + ".xml")
        elif arg == "tab":
            os.system("USBDeview.exe /stab output/" + os.environ["COMPUTERNAME"] + ".tsv")
        else:
            print("Comando", arg, "non riconosciuto!")

##    # xml elaboration from another xml file
##    doc = etree.parse(os.environ["COMPUTERNAME"] + ".xml")
##    for element in doc.getiterator("item"):
##        if element.findtext("device_type") == "HID (Human Interface Device)":
##            element.clear()
##
##    doc.write(os.environ["COMPUTERNAME"] + "_filtered.xml", pretty_print=True)

    # getting filtered devices
    devices = []
    with open("output/" + os.environ["COMPUTERNAME"] + ".tsv", "r") as input_file:
        for line in input_file.readlines():
            if "HID (Human Interface Device)" not in line:
                line = line.split("\t")
                line = [line[1], line[2], line[3], line[4], line[5], line[6], line[9], line[10]]
                # for term in line:
                    # print(term)
                # print("----------\n\n\n")
                devices.append(line)

    # making xml file structure and writing to file
    structure = etree.Element("devices")
    output = etree.ElementTree(structure)

    for device in devices:
        item = etree.SubElement(structure, "item")

        description = etree.SubElement(item, "description")
        description.text = device[0]

        device_type = etree.SubElement(item, "device_type")
        device_type.text = device[1]

        connected = etree.SubElement(item, "connected")
        connected.text = device[2]

        safe_to_unplug = etree.SubElement(item, "safe_to_unplug")
        safe_to_unplug.text = device[3]

        disabled = etree.SubElement(item, "disabled")
        disabled.text = device[4]

        usb_hub = etree.SubElement(item, "usb_hub")
        usb_hub.text = device[5]

        created = etree.SubElement(item, "created")
        created.text = device[6]

        last_plug_unplug = etree.SubElement(item, "last_plug_unplug")
        last_plug_unplug.text = device[7]


    output.write("output/" + os.environ["COMPUTERNAME"] + "_nohid.xml")

    #os.system("cp usb/" + os.environ["COMPUTERNAME"] + ".xml " + directory_to_upload + "\\xml\\")
    #os.system("cp usb/" + os.environ["COMPUTERNAME"] + ".tsv " + directory_to_upload + "\\tsv\\")

    # print("Work complete!")


if __name__ == "__main__":
    print("\nOS Scan...", end="")
    show_os_info()
    print("Done!")

    print("USB Scan...", end="")
    scan_usb()
    print("Done!")

    print("Software Scan...", end="")
    show_software_installed()
    print("Done!")

    print("NetConfigs Scan...", end="")
    show_netconfigs()
    print("Done!")

    print("BIOS Scan...", end="")
    show_bios_info()
    print("Done!")

    print("CPU Scan...", end="")
    show_cpu_info()
    print("Done!")

    # pprint(data)
    r = json.dumps(data)

    with open('output/' + os.environ["COMPUTERNAME"] + '_dump.json', "w") as output_file:
        output_file.write(r)
