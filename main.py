import os
from lxml import etree
import wmi


cim = wmi.WMI()


def show_software_installed():
    for product in cim.Win32_Product():
        print("Package name: " + product.Caption)
        print("Vendor: " + product.Vendor)
        print("Version" + product.Version)
        print("Installation date: " + product.InstallDate)
        print()


def show_netconfigs():
    for config in cim.Win32_NetworkAdapterConfiguration():
        print("Network Adapter name: " + config.caption)
        print("DHCP: " + str(config.DHCPEnabled))
        print("IPEnabled: " + str(config.IPEnabled))
        print("Service name: " + config.ServiceName)
        print()


def show_os_info():
    for info in cim.Win32_OperatingSystem():
        print("Build Number: " + info.BuildNumber)
        print("Caption: " + info.Caption)
        print("Computer Name: " + info.CSName)
        print("Distributed: " + str(info.Distributed))
        print("Install Date: " + info.InstallDate)
        print("Number Of Users: " + str(info.NumberOfUsers))
        print("OS Type: " + str(info.OSType))
        print("Service Pack: " + str(info.ServicePackMajorVersion))
        print("Windows Directory: " + info.WindowsDirectory)
        print("System Directory: " + info.SystemDirectory)
        print("Version: " + info.Version)


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
    print("OS Scan...", end="")
    show_os_info()
    print(" Done!\n")

    print("USB Scan...", end="")
    scan_usb()
    print(" Done!\n")

    print("Software Scan...\n")
    show_software_installed()
    print("\n...Done!\n")

    print("NetConfigs Scan...\n")
    show_netconfigs()
    print("\n...Done!\n")
