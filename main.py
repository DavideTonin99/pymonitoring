import os
import sys
from lxml import etree


directory_to_upload = "X:\\"


if __name__ == "__main__":
    args = sys.argv[1:]

    # Export of csv, tsv & xml using usbdeview.exe
    if args == []:
        args = ["xml", "csv", "tab"]
        for arg in args:
            if arg == "csv":
                os.system("USBDeview.exe /scomma " + os.environ["COMPUTERNAME"] + ".csv")
            elif arg == "xml":
                os.system("USBDeview.exe /sxml " + os.environ["COMPUTERNAME"] + ".xml")
            elif arg == "tab":
                os.system("USBDeview.exe /stab " + os.environ["COMPUTERNAME"] + ".tsv")
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
    with open(os.environ["COMPUTERNAME"] + ".tsv") as input_file:
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


    output.write(os.environ["COMPUTERNAME"] + "_nohid.xml")

    os.system("cp " + os.environ["COMPUTERNAME"] + ".xml " + directory_to_upload + "\\xml\\")
    os.system("cp " + os.environ["COMPUTERNAME"] + ".tsv " + directory_to_upload + "\\tsv\\")

    print("Work complete!")
