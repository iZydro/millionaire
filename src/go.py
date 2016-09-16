import xml.etree.ElementTree as ET
import collections
import xlwt
# import json

names = [
            "itemDefinitions", "commerceDefinitions", "decorationDefinitions", "wonderDefinitions",
            "contracts", "contractsTypes",
            "dailyRewardsDefinitions",
            "missionDefinitions",
            "XPTable"
         ]

workbook = xlwt.Workbook()

for name in names:

    xmlfile = 'mc/' + name + '.xml'
    h = open(xmlfile, "r", encoding="iso-8859-1")
    xmlstring = h.read()

    root = ET.fromstring(xmlstring)

    headers = []

    elements = collections.OrderedDict()

    for element in root:

        idname = ""
        if "sku" in element.attrib:
            idname = "sku"
        if "id" in element.attrib:
            idname = "id"

        tid = element.attrib[idname]
        elements[tid] = {}
        for attrib in element.attrib:
            if attrib not in headers:
                headers.append(attrib)
            elements[tid][attrib] = element.attrib[attrib]

    # print(json.dumps(elements, indent=4))

    if "sku" in headers:
        headers.remove("sku")
        headers.insert(0, "sku")
    if "id" in headers:
        headers.remove("id")
        headers.insert(0, "id")

    sheet = workbook.add_sheet(name)

    row = 0
    col = 0
    for header in headers:
        sheet.write(row, col, header)
        col += 1

    row += 1
    for element in elements:
        col = 0
        for header in headers:
            if header in elements[element]:
                sheet.write(row, col, elements[element][header])
            col += 1
        row += 1

workbook.save('mc.xls')
