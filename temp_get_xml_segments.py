import xml.etree.ElementTree as ET

tree = ET.parse("d:/WORK/OneDrive/Dokumenty/apps_data/claude/code/idoc-to-excel-repo/test_idoc.xml")
root = tree.getroot()

idoc = root.find('IDOC')
if idoc:
    segments = set()
    for elem in idoc:
        segments.add(elem.tag)
    print(f"XML segments needed: {sorted(segments)}")
