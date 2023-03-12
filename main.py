import xml.etree.ElementTree as ET

from exmlrd.archive import ExcelArchive

filepath = "tests/samle-dataset.xlsx"
e = ExcelArchive(filepath)

xml_path = "tests/dataset/samle-dataset/xl/charts/chart1.xml"
ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
tree = ET.parse(xml_path)

root = tree.getroot()

plot_area_elem = root.find(".//c:plotArea", namespaces=ns)

if plot_area_elem.find(".//c:scatterChart", namespaces=ns) is not None:
    print("散布図")
elif plot_area_elem.find(".//c:barChart", namespaces=ns) is not None:
    print("棒グラフ/横棒グラフ")
else:
    print("不明")

# start label
strlabel = plot_area_elem.find(".//c:scatterChart/c:ser/c:tx/c:strRef/c:f", namespaces=ns)

# range
range = plot_area_elem.find(".//c:scatterChart/c:ser/c:xVal/c:numRef/c:f", namespaces=ns)

# formatcode
formatcode = plot_area_elem.find(".//c:scatterChart/c:ser/c:xVal/c:numRef/c:numCache/c:f", namespaces=ns)

# xVal
xValnum = plot_area_elem.find(".//c:scatterChart/c:ser/c:xVal/c:numRef/c:numCache", namespaces=ns)

for v in xValnum:
    print(v.attrib, v.tag, v.text)

