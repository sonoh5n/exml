import re
from typing import List, Tuple
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from dataclasses import dataclass, field


@dataclass
class ChartAttr:
    type: str
    ref: str
    sheet: str
    title: str
    dataset: field(default_factory=list)


class Chart:

    ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}

    def __init__(self, archive: ZipFile, xmlpath: str):
        try:
            self.root_tree = ET.parse(archive.open(xmlpath)).getroot()
        except KeyError:
            self.root_tree = None

        self.plot_area_elem = self.root.find(".//c:plotArea", namespaces=self.ns)
        self.chart_type_path = self.get_chart_type()

    def get_chart_type(self) -> str:

        _chart_type: str
        if self.plot_area_elem.find(".//c:scatterChart", namespaces=self.ns) is not None:
            _chart_type = "scatterChart"
        elif self.plot_area_elem.find(".//c:barChart", namespaces=self.ns) is not None:
            _chart_type = "barChart"
        elif self.plot_area_elem.find(".//c:lineChart", namespaces=self.ns) is not None:
            _chart_type = "lineChart"
        elif self.plot_area_elem.find(".//c:pieChart", namespaces=self.ns) is not None:
            _chart_type = "pieChart"
        else:
            # Not Support type: doughnutChart/bubbleChart/radarChart/surfaceChart
            _chart_type = ""

        return f".//c:{_chart_type}"

    def get_chart_range(self) -> Tuple[str, str]:
        range_pattern = f".//{self.chart_type_path}/c:ser/c:xVal/c:numRef/c:f"
        range = self.plot_area_elem.find(range_pattern, namespaces=self.ns)
        if not range:
            return "", ""

        match = re.match(r'^(.+!)(\$[A-Z]+\$\d+:\$[A-Z]+\$\d+)$', range)
        if match:
            sheet_name, cell_range = match.group()
        else:
            return "", ""

        return sheet_name, cell_range


    def get_chart():
        ...








class XmlChart:

    chart_xml = r'^xl/charts/chart.*\.xml$'


    def __init__(self, archive: ZipFile):
        self.archive = archive
        self.chartfiles: List[str] = self.__chartfiles()

        # try:
        #     self.root_tree = ET.parse(self.archive.open(self.sharedstyle_xml)).getroot()
        # except KeyError:
        #     self.root_tree = None
        # self.si = self.__get_shareitem()

    def __chartfiles(self) -> List:
        return [f for f in self.archive.namelist() if re.match(self.chart_xml, f)]


    def __chartfile(self, index: int) -> str:
        try:
            return self.chartfiles[index]
        except IndexError:
            return ""


    def get_chart(self, index: int):
        chartpath = self.__chartfile(index)
        if not chartpath:
            return

        index = str(index)

