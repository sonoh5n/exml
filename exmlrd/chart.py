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

        self.plot_area_elem = self.root_tree.find(".//c:plotArea", namespaces=self.ns)
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

        return _chart_type

    def get_chart_range(self, value_type: str) -> Tuple[str, str]:
        range_pattern = f".//c:{self.chart_type_path}/c:ser/c:{value_type}/c:numRef/c:f"
        range = self.plot_area_elem.find(range_pattern, namespaces=self.ns)
        if not range:
            return "", ""

        match = re.match(r'^(.+!)(\$[A-Z]+\$\d+:\$[A-Z]+\$\d+)$', range)
        if match:
            sheet_name, cell_range = match.group()
        else:
            return "", ""

        return sheet_name, cell_range


    def get_chart(self):
        xval_pattern = f".//c:{self.chart_type_path}/c:ser/c:xVal/c:numRef/c:numCache"
        xvals = []
        for elem in self.plot_area_elem.find(xval_pattern, namespaces=self.ns):
            if not elem.attrib.get("idx"):
                continue
            for v in elem:
                xvals.append(convert_numeric_string(v.text))

        yval_pattern = f".//c:{self.chart_type_path}/c:ser/c:yVal/c:numRef/c:numCache"
        yvals = []
        for elem in self.plot_area_elem.find(yval_pattern, namespaces=self.ns):
            if not elem.attrib.get("idx"):
                continue
            for v in elem:
                yvals.append(convert_numeric_string(v.text))

        return xvals, yvals


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
        chart = Chart(self.archive, chartpath)

        x, y = chart.get_chart()
        return chart


def convert_numeric_string(sval: str):
    try:
        return int(sval)
    except ValueError:
        try:
            return float(sval)
        except ValueError:
            return None
