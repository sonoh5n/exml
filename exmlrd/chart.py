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
    x_dataset: field(default_factory=list)
    y_dataset: field(default_factory=list)


@dataclass
class RangeRef:
    sheet: str = ""
    range: str = ""


class Chart:

    ns = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    }

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

    def get_chart_range(self, chart_type: str) -> List[RangeRef]:

        refs = []
        numRef_tags = self.plot_area_elem.findall('.//c:numRef', namespaces=self.ns)
        for numRef in numRef_tags:
            f = numRef.find('.//c:f', namespaces=self.ns)
            sheet, range = self.extract_sheetref_range(f.text)
            ref = RangeRef(sheet=sheet, range=range)
            refs.append(ref)

        if chart_type == "pieChart":
            strRef_tag = self.plot_area_elem.find('.//c:strRef', namespaces=self.ns)
            f = strRef_tag.find('.//c:f', namespaces=self.ns)
            sheet, range = self.extract_sheetref_range(f.text)
            ref = RangeRef(sheet=sheet, range=range)
            refs.append(ref)

        return refs

    def extract_sheetref_range(self, input_pattern: str):
        scope = r'^(.+!)(\$[A-Z]+\$\d+)(:\$[A-Z]+\$\d+)?$'
        match = re.match(scope, input_pattern)
        if match:
            sheet_name = match.group(1)
            start_cell = match.group(2)
            end_cell = match.group(3)
            if sheet_name:
                sheet_name = sheet_name.replace("!", "")
            if end_cell:
                cell_range = start_cell + end_cell
            else:
                cell_range = start_cell
        else:
            return "", ""

        return sheet_name, cell_range

    def get_title(self):
        title = ""
        title_elem_tree = self.root_tree.find(".//c:title", namespaces=self.ns)
        text_elem = title_elem_tree.findall(".//a:t", namespaces=self.ns)
        for e in text_elem:
            title = title + e

    def get_values(self):
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

        chart = Chart(self.archive, chartpath)
        chart_type = chart.get_chart_type()
        refs = chart.get_chart_range(chart_type)
        title = chart.get_title()
        x, y = chart.get_values()
        attr = ChartAttr(
            type=chart.get_chart_type(),
            ref=range,
            sheet=st
        )
        return chart


def convert_numeric_string(sval: str):
    try:
        return int(sval)
    except ValueError:
        try:
            return float(sval)
        except ValueError:
            return None
