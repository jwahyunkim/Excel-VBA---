"""Compare saved workbook data and formulas with the pre-edit workbook."""
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def inspect(path):
    with ZipFile(path) as archive:
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            for item in ET.fromstring(archive.read("xl/sharedStrings.xml")):
                strings.append("".join(t.text or "" for t in item.findall(".//s:t", NS)))
        relationships = {
            item.attrib["Id"]: posixpath.normpath(posixpath.join("xl", item.attrib["Target"]))
            for item in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        }
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        result = {}
        for sheet in workbook.findall("s:sheets/s:sheet", NS):
            root = ET.fromstring(archive.read(relationships[sheet.attrib[f"{{{REL}}}id"]]))
            cells = {}
            for cell in root.findall("s:sheetData/s:row/s:c", NS):
                value = cell.find("s:v", NS)
                formula = cell.find("s:f", NS)
                kind = cell.get("t")
                text = value.text if value is not None else None
                if kind == "s":
                    text = strings[int(text)]
                elif kind == "inlineStr":
                    text = "".join(t.text or "" for t in cell.findall(".//s:t", NS))
                if formula is not None:
                    cells[cell.get("r")] = ("formula", formula.text, tuple(sorted(formula.attrib.items())))
                elif text is not None:
                    cells[cell.get("r")] = ("value", kind if kind not in ("s", "inlineStr") else "text", text)
            validations = [ET.tostring(v, encoding="unicode") for v in root.findall("s:dataValidations/s:dataValidation", NS)
                           if v.get("sqref") != "D5:D1048576"]
            protection = root.find("s:sheetProtection", NS)
            result[sheet.get("name")] = (cells, validations, None if protection is None else protection.attrib)
        names = {n.get("name"): n.text for n in workbook.findall("s:definedNames/s:definedName", NS)}
        return result, names


before, before_names = inspect(sys.argv[1])
after, after_names = inspect(sys.argv[2])
assert list(before) == list(after), "Sheet order changed"
for name in before:
    old_cells, old_validations, old_protection = before[name]
    new_cells, new_validations, new_protection = after[name]
    for address in set(old_cells) | set(new_cells):
        permitted = name == "config" and (
            address in {"L9", "M10"} or re.fullmatch(r"L(1[0-9]|2[0-9])", address)
            or re.fullmatch(r"AB([1-9]|1[0-9]|20)", address)
        )
        if not permitted:
            assert old_cells.get(address) == new_cells.get(address), f"Data/formula changed: {name}!{address}"
    assert old_validations == new_validations, f"Other validations changed: {name}"
    assert old_protection == new_protection, f"Protection changed: {name}"
for name, reference in before_names.items():
    assert after_names[name] == reference, f"Existing name changed: {name}"
assert after_names["GanttTaskTypes"].replace("'", "") == "config!$AB$1:$AB$4", "Wrong type list reference"
assert [after["config"][0][f"L{r}"][2] for r in range(10, 14)] == ["개발", "ETC", "유지보수", "프로젝트"]
print("PASS: saved defaults, sheet order, existing values/formulas, unrelated validations/names, protection")
