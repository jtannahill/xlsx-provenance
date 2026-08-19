"""Tests for metadata editing and the declared-vs-structural verdict rule.

Fixtures are built with zipfile only, so the suite stays stdlib-only."""

import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import xlsx_provenance as xp  # noqa: E402

CT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    "</Types>"
)
RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    "</Relationships>"
)
WORKBOOK_LIB = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets>'
    "</workbook>"
)
CORE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    "<dc:creator>Lib Bot</dc:creator><cp:lastModifiedBy>Lib Bot</cp:lastModifiedBy>"
    '<dcterms:created xsi:type="dcterms:W3CDTF">2024-05-01T10:00:00Z</dcterms:created>'
    '<dcterms:modified xsi:type="dcterms:W3CDTF">2024-05-01T10:00:00Z</dcterms:modified>'
    "</cp:coreProperties>"
)
APP_LIB = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    "<Application>openpyxl</Application><Company>Acme</Company></Properties>"
)


def make_xlsx(path, app_xml=APP_LIB, core_xml=CORE, extra=None):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("docProps/app.xml", app_xml)
        z.writestr("docProps/core.xml", core_xml)
        z.writestr("xl/workbook.xml", WORKBOOK_LIB)
        z.writestr("xl/worksheets/sheet1.xml", "<worksheet/>")
        for name, data in (extra or {}).items():
            z.writestr(name, data)
    return path


def members(path):
    with zipfile.ZipFile(path) as z:
        return [(i.filename, i.compress_type) for i in z.infolist()]


@pytest.fixture
def lib_file(tmp_path):
    return make_xlsx(tmp_path / "lib.xlsx")


def test_strip_blanks_identity_and_keeps_structure(lib_file, tmp_path):
    before = members(lib_file)
    out = tmp_path / "out.xlsx"
    written = xp.rewrite_metadata(lib_file, out, strip=True)
    assert written == {}
    assert members(out) == before
    fp = xp.fingerprint(out)
    assert fp.creator is None and fp.application is None and fp.company is None
    assert fp.metadata_stripped is True
    assert "metadata=stripped" in fp.signals


def test_strip_in_place_atomic(lib_file):
    xp.rewrite_metadata(lib_file, lib_file, strip=True)
    assert not list(lib_file.parent.glob("*.tmp-xlsxprov"))
    assert xp.fingerprint(lib_file).metadata_stripped


def test_set_properties_and_dates(lib_file, tmp_path):
    out = tmp_path / "out.xlsx"
    written = xp.rewrite_metadata(
        lib_file, out,
        updates={"creator": "Jane Doe", "company": "Globex", "created": "2024-01-02",
                 "title": "Q1", "manager": ""},
    )
    assert written["creator"] == "Jane Doe"
    assert written["company"] == "Globex"
    assert written["created"] == "2024-01-02T00:00:00Z"
    assert written["title"] == "Q1"
    assert "manager" not in written
    fp = xp.fingerprint(out)
    assert fp.creator == "Jane Doe" and fp.company == "Globex"
    # untouched properties survive
    assert fp.last_modified_by == "Lib Bot"


def test_set_date_carries_w3cdtf_type(lib_file, tmp_path):
    out = tmp_path / "out.xlsx"
    xp.rewrite_metadata(lib_file, out, updates={"modified": "2025-06-07T08:09:10+02:00"})
    with zipfile.ZipFile(out) as z:
        root = ET.fromstring(z.read("docProps/core.xml"))
    elt = root.find("{http://purl.org/dc/terms/}modified")
    assert elt.text == "2025-06-07T06:09:10Z"
    assert elt.get("{http://www.w3.org/2001/XMLSchema-instance}type") == "dcterms:W3CDTF"


def test_unknown_property_rejected(lib_file, tmp_path):
    with pytest.raises(KeyError):
        xp.rewrite_metadata(lib_file, tmp_path / "o.xlsx", updates={"bogus": "1"})


def test_bad_date_rejected(lib_file, tmp_path):
    with pytest.raises(ValueError):
        xp.rewrite_metadata(lib_file, tmp_path / "o.xlsx", updates={"created": "yesterday"})


def test_declared_excel_without_structure_is_suspect(lib_file, tmp_path):
    """Editing Application to claim Excel must not flip a library file's verdict."""
    assert xp.fingerprint(lib_file).verdict == "OPENPYXL"
    out = tmp_path / "forged.xlsx"
    xp.rewrite_metadata(lib_file, out, updates={"application": "Microsoft Excel"})
    fp = xp.fingerprint(out)
    assert fp.verdict == "SUSPECT"
    assert "lacks Excel structure" in fp.verdict_reason


def test_cli_strip_and_set(lib_file, tmp_path):
    out = tmp_path / "cli.xlsx"
    r = subprocess.run(
        [sys.executable, str(ROOT / "xlsx_provenance.py"), "--strip",
         "--set", "creator=CLI User", "-o", str(out), "--no-color", str(lib_file)],
        capture_output=True, text=True,
    )
    assert r.returncode == 0, r.stderr
    assert "updated" in r.stdout and "CLI User" in r.stdout
    fp = xp.fingerprint(out)
    assert fp.creator == "CLI User" and fp.application is None


def test_cli_list_properties():
    r = subprocess.run(
        [sys.executable, str(ROOT / "xlsx_provenance.py"), "--list-properties"],
        capture_output=True, text=True,
    )
    assert r.returncode == 0
    assert "creator" in r.stdout and "application" in r.stdout
