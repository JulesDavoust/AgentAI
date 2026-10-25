"""Fixtures des tests de l'outil vba2py (aucun appel à l'IA)."""

from pathlib import Path

import pytest
import xlsxwriter

DATA = Path(__file__).parent / "data"
# vbaProject.bin provient des exemples de XlsxWriter (licence BSD) :
# un Module1 contenant `say_hello` et des modules de documents vides.
VBA_BIN = DATA / "vbaProject.bin"

MODULE1_CODE = 'Sub say_hello()\r\n    MsgBox ("Hello from Python!")\r\nEnd Sub\r\n'


def make_xlsm(path: Path) -> Path:
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Ventes")
    ws.write_row(0, 0, ["Date", "Client", "Montant"])
    ws.write_row(1, 0, ["2024-01-02", "Client confidentiel", 1234.5])
    ws.write_row(2, 0, ["2024-01-03", "Autre client", 99])
    ws.write_formula(4, 2, "=SUM(C2:C3)")
    wb.define_name("Total", "=Ventes!$C$5")
    wb.add_vba_project(str(VBA_BIN))
    wb.close()
    return path


def make_export_dir(directory: Path, module1_code: str = MODULE1_CODE) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    (directory / "Module1.bas").write_bytes(
        ('Attribute VB_Name = "Module1"\r\n' + module1_code).encode("cp1252")
    )
    (directory / "ThisWorkbook.cls").write_bytes(
        b'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n'
        b'Attribute VB_Name = "ThisWorkbook"\r\nAttribute VB_PredeclaredId = True\r\n'
    )
    return directory


@pytest.fixture
def xlsm(tmp_path) -> Path:
    return make_xlsm(tmp_path / "Ventes 2024.xlsm")


@pytest.fixture
def export_dir(tmp_path) -> Path:
    return make_export_dir(tmp_path / "Ventes_vba")
