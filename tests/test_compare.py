from conftest import MODULE1_CODE, make_export_dir
from vba2py.compare import compare, format_report


def test_export_identique_a_l_extraction(xlsm, export_dir):
    results = compare(export_dir, xlsm)
    assert [(r.name, r.status) for r in results] == [("Module1", "identique")]
    assert "1/1 module(s) identique(s)" in format_report(results)


def test_difference_detectee(xlsm, tmp_path):
    export = make_export_dir(tmp_path / "modifie", MODULE1_CODE.replace("Hello", "Bonjour"))
    [result] = compare(export, xlsm)
    assert result.status == "different"
    assert '-    MsgBox ("Bonjour from Python!")' in result.diff


def test_module_absent(xlsm, export_dir):
    (export_dir / "Module2.bas").write_text('Attribute VB_Name = "Module2"\nSub B()\nEnd Sub\n')
    statuses = {r.name: r.status for r in compare(export_dir, xlsm)}
    assert statuses == {"Module1": "identique", "Module2": "absent_classeur"}
