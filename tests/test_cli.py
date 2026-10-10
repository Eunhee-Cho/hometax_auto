from openpyxl import Workbook
from hometax_vat_downloader import main


def test_service_selection_excludes_unsupported_trade_scope_before_planning(tmp_path,capsys):
    p=tmp_path/'a.xlsx';w=Workbook();s=w.active
    s.append(['사업자명','사업자등록번호','과세기간','과세연도','조회범위'])
    s.append(['테스트','1234567890','2기확정',2026,'분기']);w.save(p)
    assert main(['--input',str(p),'--services','1','--validate-only'])==0
    assert '세부 작업 1건' in capsys.readouterr().out
    assert main(['--input',str(p),'--services','7','--validate-only'])==2
    assert '7번' in capsys.readouterr().err
    assert main(['--input',str(p),'--services','8','--validate-only'])==2


def test_one_file_runs_alone_with_original_excel_and_2026_fallback(tmp_path):
    import os
    from pathlib import Path
    import shutil
    import subprocess
    import sys
    import hometax_vat_downloader
    script=tmp_path/'hometax_vat_downloader.py'
    shutil.copy2(hometax_vat_downloader.__file__,script)
    w=Workbook();s=w.active
    s.append(['사업자명','사업자등록번호','과세기간'])
    s.append(['테스트','1234567890','2기예정']);w.save(tmp_path/'INPUT-부가세.xlsx')
    env=os.environ.copy();env.pop('PYTHONPATH',None)
    r=subprocess.run([sys.executable,str(script),'--validate-only'],cwd=tmp_path,env=env,capture_output=True,text=True,timeout=30)
    assert r.returncode==0,r.stderr
    assert '세부 작업 13건' in r.stdout and '2026_2기예정' in r.stdout and '모드: auto' in r.stdout
    assert set(p.name for p in tmp_path.iterdir() if p.is_file())=={'hometax_vat_downloader.py','INPUT-부가세.xlsx'}


def test_input_discovery_does_not_choose_dummy_template(tmp_path,monkeypatch):
    import pytest
    import hometax_vat_downloader as module
    monkeypatch.setattr(module,'ROOT',tmp_path)
    (tmp_path/'INPUT-부가세_빈양식.xlsx').touch()
    with pytest.raises(ValueError,match='INPUT-부가세.xlsx'):module.discover_input()
    actual=tmp_path/'INPUT-부가세.xlsx';actual.touch()
    assert module.discover_input()==actual
    expanded=tmp_path/'INPUT-부가세_확장.xlsx';expanded.touch()
    assert module.discover_input()==expanded
