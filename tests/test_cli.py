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
