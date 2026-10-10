"""실제 Chromium + 네트워크 차단 fixture. 홈택스 실사이트를 검증하는 테스트는 아닙니다."""
import base64
from dataclasses import replace
import io
import json
from pathlib import Path
from types import SimpleNamespace
import pytest
from playwright.sync_api import sync_playwright
from pypdf import PdfWriter
from vat_core import Client,Job,Period,validate_artifact
from hometax_vat_downloader import Collector,hometax_host

@pytest.fixture
def runtime(tmp_path):
    with sync_playwright() as pw:
        browser=pw.chromium.launch(executable_path='/usr/bin/chromium',headless=True,args=['--no-sandbox'])
        context=browser.new_context(accept_downloads=True)
        page=context.new_page()
        args=SimpleNamespace(selectors=None,timeout=3,generation_timeout=3,mode='strict')
        incoming=tmp_path/'incoming';incoming.mkdir()
        collector=Collector(context,page,args,incoming,True)
        yield collector,context,page,tmp_path
        browser.close()


def job(service=1,filters=None):
    c=Client(2,'테스트업체','1234567890',Period.parse('2기예정',2026),tuple(range(1,8)),
             ('전자세금계산서','전자계산서'),('매출','매입'),('내국신용장','구매확인서'),
             ('거래내역조회',),'전체','전체',('PDF','원본'))
    return Job(c,service,'test',filters or {'사업자등록번호':c.biz,'과세연도':'2026','기수':'2기','신고구분':'예정'})


def serve(page,html):
    page.route('**/*',lambda route:route.fulfill(status=200,content_type='text/html; charset=utf-8',body=html))
    page.goto('https://www.hometax.go.kr/local-test')


def base_form():
    return '''<label for="biz">사업자등록번호</label><input id="biz" title="사업자등록번호">
    <table><tr><th>과세기간(년/기)</th><td><select id="year"><option>2025년</option><option selected>2026년</option></select>
    <select id="term"><option>1기</option><option>2기</option></select></td></tr></table>
    <input type="radio" id="early" name="filing"><label for="early">예정</label>
    <input type="radio" id="final" name="filing"><label for="final">확정</label>
    <button onclick="setTimeout(()=>{document.getElementById('result').innerText=document.getElementById('biz').value+' 2026 2기 예정 결과';document.getElementById('result').hidden=false},250)">조회</button>
    <div id="result" hidden></div>'''


def test_native_filters_async_readiness_and_result_identity(runtime):
    c,ctx,page,tmp=runtime;serve(page,base_form())
    c.config={'services':{'1':{'result_ready':{'css':'#result'},'result_identity':{'css':'#result'}}}}
    assert c.configure_query(job()) is True
    assert page.locator('#biz').input_value()=='1234567890'
    assert page.locator('#term').input_value()=='2기'
    assert page.locator('#early').is_checked()
    assert '1234567890' in page.locator('#result').inner_text()


def test_wrong_result_client_is_not_success(runtime):
    c,ctx,page,tmp=runtime;serve(page,base_form().replace("document.getElementById('biz').value+' 2026 2기 예정 결과'","'9999999999 2026 2기 예정 결과'"))
    c.config={'services':{'1':{'result_ready':{'css':'#result'},'result_identity':{'css':'#result'}}}}
    with pytest.raises(RuntimeError,match='사업자번호'):c.configure_query(job())


def test_split_business_number(runtime):
    c,ctx,page,tmp=runtime
    serve(page,'<table><tr><th>사업자번호</th><td><input maxlength="3"><input maxlength="2"><input maxlength="5"></td></tr></table>')
    c.set_field(job(4),'사업자등록번호','1234567890')
    assert [e.input_value() for e in page.locator('input').all()]==['123','45','67890']


def test_duplicate_control_does_not_guess(runtime):
    c,ctx,page,tmp=runtime;serve(page,'<button>조회</button><button>조회</button>')
    with pytest.raises(RuntimeError,match='여러 개'):c.click(('조회',))


def test_ambiguous_select_does_not_choose_first(runtime):
    c,ctx,page,tmp=runtime;serve(page,'<label for="a">조회월</label><select id="a"><option>7월</option><option>7월</option></select>')
    with pytest.raises(RuntimeError,match='정확히'):c.set_field(job(4),'조회월','7월')


def test_nested_same_origin_frames(runtime):
    c,ctx,page,tmp=runtime
    inner='<label for="b">사업자등록번호</label><input id="b">'
    middle='<iframe srcdoc="'+inner.replace('"','&quot;')+'"></iframe>'
    serve(page,'<iframe srcdoc="'+middle.replace('&','&amp;').replace('"','&quot;')+'"></iframe>')
    page.wait_for_timeout(200)
    c.set_field(job(),'사업자등록번호','1234567890')
    assert any(f.locator('#b').count() and f.locator('#b').input_value()=='1234567890' for f in page.frames)


def test_download_event_collects_real_file_and_preserves_old(runtime):
    c,ctx,page,tmp=runtime
    text='column,value\nfoo,123\n'
    serve(page,"<a download='native.csv' href='data:text/csv,"+text.replace('\n','%0A')+"'>엑셀 내려받기</a>")
    target=tmp/'out';target.mkdir()
    old=tmp/'old.csv';old.write_text('preserve')
    a=c.download(job(5),target,('엑셀 내려받기',))
    assert Path(a['path']).read_text()==text and old.read_text()=='preserve'
    assert a['source']=='native_download'


def test_download_of_html_login_is_rejected(runtime):
    c,ctx,page,tmp=runtime
    serve(page,"<a download='report.xls' href='data:application/octet-stream,%3Chtml%3Elogin%3C/html%3E'>다운로드</a>")
    target=tmp/'out';target.mkdir()
    with pytest.raises(ValueError,match='HTML'):c.download(job(6),target,('다운로드',))


def test_file_generation_waits_for_new_completion(runtime):
    c,ctx,page,tmp=runtime
    serve(page,"""<button onclick="document.getElementById('dl').disabled=true;setTimeout(()=>{document.getElementById('dl').disabled=false;alert('파일생성이 완료되었습니다.')},300)">파일생성</button>
    <button id='dl' onclick="let a=document.createElement('a');a.href='data:text/csv,head%0Avalue';a.download='result.csv';a.click()">다운로드</button>""")
    target=tmp/'out';target.mkdir()
    result=c.generate_download(job(6),target)
    assert Path(result['path']).read_text()=='head\nvalue'


def test_stale_enabled_download_not_used(runtime):
    c,ctx,page,tmp=runtime;c.args.generation_timeout=1
    serve(page,"<button>파일생성</button><button>다운로드</button>")
    target=tmp/'out';target.mkdir()
    with pytest.raises(RuntimeError,match='시간초과'):c.generate_download(job(6),target)
    assert not list(target.iterdir())


def test_screen_pdf_is_labeled_as_evidence(runtime):
    c,ctx,page,tmp=runtime;serve(page,'<h1>조회자료</h1><p>총 0건</p>')
    target=tmp/'out';target.mkdir()
    a=c.screen_pdf(job(5),target,True)
    assert a['pages']>=1 and a['source']=='no_data_screen'
    assert '조회화면증빙' in Path(a['path']).name
    assert c.no_data(job(5)) and not c.no_data(job(3))


def test_new_report_file_only_not_stale_pdf(runtime):
    c,ctx,page,tmp=runtime
    old=c.incoming/'old.pdf'
    w=PdfWriter();w.add_blank_page(width=100,height=100)
    with old.open('wb') as f:w.write(f)
    serve(page,'<button>dummy</button>')
    target=tmp/'out';target.mkdir();c.args.timeout=1
    with pytest.raises(RuntimeError,match='시간초과'):c.wait_print_file(job(),target,{old})
    assert old.is_file() and not list(target.iterdir())


def test_hometax_domain_check():
    assert hometax_host('https://www.hometax.go.kr')
    assert hometax_host('https://teht.hometax.go.kr')
    assert not hometax_host('https://www.hometax.go.kr.example.com')
    assert not hometax_host('https://hometax.go.kr@evil.example')


def test_same_named_button_across_frames_is_ambiguous(runtime):
    c,ctx,page,tmp=runtime
    serve(page,'<button>조회</button><iframe srcdoc="&lt;button&gt;조회&lt;/button&gt;"></iframe>')
    page.wait_for_timeout(100)
    with pytest.raises(RuntimeError,match='여러 개'):c.click(('조회',))


def test_wrong_result_period_is_not_success(runtime):
    c,ctx,page,tmp=runtime;serve(page,base_form().replace('2026 2기 예정 결과','2026 1기 예정 결과'))
    c.config={'services':{'1':{'result_ready':{'css':'#result'},'result_identity':{'css':'#result'}}}}
    with pytest.raises(RuntimeError,match='기수'):c.configure_query(job())


def test_no_data_selector_only_applies_when_visible(runtime):
    c,ctx,page,tmp=runtime
    c.config={'services':{'5':{'no_data':{'css':'#empty'}}}}
    serve(page,'<div id="empty" hidden>조회된 결과가 없습니다.</div><p>총 1건</p>')
    assert not c.no_data(job(5))
    page.locator('#empty').evaluate('e=>e.hidden=false')
    assert c.no_data(job(5))
