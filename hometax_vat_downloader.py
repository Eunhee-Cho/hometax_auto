"""1~7 홈택스 자료 수집. 로그인/권한은 사용자가 완료하며 신고·납부하지 않습니다.

기본 assisted 모드: 화면에 맞지 않는 선택자는 수동 확인으로 보완합니다.
strict 모드: 선택자/대상 검증이 안 되면 실패로 기록합니다.
--validate-only 는 브라우저를 열지 않습니다.
"""
from __future__ import annotations
import argparse
from dataclasses import replace
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.request
from pathlib import Path
from urllib.parse import urlparse
from vat_core import SERVICES, Job, Journal, build_jobs, load_clients, validate_artifact

ROOT = Path(__file__).resolve().parent
URL = 'https://www.hometax.go.kr'
MENUS = {
  1: ('부가가치세 신고도움 서비스','부가가치세 신고도움서비스'),
  2: ('전자(세금)계산서 합계표 조회','전자(세금)계산서 합계표조회'),
  3: ('신용카드/판매(결제)대행 매출자료 조회','신용카드 판매(결제)대행 매출자료 조회'),
  4: ('현금영수증 매출총액 조회',),
  5: ('부가가치세 매입자납부 특례 조회','부가가치세 매입자납부특례 조회'),
  6: ('수출실적명세서 조회',),
  7: ('내국신용장·구매확인서 전자발급명세서조회','내국신용장·구매확인서 전자발급명세서 조회'),
}
FIELD_ALIASES = {
 '사업자등록번호': ('사업자등록번호','사업자번호'),
 '과세연도': ('과세기간(년/기)','과세기간(년)','조회년도','조회연도','귀속연도','결제년도','결제연도','거래년도','선적년월','구매공급일','분기별','일자별'),
 '기수': ('과세기간(년/기)','과세기간','기수','구매공급일'),
 '조회월': ('조회년도','조회연도','월'),
 '시작분기': ('분기','결제년도'), '종료분기': ('분기','결제년도'),
 '시작일': ('조회시작일자','시작일자','시작일','일자별'),
 '종료일': ('조회종료일자','종료일자','종료일','일자별'),
 '조회기간': ('선적년월','조회기간'), '적용분기': ('적용분기',),
 '품목구분': ('품목 구분','품목구분'), '매입매출구분': ('매입/매출 구분','매입매출구분'),
}
RADIO_FIELDS = {'신고구분','분류','구분','작성일자','조회단위','서류구분'}


def norm(text):
    return re.sub(r'\s+', '', str(text))


def hometax_host(url):
    host = urlparse(url).hostname or ''
    return host == 'hometax.go.kr' or host.endswith('.hometax.go.kr')


def prompt(message):
    print('\n'+message)
    answer = input('완료 후 Enter / 건너뛰기 s / 종료 q: ').strip().lower()
    while answer not in ('','s','q'):
        answer = input('Enter(완료), s(건너뛰기), q(종료) 중 입력하세요: ').strip().lower()
    if answer == 'q':
        raise KeyboardInterrupt
    if answer == 's':
        raise RuntimeError('사용자가 이 작업을 건너뛰었습니다.')


def chrome_binary(explicit):
    if explicit:
        p = Path(explicit)
        if p.is_file(): return str(p)
        raise ValueError('--chrome 실행파일을 찾지 못했습니다.')
    candidates = [
        Path(os.environ.get('PROGRAMFILES','C:/Program Files'))/'Google/Chrome/Application/chrome.exe',
        Path(os.environ.get('PROGRAMFILES(X86)','C:/Program Files (x86)'))/'Google/Chrome/Application/chrome.exe',
        Path(os.environ.get('LOCALAPPDATA','C:/Users/Default/AppData/Local'))/'Google/Chrome/Application/chrome.exe',
        Path('/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'),
    ]
    for p in candidates:
        if p.is_file(): return str(p)
    for name in ('google-chrome','chromium','chromium-browser'):
        if shutil.which(name): return shutil.which(name)
    raise ValueError('Chrome이 필요합니다. 설치하거나 --chrome으로 실행파일 경로를 지정하세요.')


def debug_info(port):
    # localhost 요청에 조직 HTTP 프록시를 적용하지 않습니다.
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    try:
        with opener.open(f'http://127.0.0.1:{port}/json/version',timeout=1) as r:
            info = json.load(r)
        return info if 'webSocketDebuggerUrl' in info else None
    except (OSError,ValueError): return None


def launch_chrome(args, incoming):
    if debug_info(args.port):
        # 이 실행으로 시작한 Chrome의 print 설정인지 알 수 없어 무인 인쇄하지 않습니다.
        return False
    profile = args.profile.resolve()
    profile.mkdir(parents=True,exist_ok=True)
    prefs = profile/'Default'/'Preferences'
    prefs.parent.mkdir(exist_ok=True)
    data = json.loads(prefs.read_text(encoding='utf-8')) if prefs.exists() else {}
    data.setdefault('download',{}).update({'default_directory':str(incoming),'prompt_for_download':False})
    data.setdefault('savefile',{})['default_directory'] = str(incoming)
    sticky = {'recentDestinations':[{'id':'Save as PDF','origin':'local','account':''}],
              'selectedDestinationId':'Save as PDF','version':2,'isHeaderFooterEnabled':False}
    data.setdefault('printing',{}).setdefault('print_preview_sticky_settings',{})['appState']=json.dumps(sticky)
    tmp = prefs.with_suffix('.new')
    tmp.write_text(json.dumps(data),encoding='utf-8'); tmp.replace(prefs)
    subprocess.Popen([chrome_binary(args.chrome),f'--remote-debugging-port={args.port}',
       '--remote-debugging-address=127.0.0.1',f'--user-data-dir={profile}',
       '--no-first-run','--no-default-browser-check','--start-maximized','--kiosk-printing',URL],
       stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    deadline=time.monotonic()+30
    while time.monotonic()<deadline:
        if debug_info(args.port): return True
        time.sleep(.3)
    raise RuntimeError('Chrome 디버그 연결 시간초과. 전용 프로필을 사용하는 다른 Chrome을 닫으세요.')


class Collector:
    def __init__(self, context, page, args, incoming, kiosk=False):
        self.context,self.page,self.args,self.incoming,self.kiosk = context,page,args,incoming,kiosk
        self.config = json.loads(args.selectors.read_text(encoding='utf-8')) if args.selectors else {}
        self.dialogs=[]
        self.owned=[]
        self.context.on('page',self._new_page)
        for p in context.pages: self._watch(p)

    def _watch(self,page):
        def dialog(d):
            # 조회/파일생성 후 안내창. 원문에는 개인정보가 있을 수 있어 로그에 남기지 않습니다.
            self.dialogs.append(d.message)
            d.accept()
        page.on('dialog',dialog)
        page.set_default_timeout(self.args.timeout*1000)

    def _new_page(self,page):
        self._watch(page)
        self.owned.append(page)

    def frames(self,page=None):
        p=page or self.page
        return [f for f in p.frames if f==p.main_frame or f.url in ('','about:blank','about:srcdoc') or hometax_host(f.url)]

    def visible(self,locator):
        return [e for e in locator.all() if e.is_visible()]

    def exact(self,text,page=None,buttons=False):
        pattern=re.compile(r'^\s*'+r'\s*'.join(re.escape(c) for c in norm(text))+r'\s*$')
        all_out=[]
        for f in self.frames(page):
            selectors = 'button,a,input[type=button],input[type=submit],[role=button],label' if buttons else 'button,a,label,span,input[type=button],input[type=submit],[role=button]'
            candidates=self.visible(f.locator(selectors))
            out=[]
            for e in candidates:
                if e.get_attribute('type') in ('button','submit'):
                    label=e.get_attribute('value') or ''
                else: label=e.inner_text()
                if pattern.fullmatch(label): out.append(e)
            # span inside a/button is not a separate click target.
            out=[e for e in out if not e.evaluate("e => e.tagName==='SPAN' && !!e.closest('button,a,label,[role=button]')")]
            all_out+=out
        if len(all_out)==1:return all_out[0]
        if len(all_out)>1:raise RuntimeError(f'같은 이름의 요소 여러 개: {text}. selectors.json으로 범위를 지정하세요.')
        return None

    def click(self,labels,page=None,required=True):
        for text in labels:
            target=self.exact(text,page,buttons=False)
            if target:
                target.scroll_into_view_if_needed();target.click(timeout=5000)
                return target
        if required:raise RuntimeError('버튼/메뉴를 찾지 못했습니다: '+', '.join(labels))
        return None

    def override(self,job,key,page=None,wait=False,optional=False):
        spec=self.config.get('services',{}).get(str(job.service),{}).get(key)
        if not spec:return None
        end=time.monotonic()+(self.args.timeout if wait else 0)
        while True:
            matches=[]
            for f in self.frames(page):
                if spec.get('frame_url_contains') and spec['frame_url_contains'] not in f.url:continue
                matches+=self.visible(f.locator(spec['css']))
            if matches or not wait or time.monotonic()>=end:break
            (page or self.page).wait_for_timeout(100)
        if 'index' in spec:
            n=int(spec['index'])
            if 0<=n<len(matches):return matches[n]
        if len(matches)==1:return matches[0]
        if not matches and optional:return None
        raise RuntimeError(f'설정 선택자가 한 요소를 지정하지 못했습니다: {key}')

    def field(self,job,key):
        override=self.override(job,'field:'+key)
        if override:return override
        names=FIELD_ALIASES.get(key,(key,))
        matches=[]
        for f in self.frames():
            for name in names:
                matches += self.visible(f.get_by_label(name,exact=True))
            # title/aria-label/placeholder을 정확히 대조하고 전역 검색창은 제외합니다.
            for e in self.visible(f.locator('input:not([type=hidden]):not([type=radio]):not([type=checkbox]),select,[role=combobox]')):
                labels=[e.get_attribute(a) or '' for a in ('title','aria-label','placeholder')]
                if any(norm(t)==norm(n) for t in labels for n in names):matches.append(e)
            # WebSquare form 표의 라벨 셀 바로 다음 셀만 사용. 문서 전체 following::input 금지.
            for name in names:
                labels=self.visible(f.locator('th,td').filter(has_text=re.compile(r'^\s*'+re.escape(name)+r'\s*$')))
                for label in labels:
                    matches += self.visible(label.locator('xpath=following-sibling::*[1]').locator('input:not([type=hidden]):not([type=radio]):not([type=checkbox]),select,[role=combobox]'))
        unique={e.evaluate("e=>{if(!e.__vat_key)e.__vat_key=crypto.randomUUID();return e.__vat_key}"):e for e in matches}
        matches=list(unique.values())
        # 동일 라벨 아래 연/기수, 연/월, 분기 범위를 값의 형식으로 구분합니다.
        select_matches=[e for e in matches if e.evaluate("e=>e.tagName==='SELECT'")]
        if key in ('기수','조회월','조회기간','시작분기','종료분기','과세연도','적용분기') and select_matches:
            token={'기수':'기','조회월':'월','조회기간':'기','시작분기':'분기','종료분기':'분기','적용분기':'분기'}.get(key)
            filtered=[]
            for e in select_matches:
                opts=e.locator('option').all_text_contents()
                if key=='과세연도' and any(re.fullmatch(r'20\d{2}년?',norm(t)) for t in opts):filtered.append(e)
                elif token and any(token in norm(t) for t in opts):filtered.append(e)
            matches=filtered or select_matches
        if key in ('시작분기','종료분기','시작일','종료일') and len(matches)==2:
            return matches[0 if key.startswith('시작') else 1]
        if len(matches)==1:return matches[0]
        raise RuntimeError(f'{key} 입력란을 정확히 하나 찾지 못했습니다 ({len(matches)}개).')

    def set_field(self,job,key,value):
        if key=='사업자등록번호' and job.service==4 and self.set_split_biz(job,value):return
        if key=='특례조회탭':
            self.click((value,));return
        if key in RADIO_FIELDS:
            target=self.override(job,'field:'+key)
            if target:
                tag=target.evaluate('e=>e.tagName')
                if tag=='SELECT':self.select(target,value);return
                if target.get_attribute('type')=='radio':target.check();return
                target.click();return
            target=self.exact(value)
            if not target:raise RuntimeError(f'{key} 선택을 찾지 못했습니다: {value}')
            target.click()
            linked=target.get_attribute('for')
            if linked:
                checked=target.evaluate('(e,id)=>{const a=e.ownerDocument.getElementById(id);return a? a.checked===true:null}',linked)
                if checked is False:raise RuntimeError(f'{key} 선택 확인 실패')
            return
        e=self.field(job,key)
        tag=e.evaluate('e=>e.tagName')
        if tag=='SELECT':self.select(e,value);return
        if e.get_attribute('role')=='combobox':
            e.click();self.click((value,));return
        if key in ('기수','조회월','적용분기','시작분기','종료분기','조회기간','품목구분','매입매출구분'):
            # 커스텀 콤보박스를 텍스트 input처럼 입력하지 않습니다.
            raise RuntimeError(f'{key}는 커스텀 선택 컨트롤입니다. 선택자 설정/수동확인이 필요합니다.')
        if e.get_attribute('readonly') is not None:raise RuntimeError(f'{key}는 읽기 전용입니다.')
        e.fill(value)
        e.press('Tab')
        actual=e.input_value()
        if re.sub(r'[-\s]','',actual)!=re.sub(r'[-\s]','',value):
            raise RuntimeError(f'{key} 입력값 확인 실패')

    def set_split_biz(self,job,value):
        groups=[]
        for f in self.frames():
            for name in FIELD_ALIASES['사업자등록번호']:
                for label in self.visible(f.locator('th,td').filter(has_text=re.compile(r'^\s*'+re.escape(name)+r'\s*$'))):
                    fields=self.visible(label.locator('xpath=following-sibling::*[1]').locator('input[type=text],input:not([type])'))
                    if len(fields)==3 and [e.get_attribute('maxlength') for e in fields]==['3','2','5']:
                        groups.append(fields)
        if not groups:return False
        if len(groups)!=1:raise RuntimeError('분할 사업자번호 입력란이 여러 개입니다.')
        for e,part in zip(groups[0],(value[:3],value[3:5],value[5:])):
            e.fill(part);e.press('Tab')
            if e.input_value()!=part:raise RuntimeError('분할 사업자번호 입력 확인 실패')
        return True

    def select(self,e,value):
        options=e.locator('option').all()
        exact=[o for o in options if norm(o.inner_text()).strip('-')==norm(value).strip('-')]
        if not exact and re.fullmatch(r'\d{4}',value):
            exact=[o for o in options if norm(o.inner_text())==value+'년']
        if len(exact)!=1:raise RuntimeError(f'선택값을 정확히 찾지 못했습니다: {value}')
        v=exact[0].evaluate('e=>e.value')  # HTML option은 value 속성 생략 시 텍스트를 값으로 사용
        e.select_option(value=v)
        if e.input_value()!=v:raise RuntimeError('선택값 확인 실패')

    def wait_busy(self):
        end=time.monotonic()+self.args.timeout
        while time.monotonic()<end:
            busy=any(any(e.is_visible() for e in f.locator('[aria-busy=true],.w2modal,.w2loading').all()) for f in self.frames())
            if not busy:return
            self.page.wait_for_timeout(200)
        raise RuntimeError('조회 처리중 표시가 사라지지 않았습니다.')

    def navigate(self,job):
        self.page.bring_to_front()
        nav=self.override(job,'menu')
        if nav:nav.click();return
        try:
            self.click(('세무대리·납세관리','세무대리ㆍ납세관리','세무대리 납세관리'))
            self.click(('세무대리 공통 조회',))
            self.click(MENUS[job.service])
            self.wait_busy()
        except Exception:
            if self.args.mode=='strict':raise
            prompt(f'Chrome에서 [세무대리·납세관리 → 세무대리 공통 조회 → {MENUS[job.service][0]}]를 열어 주세요.')

    def select_client(self,job):
        if job.service!=2:return
        # 전자(세금)계산서는 입력란이 아닌 수임사업자 전환 모달을 사용합니다.
        self.click(('수임사업자전환','수임사업자 전환'))
        for f in self.frames():
            candidates=self.visible(f.locator('[role=dialog],.w2popup_window'))
            for modal in candidates:
                if '수임사업자' not in modal.inner_text():continue
                combos=self.visible(modal.locator('select'))
                possible=[]
                for e in combos:
                    if any(norm(t)=='사업자등록번호' for t in e.locator('option').all_text_contents()):possible.append(e)
                if len(possible)!=1:raise RuntimeError('수임사업자번호 검색방식 선택자 확인 필요')
                self.select(possible[0],'사업자등록번호')
                # 번호 입력란은 모달 안으로 한정합니다. ID를 추측하지 않습니다.
                inputs=self.visible(modal.locator('input[type=text],input:not([type])'))
                numbered=[e for e in inputs if '사업자' in (e.get_attribute('title') or '') or e.get_attribute('maxlength') in ('10','12')]
                if len(numbered)!=1:raise RuntimeError('수임사업자 전환 번호 입력란 확인 필요')
                numbered[0].fill(job.client.biz)
                buttons=modal.get_by_text('조회',exact=True)
                if len(self.visible(buttons))!=1:raise RuntimeError('모달 조회 버튼 모호')
                self.visible(buttons)[0].click();self.wait_busy()
                rows=self.visible(modal.locator('tr,[role=row]'))
                rows=[r for r in rows if job.client.biz in re.sub(r'[-\s]','',r.inner_text())]
                if len(rows)!=1:raise RuntimeError('정확한 사업자번호의 수임사업자 결과가 한 건이 아닙니다.')
                radios=self.visible(rows[0].locator('input[type=radio]'))
                if len(radios)!=1:raise RuntimeError('수임사업자 결과 선택 컨트롤 확인 필요')
                radios[0].check()
                confirm=self.visible(modal.get_by_text('확인',exact=True))
                if len(confirm)!=1:raise RuntimeError('수임사업자 확인 버튼 모호')
                confirm[0].click();self.wait_busy();return
        raise RuntimeError('수임사업자 전환 모달을 찾지 못했습니다.')

    def configure_query(self,job):
        failures=[]
        try:self.select_client(job)
        except Exception as exc:failures.append(str(exc))
        ordered=list(job.filters.items())
        ordered.sort(key=lambda item: 0 if item[0]=='특례조회탭' else (1 if item[0] in RADIO_FIELDS else 2))
        for key,value in ordered:
            if job.service==2 and key in ('사업자등록번호','과세연도'):continue  # 연도는 시작/종료일에 포함
            try:self.set_field(job,key,value)
            except Exception as exc:failures.append(str(exc))
        if failures:
            if self.args.mode=='strict':raise RuntimeError('; '.join(failures))
            masked={k:(f'끝 4자리 {value[-4:]}' if k=='사업자등록번호' else value) for k,value in job.filters.items()}
            print('자동 설정하지 못한 항목:', '; '.join(failures))
            print('이번 작업의 필수 조건:',json.dumps(masked,ensure_ascii=False))
            prompt('해당 거래처의 사업자번호 전체를 엑셀과 대조하고 위 조건을 모두 설정하세요. 수임사업자 전환이 필요한 경우 전환도 완료하세요.')
        self.dialogs.clear()
        search=self.override(job,'search')
        if search:search.click()
        else:self.click(('조회',))
        self.page.wait_for_timeout(600)
        self.wait_busy()
        # 기존 결과가 남는 비동기 페이지는 사람이 확인하거나 result_ready 선택자로 검증합니다.
        ready=self.override(job,'result_ready',wait=True)
        if ready:ready.wait_for(state='visible')
        fatal=[m for m in self.dialogs if any(t in m for t in ('권한','로그인','유효하지','잘못','필수','실패','오류'))]
        if fatal:raise RuntimeError('홈택스가 조회 오류/권한/로그인 알림을 표시했습니다. 화면을 확인하세요.')
        if self.args.mode=='assisted':
            prompt(f'[{job.client.company} / 번호 끝 {job.client.biz[-4:]} / {job.client.period.label} / {job.variant}]\n조회가 완료되었는지, 결과의 거래처·조회기간이 맞는지 확인하세요. 조회결과 없음도 이 단계에서 확인합니다.')
            return True
        # strict에서 입력창 값만 보고 대상·결과가 검증되었다고 간주하지 않습니다.
        proof=self.override(job,'result_identity')
        if not proof or not ready:raise RuntimeError('strict 모드는 result_ready와 result_identity 선택자가 필요합니다.')
        self.verify_identity(job,proof)
        return True

    def verify_identity(self,job,element):
        text=re.sub(r'[-\s]','',element.inner_text())
        if job.client.biz not in text or str(job.client.period.year) not in text:
            raise RuntimeError('결과 영역에서 사업자번호/연도를 확인하지 못했습니다.')
        for key,value in job.filters.items():
            if key in ('사업자등록번호','과세연도','작성일자','특례조회탭'):continue
            if re.sub(r'[-\s]','',value) not in text:
                raise RuntimeError(f'결과 영역에서 {key} 조건을 확인하지 못했습니다.')

    def no_data(self,job):
        custom=self.override(job,'no_data',optional=True)
        if custom:return True
        # 3번은 한 표가 비어도 다른 표에 매출이 있으므로 통합 무자료로 판단하지 않습니다.
        if job.service==3:return False
        bodies=[f.locator('body').inner_text() for f in self.frames()]
        text='\n'.join(bodies)
        total=re.findall(r'총\s*([0-9,]+)\s*건',text)
        if total:return all(int(n.replace(',',''))==0 for n in total)
        # 모호한 공통 '조회결과 없음'은 판정하지 않고 일반 출력 절차를 거칩니다.
        return False

    def download(self,job,target,labels):
        button=self.override(job,'download')
        if not button:
            for label in labels:
                button=self.exact(label)
                if button:break
        if not button:raise RuntimeError('다운로드 버튼을 찾지 못했습니다.')
        with self.page.expect_download(timeout=self.args.timeout*1000) as event:
            button.click()
        download=event.value
        if download.failure():raise RuntimeError('브라우저 다운로드 실패')
        name=Path(download.suggested_filename).name
        suffix=Path(name).suffix.lower()
        if not suffix:raise RuntimeError('다운로드 파일의 확장자가 없습니다.')
        dest=target/(job.stem+suffix)
        if dest.exists():raise RuntimeError('같은 이름의 새 다운로드가 중복됩니다.')
        download.save_as(str(dest))
        return self.artifact(dest,'native_download')

    def generate_download(self,job,target):
        generation=self.override(job,'generate')
        if generation:generation.click()
        else:self.click(('파일생성','파일 생성'))
        # 파일 생성은 화면 안내대로 시간이 걸립니다. 재클릭하지 않고 활성화/안내를 기다립니다.
        end=time.monotonic()+self.args.generation_timeout
        saw_disabled=False
        while time.monotonic()<end:
            if any(any(t in m for t in ('실패','오류','권한')) for m in self.dialogs):
                raise RuntimeError('파일생성 오류. 홈택스 알림을 확인하세요.')
            button=self.override(job,'download',optional=True)
            if not button:button=self.exact('다운로드')
            enabled=bool(button and button.is_enabled() and button.get_attribute('aria-disabled')!='true')
            if not enabled:saw_disabled=True
            completed=any('파일생성' in norm(m) and '완료' in m for m in self.dialogs)
            if enabled and (saw_disabled or completed):
                return self.download(job,target,('다운로드',))
            self.page.wait_for_timeout(500)
        raise RuntimeError('파일생성 시간초과. 다음 실행에서 다시 조회/생성할 수 있습니다.')

    def artifact(self,path,source):
        return {'path':str(path.resolve()),'source':source,**validate_artifact(path)}

    def screen_pdf(self,job,target,no_data=False):
        # 보고서 PDF와 구별되는 조회화면 증빙입니다. 테이블의 전체 페이지를 대체하지 않습니다.
        session=self.context.new_cdp_session(self.page)
        try:
            import base64
            result=session.send('Page.printToPDF',{'printBackground':True,'landscape':True,'preferCSSPageSize':False})
            path=target/(job.stem+'_조회화면증빙.pdf')
            path.write_bytes(base64.b64decode(result['data']))
        finally:session.detach()
        return self.artifact(path,'no_data_screen' if no_data else 'query_screen')

    def report_pdf(self,job,target):
        started_at=time.time()
        incoming_before={p for p in self.incoming.iterdir() if p.is_file()}
        pages_before=set(self.context.pages)
        button=self.override(job,'report')
        if button:button.click()
        elif job.service==2:self.click(('명세서 출력','출력'))
        else:self.click(('인쇄하기','인쇄'))
        end=time.monotonic()+10
        while time.monotonic()<end and not (set(self.context.pages)-pages_before):
            self.page.wait_for_timeout(200)
        new=list(set(self.context.pages)-pages_before)
        if len(new)>1:raise RuntimeError('여러 리포트창이 열렸습니다. 출력 대상을 확인하세요.')
        report=new[0] if new else self.page
        report.bring_to_front()
        report.wait_for_load_state('domcontentloaded')
        # WebSquare/CLIP report 로딩 완료 확인. 제목만 열린 빈 팝업은 성공 처리하지 않습니다.
        report.wait_for_timeout(1500)
        if self.args.mode=='strict':
            ready=self.override(job,'report_ready',report,wait=True)
            identity=self.override(job,'report_identity',report)
            if not ready or not identity:raise RuntimeError('strict 모드는 report_ready와 report_identity 선택자가 필요합니다.')
            self.verify_identity(job,identity)
            if not self.kiosk:raise RuntimeError('strict PDF 인쇄는 이 실행에서 준비한 전용 Chrome이 필요합니다.')
        if not hometax_host(report.url) and report.url not in ('','about:blank'):
            raise RuntimeError('홈택스 외부 리포트창입니다. URL을 확인하세요.')
        if self.args.mode=='assisted':
            prompt('리포트의 거래처·기간·전체 페이지가 맞고 로딩이 끝났는지 확인하세요. 아직 인쇄하지 말고 Enter를 누르세요.')
        print_button=self.override(job,'print',report)
        if not print_button:
            # 프린터 아이콘의 접근성 속성만 사용합니다. 임의 좌표 클릭 금지.
            icons=[]
            for f in self.frames(report):
                icons+=self.visible(f.locator('button[title="인쇄"],a[title="인쇄"],button[aria-label="인쇄"],a[aria-label="인쇄"]'))
            if len(icons)==1:icons[0].click()
            print_button=self.exact('인쇄',report)
        if print_button:
            # 전체페이지 라디오가 존재하면 명확히 선택합니다.
            full=self.exact('전체 페이지',report) or self.exact('전체페이지',report)
            if full:full.click()
            if not self.kiosk and self.args.mode=='assisted':
                print('이 Chrome은 연결된 기존 창입니다. 인쇄 미리보기에서 [PDF로 저장]을 선택하세요.')
            print_button.click()
            if not self.kiosk and self.args.mode=='assisted':
                prompt(f'인쇄 미리보기에서 [PDF로 저장]으로 저장하세요. 저장 위치: {self.incoming}. 저장 완료 후 Enter.')
        else:
            if self.args.mode=='strict':raise RuntimeError('리포트 인쇄 버튼 선택자 확인 필요')
            prompt('리포트 프린터 아이콘 → 전체페이지 → 인쇄를 눌러 주세요. Chrome에서는 [PDF로 저장]을 선택하세요.\n새 PDF는 실행 폴더의 _incoming에 저장하세요. PDF 준비 후 Enter.')
        try:
            return self.wait_print_file(job,target,incoming_before)
        except Exception:
            if self.args.mode=='strict':raise
            print('자동 PDF 저장이 확인되지 않았습니다.')
            path=input('PDF를 수동 저장한 다음 파일의 전체 경로 입력 (s: 건너뛰기): ').strip().strip('"')
            if path.lower()=='s':raise RuntimeError('리포트 저장 건너뜀')
            source=Path(path)
            validate_artifact(source)
            if source.suffix.lower()!='.pdf':raise RuntimeError('PDF만 지정할 수 있습니다.')
            if source.resolve() in {p.resolve() for p in incoming_before} or source.stat().st_mtime < started_at-2:
                raise RuntimeError('이전 작업 파일은 사용할 수 없습니다. 이번 조회 PDF를 새로 저장하세요.')
            dest=target/(job.stem+'.pdf');shutil.copy2(source,dest)
            return self.artifact(dest,'manual_report')
        finally:
            for p in new:
                if not p.is_closed():p.close()
            if not self.page.is_closed():self.page.bring_to_front()

    def wait_print_file(self,job,target,before):
        end=time.monotonic()+self.args.timeout
        previous={}
        while time.monotonic()<end:
            candidates=[p for p in self.incoming.iterdir() if p.is_file() and p not in before and p.suffix.lower()=='.pdf']
            if len(candidates)>1:raise RuntimeError('새 PDF가 여러 개여서 출력 대상을 판정할 수 없습니다.')
            if len(candidates)==1:
                p=candidates[0];size=p.stat().st_size
                if size>0 and previous.get(p)==size:
                    validate_artifact(p)
                    dest=target/(job.stem+'.pdf');shutil.move(str(p),str(dest))
                    return self.artifact(dest,'clipreport_print')
                previous[p]=size
            self.page.wait_for_timeout(500)
        raise RuntimeError('PDF 저장 시간초과')

    def collect(self,job,target):
        self.navigate(job)
        confirmed=self.configure_query(job)
        artifacts=[]
        empty=self.no_data(job)
        notes=[]
        # 특례 탭 외에는 화면 전체를 원본 데이터인 것처럼 XLSX로 변환하지 않습니다.
        if '원본' in job.client.formats and job.service in (5,6,7):
            if empty:notes.append('조회 건수 0: 원본 다운로드를 시도하지 않았음')
            elif job.service in (6,7):artifacts.append(self.generate_download(job,target))
            else:artifacts.append(self.download(job,target,('엑셀 내려받기','엑셀내려받기','엑셀 다운로드','엑셀다운로드')))
        if 'PDF' in job.client.formats:
            if job.service==5 or empty:
                artifacts.append(self.screen_pdf(job,target,no_data=empty))
                notes.append('조회화면 증빙 PDF: 전체 거래 원본/공식 보고서를 대체하지 않음')
            else:artifacts.append(self.report_pdf(job,target))
        if '원본' in job.client.formats and job.service in (1,2,3,4):
            notes.append('이 구현은 해당 항목의 PDF를 수집함; 엑셀 원본 미수집')
        if not artifacts:raise RuntimeError('요청 형식에 대해 수집한 파일이 없습니다. PDF 또는 모두로 지정하세요.')
        return artifacts,confirmed,'; '.join(notes)


def parser():
    p=argparse.ArgumentParser(description=__doc__,formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--input',type=Path,default=ROOT/'INPUT-부가세_확장.xlsx')
    p.add_argument('--sheet')
    p.add_argument('--year',type=int,help='과세연도 열이 없는 엑셀의 보완값. 엑셀 값이 우선.')
    p.add_argument('--output',type=Path,default=ROOT/'다운로드')
    p.add_argument('--profile',type=Path,default=ROOT/'.chrome-profile')
    p.add_argument('--chrome',help='Chrome 실행파일 경로')
    p.add_argument('--port',type=int,default=9222)
    p.add_argument('--mode',choices=('assisted','strict'),default='assisted')
    p.add_argument('--selectors',type=Path,help='실제 화면에서 확인한 선택자 JSON')
    p.add_argument('--timeout',type=int,default=60)
    p.add_argument('--generation-timeout',type=int,default=360)
    p.add_argument('--validate-only',action='store_true')
    p.add_argument('--services',help='이번 실행의 항목만 선택, 예: 1,2,7')
    p.add_argument('--inspect',action='store_true',help='로그인 후 현재 화면의 제어요소 메타데이터 저장. 입력값/쿠키 제외.')
    p.add_argument('--resume',action='store_true',help='저장기록·확인상태·파일 해시가 유효한 작업만 건너뜀')
    return p


def main(argv=None):
    args=parser().parse_args(argv)
    if not 1024<=args.port<=65535 or min(args.timeout,args.generation_timeout)<=0:
        print('포트/시간 제한값이 잘못되었습니다.',file=sys.stderr);return 2
    try:
        clients=load_clients(args.input,args.year,args.sheet)
        if args.services:
            ids={int(x) for x in args.services.split(',')}
            if not ids or not ids.issubset(SERVICES):raise ValueError('--services는 1~7이어야 합니다.')
            clients=[replace(c,services=tuple(s for s in c.services if s in ids)) for c in clients]
            clients=[c for c in clients if c.services]
        jobs=build_jobs(clients)
        if not jobs:raise ValueError('선택된 작업이 없습니다.')
        if any(j.client.formats==('원본',) and j.service in (1,2,3,4) for j in jobs):
            raise ValueError('1~4번은 이 구현에서 원본 전용 저장을 지원하지 않습니다. 저장형식을 PDF 또는 모두로 지정하세요.')
        print(f'거래처 {len(clients)}곳, 세부 작업 {len(jobs)}건')
        for j in jobs:
            print(f'  엑셀 {j.client.row}행 / {j.client.period.label} / {j.service} / {j.variant}')
        if args.validate_only:
            print('입력·조회계획 검증 완료. 홈택스 로그인/조회/다운로드는 실행하지 않았습니다.');return 0
        from playwright.sync_api import sync_playwright
        args.output=args.output.resolve();args.output.mkdir(parents=True,exist_ok=True)
        incoming=args.output/'_incoming';incoming.mkdir(exist_ok=True)
        journal=Journal(args.output/'실행기록.jsonl')
        success=failed=skipped=0
        with sync_playwright() as pw:
            kiosk=launch_chrome(args,incoming)
            browser=pw.chromium.connect_over_cdp(f'http://127.0.0.1:{args.port}')
            context=browser.contexts[0]
            pages=[p for p in context.pages if hometax_host(p.url) and 'clipreport' not in p.url]
            page=pages[0] if pages else context.new_page()
            if not pages:page.goto(URL,wait_until='domcontentloaded')
            print('전용 Chrome에서 [세무대리인] 로그인 및 필요한 수임 동의를 완료하세요. 인증정보를 입력 엑셀에 저장하지 마세요.')
            end=time.monotonic()+600
            while time.monotonic()<end:
                if any(f.get_by_text('로그아웃',exact=True).count() and any(e.is_visible() for e in f.get_by_text('로그아웃',exact=True).all()) for f in page.frames):break
                page.wait_for_timeout(1000)
            else:raise RuntimeError('로그인 대기 10분 초과')
            collector=Collector(context,page,args,incoming,kiosk)
            if args.inspect:
                prompt('선택자를 확인할 조회화면을 열어 주세요. 고객 자료를 입력할 필요는 없습니다.')
                controls=[]
                for frame in collector.frames():
                    for e in collector.visible(frame.locator('input,select,button,a,[role=combobox]')):
                        if e.get_attribute('type') in ('password','hidden'):continue
                        item={k:e.get_attribute(k) for k in ('id','name','title','aria-label','role','type','maxlength')}
                        item['tag']=e.evaluate('e=>e.tagName')
                        item['frame_path']=urlparse(frame.url).path
                        if item['tag']=='SELECT':item['options']=e.locator('option').all_text_contents()
                        # 버튼·링크 텍스트에는 고객명이 있을 수 있어 기본 수집하지 않습니다.
                        controls.append(item)
                dest=args.output/('controls_'+time.strftime('%Y%m%d_%H%M%S')+'.json')
                dest.write_text(json.dumps(controls,ensure_ascii=False,indent=2),encoding='utf-8')
                print(f'선택자 점검자료 저장: {dest}. 공유 전 개인정보 포함 여부를 확인하세요.')
                return 0
            for job in jobs:
                if args.resume and journal.reusable(job):
                    skipped+=1;print(f'재사용: {job.service}/{job.variant}');continue
                # 고유 실행 폴더: 기존 파일을 삭제/덮어쓰지 않습니다.
                target=args.output/job.key/time.strftime('%Y%m%d_%H%M%S')
                n=0
                while target.exists():
                    n+=1;target=args.output/job.key/(time.strftime('%Y%m%d_%H%M%S')+f'_{n}')
                target.mkdir(parents=True)
                metadata={'job_key':job.key,'row':job.client.row,'company':job.client.company,
                          'service':job.service,'variant':job.variant,'filters':job.filters,
                          'formats':job.client.formats}
                (target/'조회조건.json').write_text(json.dumps(metadata,ensure_ascii=False,indent=2),encoding='utf-8')
                print(f'처리: 엑셀 {job.client.row}행 / 항목 {job.service} / {job.variant}')
                try:
                    artifacts,confirmed,note=collector.collect(job,target)
                    journal.record(job,'saved',artifacts,confirmed,note);success+=1
                except KeyboardInterrupt:
                    journal.record(job,'interrupted',[],False,'사용자가 중단함')
                    raise
                except Exception as exc:
                    # 브라우저 예외 본문에 DOM/개인정보가 포함될 수 있어 원문 로그를 저장하지 않습니다.
                    journal.record(job,'failed',[],False,type(exc).__name__)
                    failed+=1
                    detail=str(exc) if isinstance(exc,(RuntimeError,ValueError)) else '브라우저 요소/대기 오류. 해당 화면과 선택자를 확인하세요.'
                    print(f'실패: {type(exc).__name__}: {detail}')
                    if args.mode=='assisted':
                        print('다음 작업으로 넘어갑니다. 실패한 작업은 --resume 실행 시 다시 처리합니다.')
                finally:
                    for popup in list(collector.owned):
                        if popup!=page and not popup.is_closed():popup.close()
                    collector.owned.clear()
                    if page.is_closed():raise RuntimeError('메인 홈택스 창이 닫혔습니다.')
                    page.bring_to_front()
            # 로그인된 사용자 브라우저는 그대로 두고 Playwright 연결만 문맥 종료로 해제합니다.
        print(f'완료: 저장 {success}, 실패 {failed}, 이전 파일 재사용 {skipped}')
        print(f'결과/실행기록: {args.output}')
        return 1 if failed else 0
    except KeyboardInterrupt:
        print('\n중단되었습니다. 다음 실행에 --resume을 사용하세요.');return 130
    except Exception as exc:
        # 입력/초기화 오류는 비밀값을 받지 않으므로 원인을 출력합니다.
        print(f'실행 중단: {exc}',file=sys.stderr);return 2


if __name__=='__main__':
    raise SystemExit(main())
