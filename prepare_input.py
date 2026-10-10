"""기존 엑셀을 변경하지 않고 확장본 생성. 로그인 정보는 넣지 않습니다."""
import argparse
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

EXTRA = {
 '과세연도':2026, '조회범위':'분기', '조회항목':'1,2,3,4,5,6,7',
 '계산서종류':'모두','매출매입구분':'모두','무역서류종류':'모두',
 '특례조회탭':'거래내역조회','특례품목':'전체','특례매입매출구분':'전체',
 '저장형식':'모두','실행여부':'Y',
}
DESCRIPTION = {
 '과세연도':'필수. 4자리 귀속연도. 사용자 답변에 따라 2026 설정.',
 '조회범위':'예정=분기. 확정 기본=반기. 확정 분기만은 7번 미지원.',
 '조회항목':'1~7 중 실행할 항목. 쉼표 구분.',
 '계산서종류':'전자세금계산서/전자계산서/모두. 쉼표로 복수 지정 가능.',
 '매출매입구분':'매출/매입/모두. 사용자 답변에 따라 모두.',
 '무역서류종류':'내국신용장/구매확인서/모두.',
 '특례조회탭':'기본 거래내역조회. 전용계좌조회/국고입금내역/기납부세액은 화면 선택자 추가 확인 필요. 쉼표 구분 가능.',
 '특례품목':'기본 전체. 실제 홈택스 품목 선택항목의 정확한 이름.',
 '특례매입매출구분':'전체/매입/매출. 품목별 자료 구분.',
 '저장형식':'PDF/원본/모두. 1~4는 PDF만 구현. 5~7은 원본+증빙/보고서 PDF.',
 '실행여부':'Y 실행/N 제외.',
}


def prepare(source: Path | None, target: Path, year: int=2026):
    if target.exists():raise FileExistsError(f'기존 파일을 덮어쓰지 않습니다: {target}')
    if source:
        w=load_workbook(source)
    else:
        w=Workbook();w.active.title='거래처'
        w.active.append(['NO','사업자명','대표이사','사업자등록번호','과세기간'])
        w.active.append([1,'예시사업자(교체하세요)','','000-00-00000','2기예정'])
    s=w.active
    headers=[c.value for c in s[1]]
    for name,default in EXTRA.items():
        if name in headers:continue
        col=len(headers)+1;headers.append(name);s.cell(1,col,name)
        for r in range(2,s.max_row+1):
            if not any(s.cell(r,c).value is not None for c in range(1,6)):continue
            value=year if name=='과세연도' else default
            if name=='조회범위':value='반기' if '확정' in str(s.cell(r,headers.index('과세기간')+1).value) else '분기'
            s.cell(r,col,value)
    s.freeze_panes='A2';s.auto_filter.ref=s.dimensions
    for c in s[1]:
        c.font=Font(bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='245782')
        c.alignment=Alignment(wrap_text=True)
        s.column_dimensions[c.column_letter].width=22 if c.value!='조회항목' else 26
    if '사업자등록번호' in headers:
        col=headers.index('사업자등록번호')+1
        for r in range(2,s.max_row+1):s.cell(r,col).number_format='@'
    for name,values in {'과세기간':'1기예정,1기확정,2기예정,2기확정','조회범위':'분기,반기',
            '계산서종류':'모두,전자세금계산서,전자계산서','매출매입구분':'모두,매출,매입',
            '무역서류종류':'모두,내국신용장,구매확인서','특례매입매출구분':'전체,매입,매출',
            '저장형식':'모두,PDF,원본','실행여부':'Y,N'}.items():
        col=s.cell(1,headers.index(name)+1).column_letter
        dv=DataValidation(type='list',formula1='"'+values+'"',allow_blank=True)
        dv.showErrorMessage=True;dv.errorTitle='입력값 확인';dv.error='목록의 값을 선택하세요.'
        s.add_data_validation(dv);dv.add(f'{col}2:{col}10001')
    if '입력안내' not in w.sheetnames:
        guide=w.create_sheet('입력안내');guide.append(['항목','설명'])
        for k,v in DESCRIPTION.items():guide.append([k,v])
        guide.column_dimensions['A'].width=25;guide.column_dimensions['B'].width=105
        for row in guide:
            for c in row:c.alignment=Alignment(wrap_text=True,vertical='top')
        guide.append(['주의','로그인 비밀번호·인증서 비밀번호는 입력하지 않습니다. 대표이사는 현재 자동화에서 사용하지 않습니다.'])
    w.save(target);w.close()


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--source',type=Path);p.add_argument('--target',type=Path,required=True)
    p.add_argument('--year',type=int,default=2026);a=p.parse_args();prepare(a.source,a.target,a.year)
    print('확장 입력파일 생성 완료')
