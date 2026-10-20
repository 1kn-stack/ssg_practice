"""Independent source/normalized/report verifier; reads originals without modification."""
from pathlib import Path
import argparse, datetime as dt, hashlib, itertools, json, zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
FIELDS = ['record_id','기준주','점포','브랜드','장르','채널','매출_만원','목표_만원','비용_만원','상태','출처','기준일']
ALIASES = [FIELDS,
 ['레코드ID','주차','점포명','브랜드명','카테고리','유통채널','실적(만원)','목표매출(만원)','집행비용(만원)','진행상태','근거자료','확인일'],
 ['기록번호','주간기준일','영업점','행사브랜드','상품군','판매채널','매출액_만원','목표액_만원','사용비용_만원','진행현황','자료출처','기준일자']]
REPORT = ['record_id','채널','기준주','점포','브랜드','장르','매출_만원','목표_만원','목표차이_만원','달성률','비용_만원','비용차감액_만원','상태','메모','출처','기준일']

def read_xlsx(path):
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            shared = [''.join(si.itertext()) for si in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        tree = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        rows=[]
        for row in tree.findall('s:sheetData/s:row',NS):
            values={}
            for cell in row.findall('s:c',NS):
                letters=''.join(c for c in cell.attrib['r'] if c.isalpha())
                col=0
                for c in letters: col=col*26+ord(c)-64
                kind=cell.get('t'); value=cell.findtext('s:v',default='',namespaces=NS)
                if kind=='s': value=shared[int(value)]
                elif kind=='inlineStr': value=''.join(cell.find('s:is',NS).itertext())
                elif kind not in ('str','b','e') and value!='': value=float(value); value=int(value) if value.is_integer() else value
                values[col-1]=value
            rows.append([values.get(i,'') for i in range(max(values,default=-1)+1)])
        return rows

def source_rows():
    all_rows=[]; summaries=[]
    for i,(name,channel) in enumerate(zip(['01_영업실적.xlsx','02_온라인실적.xlsx','03_팝업실적.xlsx'],['영업','온라인','팝업'])):
        raw=read_xlsx(ROOT/'A_매출비교'/name); headers=raw[0]
        assert len(headers)==12 and set(headers)==set(ALIASES[i]), (name,headers)
        rows=[]
        for number,values in enumerate(raw[1:],2):
            original=dict(zip(headers,values)); normalized={field:original[alias] for field,alias in zip(FIELDS,ALIASES[i])}
            for key in ['기준주','기준일']:
                if isinstance(normalized[key],(float,int)): normalized[key]=(dt.date(1899,12,30)+dt.timedelta(days=normalized[key])).isoformat()
            assert normalized['채널']==channel,(name,number,'channel')
            assert normalized['출처']==name,(name,number,'source')
            rows.append(normalized)
        assert len(rows)==160 and len({r['record_id'] for r in rows})==160 and all(r['record_id'] for r in rows)
        summaries.append({'file':name,**total(rows),'zero_target_rows':sum(r['목표_만원']==0 for r in rows)})
        all_rows.extend(rows)
    assert len({r['record_id'] for r in all_rows})==480
    return all_rows,summaries

def total(rows):
    return {'rows':len(rows),'sales':sum(r['매출_만원'] for r in rows),'target':sum(r['목표_만원'] for r in rows),'cost':sum(r['비용_만원'] for r in rows)}

def fixtures():
    directory=ROOT/'results/verification'; directory.mkdir(parents=True,exist_ok=True)
    source=ROOT/'A_매출비교/01_영업실적.xlsx'
    for name,kind in [('01_영업실적_파일내중복.xlsx','duplicate'),('01_영업실적_채널불일치.xlsx','channel')]:
        with zipfile.ZipFile(source) as original, zipfile.ZipFile(directory/name,'w',zipfile.ZIP_DEFLATED) as output:
            for item in original.infolist():
                data=original.read(item.filename)
                if item.filename=='xl/worksheets/sheet1.xml':
                    tree=ET.fromstring(data); rows=tree.findall('s:sheetData/s:row',NS)
                    if kind=='duplicate':
                        a=rows[1].find('s:c/s:v',NS).text
                        rows[2].find('s:c/s:v',NS).text=a
                    else:
                        rows[1].find("s:c[@r='F2']/s:v",NS).text='온라인'
                    data=ET.tostring(tree,encoding='utf-8',xml_declaration=True)
                output.writestr(item,data)
    oracle=[]
    for name in ['01_영업실적.xlsx','02_온라인실적.xlsx','03_팝업실적.xlsx','05_오류_열누락.xlsx']:
        raw=read_xlsx(ROOT/'A_매출비교'/name)
        oracle.append({'name':name,'headers':raw[0],'body':raw[1:]})
    (directory/'source-oracle.json').write_text(json.dumps(oracle,ensure_ascii=False),encoding='utf-8')

def main():
    parser=argparse.ArgumentParser(); parser.add_argument('--normalized',default='results/sales-app/data/sales.json'); parser.add_argument('--report'); parser.add_argument('--expected-memo'); args=parser.parse_args()
    source,summaries=source_rows(); template=read_xlsx(ROOT/'A_매출비교/04_보고양식.xlsx')
    assert template[0]==REPORT
    bad=read_xlsx(ROOT/'A_매출비교/05_오류_열누락.xlsx')
    assert '목표_만원' not in bad[0] and len(bad[1:])==3
    normalized=json.loads((ROOT/args.normalized).read_text(encoding='utf-8-sig'))
    actual=normalized['rows']; assert normalized['reportHeaders']==REPORT
    assert len(actual)==480 and len({r['record_id'] for r in actual})==480
    expected_by_id={r['record_id']:r for r in source}
    for row in actual:
        assert {key:row[key] for key in FIELDS}==expected_by_id[row['record_id']],row['record_id']
        assert row['_sourceFile']==row['출처'] and row['_sourceSheet']=='원본' and isinstance(row['_sourceRow'],int)
        expected_source_row=2+[r['record_id'] for r in source if r['출처']==row['출처']].index(row['record_id'])
        assert row['_sourceRow']==expected_source_row
    filter_results=[]
    choices=[[None]+sorted({r[key] for r in source}) for key in ['점포','채널','기준주']]
    for store,channel,week in itertools.product(*choices):
        condition={key:value for key,value in zip(['점포','채널','기준주'],[store,channel,week]) if value is not None}
        expected=[r for r in source if all(r[k]==v for k,v in condition.items())]
        selected=[r for r in actual if all(r[k]==v for k,v in condition.items())]
        assert total(expected)==total(selected),condition
        filter_results.append({'filters':condition,**total(expected)})
    hashes=[]
    for entry in json.loads((ROOT/'results/source-hashes-before.json').read_text(encoding='utf-8-sig')):
        current=hashlib.sha256(Path(entry['Path']).read_bytes()).hexdigest().upper(); assert current==entry['Hash'],entry['Path']
        hashes.append({'path':entry['Path'],'unchanged':True,'sha256':current})
    fixtures()
    result={'source_files':summaries,'total':total(source),'total_attainment':total(source)['sales']/total(source)['target'],'zero_target_rows':sum(r['목표_만원']==0 for r in source),'filter_combinations':len(filter_results),'filter_results':filter_results,'hashes':hashes,'headers':REPORT,'checked_fields':12,'row_by_row_match':True}
    (ROOT/'results/verification/source-comparison.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k not in ['filter_results','hashes','headers']},ensure_ascii=True,indent=2))
    if args.report:
        report=read_xlsx(ROOT/args.report); assert report[0]==REPORT
        assert len({values[0] for values in report[1:]})==len(report)-1
        report_rows=[dict(zip(REPORT,values)) for values in report[1:]]
        if report_rows:
            inferred_filter={key:report_rows[0][key] for key in ['점포','채널','기준주'] if len({r[key] for r in report_rows})==1}
            expected_report=[r for r in source if all(r[key]==value for key,value in inferred_filter.items())]
            assert {r['record_id'] for r in report_rows}=={r['record_id'] for r in expected_report},inferred_filter
        zero_blank=0
        for values in report[1:]:
            row=dict(zip(REPORT,values)); original=expected_by_id[row['record_id']]
            assert all(row[k]==original[k] for k in FIELDS),row['record_id']
            assert row['목표차이_만원']==row['매출_만원']-row['목표_만원']; assert row['비용차감액_만원']==row['매출_만원']-row['비용_만원']
            if row['목표_만원']==0:
                assert row['달성률']=='',row['record_id']; zero_blank+=1
            else: assert abs(row['달성률']-row['매출_만원']/row['목표_만원'])<1e-12
        if args.expected_memo is not None:
            memo=[dict(zip(REPORT,v))['메모'] for v in report[1:] if v[0]=='SL-20260907-A-01']
            assert memo==[args.expected_memo],memo
        print('REPORT VERIFIED',len(report)-1,'ZERO-GOAL BLANK',zero_blank,'SOURCE FIELDS EXACT', 'MEMO EXACT' if args.expected_memo is not None else 'MEMO NOT ASSERTED')

if __name__=='__main__': main()
