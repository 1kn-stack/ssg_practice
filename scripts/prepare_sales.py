import json, zipfile, xml.etree.ElementTree as ET, hashlib
from pathlib import Path
from datetime import datetime, timedelta
ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'A_매출비교'
OUT = ROOT / 'results/sales-app/data/sales.json'
NS = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
FIELDS = ['record_id','기준주','점포','브랜드','장르','채널','매출_만원','목표_만원','비용_만원','상태','출처','기준일']
ALIASES = [ ['record_id','레코드ID','기록번호'],['기준주','주차','주간기준일'],['점포','점포명','영업점'],['브랜드','브랜드명','행사브랜드'],['장르','카테고리','상품군'],['채널','유통채널','판매채널'],['매출_만원','실적(만원)','매출액_만원'],['목표_만원','목표매출(만원)','목표액_만원'],['비용_만원','집행비용(만원)','사용비용_만원'],['상태','진행상태','진행현황'],['출처','근거자료','자료출처'],['기준일','확인일','기준일자'] ]
def read_sheet(path):
    with zipfile.ZipFile(path) as z:
        strings=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            strings=[''.join(si.itertext()) for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si',NS)]
        workbook=ET.fromstring(z.read('xl/workbook.xml'))
        sheets=workbook.find('m:sheets',NS)
        sheet=next((s for s in sheets if s.get('name')=='원본'),sheets[0])
        rid=sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        target=next(r.get('Target') for r in rels if r.get('Id')==rid)
        target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
        rows=[]
        for row in ET.fromstring(z.read(target)).findall('m:sheetData/m:row',NS):
            values=[]
            for c in row.findall('m:c',NS):
                letters=''.join(x for x in c.get('r') if x.isalpha())
                index=0
                for ch in letters:index=index*26+ord(ch)-64
                while len(values)<index: values.append('')
                value=c.find('m:v',NS)
                val=value.text if value is not None else ''
                if c.get('t')=='inlineStr':val=''.join(c.find('m:is',NS).itertext())
                elif c.get('t')=='s':val=strings[int(val)]
                elif val!='':
                    try:val=float(val);val=int(val) if val.is_integer() else val
                    except ValueError:pass
                values[index-1]=val
            rows.append((int(row.get('r')),values))
        return sheet.get('name'),rows

def main():
    before={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in SOURCE.glob('*.xlsx')}
    allrows=[];sources=[];seen=set()
    for name,channel in [('01_영업실적.xlsx','영업'),('02_온라인실적.xlsx','온라인'),('03_팝업실적.xlsx','팝업')]:
        sheet,raw=read_sheet(SOURCE/name);headers=raw[0][1]
        indices={field:next((headers.index(a) for a in aliases if a in headers),None) for field,aliases in zip(FIELDS,ALIASES)}
        assert all(i is not None for i in indices.values()),name
        count=0
        for number,values in raw[1:]:
            if not any(v!='' for v in values):continue
            row={f:values[i] if i<len(values) else '' for f,i in indices.items()}
            for field in ['기준주','기준일']:
                if isinstance(row[field], (int,float)): row[field]=(datetime(1899,12,30)+timedelta(days=row[field])).strftime('%Y-%m-%d')
            assert row['record_id'] and row['record_id'] not in seen,(name,number)
            assert row['채널']==channel,(name,number)
            seen.add(row['record_id']);row.update(_sourceFile=name,_sourceSheet=sheet,_sourceRow=number)
            allrows.append(row);count+=1
        assert count==160,(name,count)
        sources.append({'filename':name,'sheet':sheet,'headers':headers,'rows':count,'sha256':before[name]})
    _,template=read_sheet(SOURCE/'04_보고양식.xlsx')
    assert len(allrows)==480
    after={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in SOURCE.glob('*.xlsx')}
    assert before==after,'Original source changed'
    payload={'mode':'educational-sample','notice':'교육용 가상 매출 데이터 · 실제 영업 정보 아님','sources':sources,'reportHeaders':template[0][1],'rows':allrows}
    OUT.parent.mkdir(parents=True,exist_ok=True)
    OUT.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'output':str(OUT),'rows':len(allrows),'sales':sum(r['매출_만원'] for r in allrows),'target':sum(r['목표_만원'] for r in allrows),'cost':sum(r['비용_만원'] for r in allrows),'sourceHashesUnchanged':before==after},ensure_ascii=False))
if __name__=='__main__':main()
