/* Column aliases and calculations follow A_매출비교/06_데이터안내.md. */
(function (global) {
  'use strict';
  const fields = ['record_id','기준주','점포','브랜드','장르','채널','매출_만원','목표_만원','비용_만원','상태','출처','기준일'];
  const aliases = [ ['record_id','레코드ID','기록번호'],['기준주','주차','주간기준일'],['점포','점포명','영업점'],['브랜드','브랜드명','행사브랜드'],['장르','카테고리','상품군'],['채널','유통채널','판매채널'],['매출_만원','실적(만원)','매출액_만원'],['목표_만원','목표매출(만원)','목표액_만원'],['비용_만원','집행비용(만원)','사용비용_만원'],['상태','진행상태','진행현황'],['출처','근거자료','자료출처'],['기준일','확인일','기준일자'] ];
  const reportHeaders = ['record_id','채널','기준주','점포','브랜드','장르','매출_만원','목표_만원','목표차이_만원','달성률','비용_만원','비용차감액_만원','상태','메모','출처','기준일'];
  const numeric = ['매출_만원','목표_만원','비용_만원'];
  function normalize(headers, body, filename, sheetName='원본', rowNumbers=[]) {
    const names=headers.map(v=>String(v??'').trim());
    const positions=aliases.map(list=>names.findIndex(n=>list.includes(n)));
    const missing=fields.filter((_,i)=>positions[i]<0);
    if(missing.length) return {rows:[],errors:[`${filename}: 필수 열 누락 — ${missing.join(', ')}`]};
    const ambiguous=fields.filter((_,i)=>names.filter(n=>aliases[i].includes(n)).length>1);
    if(ambiguous.length) return {rows:[],errors:[`${filename}: 같은 의미의 열이 중복됩니다 — ${ambiguous.join(', ')}`]};
    const expected=({'01_영업실적.xlsx':'영업','02_온라인실적.xlsx':'온라인','03_팝업실적.xlsx':'팝업'})[filename];
    const rows=[], errors=[], seen=new Set();
    body.forEach((values,index)=>{
      if(!values.some(v=>v!==''&&v!==null&&v!==undefined)) return;
      const line=rowNumbers[index]??index+2;
      const row=Object.fromEntries(fields.map((f,i)=>[f,values[positions[i]]??'']));
      row.record_id=String(row.record_id).trim();
      for(const field of ['기준주','기준일']) {
        if(typeof row[field]==='number' && Number.isFinite(row[field])) row[field]=new Date(Date.UTC(1899,11,30)+row[field]*86400000).toISOString().slice(0,10);
      }
      if(!row.record_id){errors.push(`${filename} ${line}행: record_id가 비어 있습니다.`);return;}
      if(seen.has(row.record_id)){errors.push(`${filename} ${line}행: 중복 record_id ${row.record_id} 제외`);return;}
      seen.add(row.record_id);
      if(expected&&row['채널']!==expected){errors.push(`${filename} ${line}행: 채널 '${row['채널']}' 불일치 (기대값 ${expected}). 원본 값은 변경하지 않았습니다.`);return;}
      for(const f of numeric){
        const raw=row[f];
        if(raw===''||raw===null||!Number.isFinite(Number(raw))){errors.push(`${filename} ${line}행: ${f} 숫자가 올바르지 않습니다.`);return;}
        row[f]=Number(raw);
      }
      row._sourceFile=filename;row._sourceSheet=sheetName;row._sourceRow=line;
      rows.push(row);
    });
    return {rows,errors};
  }
  function merge(existing,incoming){
    const rows=[],errors=[],seen=new Set();
    for(const row of [...existing,...incoming]){
      if(!row.record_id){errors.push(`${row._sourceFile??'자료'}: record_id 누락 제외`);continue;}
      if(seen.has(row.record_id)){errors.push(`${row._sourceFile??'자료'}: 중복 record_id ${row.record_id} 제외 (기존 데이터 유지)`);continue;}
      seen.add(row.record_id); rows.push(row);
    }
    return {rows,errors};
  }
  function filter(rows, filters={}){
    return rows.filter(row=>[['store','점포'],['channel','채널'],['week','기준주']].every(([key,field])=>!filters[key]||filters[key]==='all'||filters[key]==='전체'||row[field]===filters[key]));
  }
  function summarize(rows){
    const total=rows.reduce((a,r)=>({sales:a.sales+Number(r['매출_만원']),target:a.target+Number(r['목표_만원']),cost:a.cost+Number(r['비용_만원'])}),{sales:0,target:0,cost:0});
    return {...total,attainment:total.target===0?null:total.sales/total.target};
  }
  function compute(row,memo=''){
    return {...row,'목표차이_만원':row['매출_만원']-row['목표_만원'],'비용차감액_만원':row['매출_만원']-row['비용_만원'],'달성률':Number(row['목표_만원'])===0?null:row['매출_만원']/row['목표_만원'],'메모':memo};
  }
  const xmlEscape=v=>String(v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c])).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,'');
  function column(index){let result='';for(let n=index+1;n>0;n=Math.floor((n-1)/26)) result=String.fromCharCode(65+(n-1)%26)+result;return result;}
  async function createWorkbook(rows,memos={},headers=reportHeaders){
    if(!global.JSZip) throw new Error('엑셀 압축 라이브러리를 불러오지 못했습니다.');
    const zip=new global.JSZip();
    const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    let content='<row r="1">'+headers.map((h,c)=>`<c r="${column(c)}1" t="inlineStr" s="1"><is><t>${xmlEscape(h)}</t></is></c>`).join('')+'</row>';
    rows.forEach((row,i)=>{
      const values=compute(row,memos instanceof Map?(memos.get(row.record_id)??''):(memos[row.record_id]??''));
      const r=i+2;
      content+=`<row r="${r}">`+headers.map((h,c)=>{
        const ref=column(c)+r, value=values[h];
        if(value===null||value===undefined||value==='')return `<c r="${ref}"/>`;
        if(typeof value==='number'&&Number.isFinite(value)) return `<c r="${ref}"${h==='달성률'?' s="2"':''}><v>${value}</v></c>`;
        return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
      }).join('')+'</row>';
    });
    zip.file('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
    zip.file('_rels/.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('xl/workbook.xml',`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="매출보고서" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    zip.file('xl/styles.xml',`<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="${ns}"><fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><sz val="11"/><name val="맑은 고딕"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
    zip.file('xl/worksheets/sheet1.xml',`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${ns}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="16" width="19" customWidth="1"/></cols><sheetData>${content}</sheetData><autoFilter ref="A1:${column(headers.length-1)}${rows.length+1}"/></worksheet>`);
    return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',compression:'DEFLATE'});
  }
  async function readWorkbook(file){
    if(!global.JSZip)throw new Error('엑셀 압축 라이브러리를 불러오지 못했습니다.');
    const zip=await global.JSZip.loadAsync(await file.arrayBuffer());
    async function parse(path){const entry=zip.file(path);if(!entry)throw new Error(`${file.name}: 엑셀 구성 파일 누락 ${path}`);const doc=new DOMParser().parseFromString(await entry.async('string'),'application/xml');if(doc.getElementsByTagName('parsererror').length)throw new Error(`${file.name}: 손상된 엑셀 XML`);return doc;}
    const elements=(node,name)=>Array.from(node.getElementsByTagNameNS('*',name));
    const strings=zip.file('xl/sharedStrings.xml')?elements(await parse('xl/sharedStrings.xml'),'si').map(si=>elements(si,'t').map(t=>t.textContent).join('')):[];
    const wb=await parse('xl/workbook.xml'),sheets=elements(wb,'sheet');
    const sheet=sheets.find(s=>s.getAttribute('name')==='원본')??sheets[0];
    if(!sheet)throw new Error(`${file.name}: 시트가 없습니다.`);
    const id=sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    const rel=elements(await parse('xl/_rels/workbook.xml.rels'),'Relationship').find(r=>r.getAttribute('Id')===id);
    if(!rel)throw new Error(`${file.name}: 시트 연결을 찾을 수 없습니다.`);
    const target=rel.getAttribute('Target'),path=target.startsWith('/')?target.slice(1):'xl/'+target;
    const rows=elements(await parse(path),'row').map(node=>{
      const values=[];
      for(const cell of elements(node,'c')){
        const letters=(cell.getAttribute('r')??'').replace(/[0-9]/g,'');
        let col=0;for(const c of letters)col=col*26+c.charCodeAt(0)-64;
        const type=cell.getAttribute('t'),v=elements(cell,'v')[0]?.textContent??'';
        let value=type==='inlineStr'?elements(cell,'t').map(t=>t.textContent).join(''):type==='s'?(strings[Number(v)]??''):v;
        if(!['inlineStr','s','str','b','e'].includes(type)&&v!==''&&Number.isFinite(Number(v)))value=Number(v);
        values[col-1]=value;
      }
      return {number:Number(node.getAttribute('r')),values:Array.from({length:values.length},(_,i)=>values[i]??'')};
    });
    if(!rows.length)throw new Error(`${file.name}: 데이터가 없는 파일입니다.`);
    return {headers:rows[0].values,body:rows.slice(1).map(r=>r.values),sheetName:sheet.getAttribute('name'),rowNumbers:rows.slice(1).map(r=>r.number)};
  }
  global.SalesData={fields,aliases,reportHeaders,normalize,merge,filter,summarize,compute,readWorkbook,createWorkbook};
})(typeof window!=='undefined'?window:globalThis);
