/* All data is educational sample data. Memos remain in this browser only. */
(() => {
  'use strict';
  const D = window.SalesData;
  const $ = (id) => document.getElementById(id);
  const number = new Intl.NumberFormat('ko-KR', { maximumFractionDigits: 2 });
  const storageKey = 'momentroom-sales-memos-v1';
  let rows = [], filtered = [], reportHeaders = [], memos = Object.create(null), ready = false;
  const esc = (value) => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fmt = (value) => number.format(Number(value) || 0);
  const rate = (value) => value == null || !Number.isFinite(value) ? '—' : `${(value * 100).toFixed(1)}%`;
  function notify(message, error = false) { $('message').textContent = message; $('message').classList.toggle('error', error); }
  function errorText(error) { return typeof error === 'string' ? error : error.message || JSON.stringify(error); }
  function populateFilters() {
    for (const [id, key, title] of [['store','점포','전체 점포'],['channel','채널','전체 채널'],['week','기준주','전체 기준주']]) {
      const selected = $(id).value;
      $(id).innerHTML = `<option value="">${title}</option>` + [...new Set(rows.map(r => String(r[key])))].sort((a,b) => a.localeCompare(b,'ko')).map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
      if ([...$(id).options].some(o => o.value === selected)) $(id).value = selected;
    }
  }
  function renderChart() {
    const stores = [...new Set(filtered.map(r => r['점포']))].sort((a,b) => a.localeCompare(b,'ko'));
    if (!stores.length) { $('chart').innerHTML = '<p class="empty">선택 조건에 해당하는 실적이 없습니다.</p>'; return; }
    const groups = stores.map(store => ({ store, ...D.summarize(filtered.filter(r => r['점포'] === store)) }));
    const max = Math.max(1, ...groups.flatMap(g => [g.sales,g.target]));
    $('chart').innerHTML = groups.map(g => `<div class="chart-row"><span class="store-name">${esc(g.store)}</span><div class="bar-pair" role="img" aria-label="${esc(g.store)} 매출 ${fmt(g.sales)}만원, 목표 ${fmt(g.target)}만원">${[['sales',g.sales,'매출'],['target',g.target,'목표']].map(([kind,value,label]) => `<div class="bar-line" style="--bar-width:${Math.max(0,value/max*100)}%"><div class="bar ${kind}" style="width:${Math.max(0,value/max*100)}%"></div><span class="bar-label">${fmt(value)} <span class="sr-only">${label} 만원</span></span></div>`).join('')}</div></div>`).join('') + `<div class="chart-axis"><span>0</span><span>${fmt(max/2)}</span><span>${fmt(max)} 만원</span></div>`;
  }
  function render() {
    filtered = D.filter(rows, { store:$('store').value, channel:$('channel').value, week:$('week').value });
    const sum = D.summarize(filtered);
    $('sales-total').textContent = fmt(sum.sales); $('target-total').textContent = fmt(sum.target); $('attainment-total').textContent = rate(sum.attainment);
    $('row-count').textContent = `${fmt(filtered.length)}건`; $('loaded-count').textContent = `전체 ${fmt(rows.length)}건 중 ${fmt(filtered.length)}건`;
    $('export').disabled = !ready || !filtered.length;
    renderChart();
    $('records').innerHTML = filtered.length ? filtered.map(row => {
      const c = D.compute(row,memos[row.record_id] || '');
      return `<tr data-record-id="${esc(row.record_id)}"><td class="id-cell">${esc(row.record_id)}</td><td><span class="channel-tag">${esc(row['채널'])}</span></td><td>${esc(row['기준주'])}</td><td>${esc(row['점포'])}</td><td>${esc(row['브랜드'])}</td><td>${esc(row['장르'])}</td><td class="num">${fmt(row['매출_만원'])}</td><td class="num">${fmt(row['목표_만원'])}</td><td class="num">${fmt(c['목표차이_만원'])}</td><td class="num">${rate(c['달성률'])}</td><td class="num">${fmt(row['비용_만원'])}</td><td class="num">${fmt(c['비용차감액_만원'])}</td><td>${esc(row['상태'])}</td><td><textarea class="memo-input" rows="1" data-memo-id="${esc(row.record_id)}" aria-label="${esc(row.record_id)} 메모" placeholder="메모 입력">${esc(memos[row.record_id] || '')}</textarea></td><td>${esc(row['출처'])}</td><td>${esc(row['기준일'])}</td><td class="provenance">${esc(row._sourceFile || row['출처'])} · ${esc(row._sourceSheet || '원본')} · ${esc(row._sourceRow ?? '—')}행</td></tr>`;
    }).join('') : '<tr><td colspan="17" class="empty">선택 조건에 해당하는 실적이 없습니다. 필터를 변경해 주세요.</td></tr>';
  }
  $('records').addEventListener('input', (event) => {
    const id = event.target.dataset.memoId;
    if (!id) return;
    memos[id] = event.target.value;
    try { localStorage.setItem(storageKey, JSON.stringify(memos)); $('memo-status').textContent = `${id} 메모를 이 브라우저에 저장했습니다. 새로고침 후에도 유지됩니다.`; }
    catch { $('memo-status').textContent = `${id} 메모를 저장하지 못했습니다. 브라우저 저장 공간·권한을 확인해 주세요. 현재 입력은 엑셀 내보내기에 반영됩니다.`; }
  });
  ['store','channel','week'].forEach(id => $(id).addEventListener('change',render));
  $('reset').addEventListener('click', () => { ['store','channel','week'].forEach(id => $(id).value = ''); render(); });
  function switchView(operations) { $('sales-view').hidden = operations; $('operations-view').hidden = !operations; $('sales-nav').classList.toggle('active',!operations); $('operations-nav').classList.toggle('active',operations); $('sales-nav').setAttribute('aria-current',operations ? 'false' : 'page'); $('operations-nav').setAttribute('aria-current',operations ? 'page' : 'false'); }
  $('sales-nav').addEventListener('click',() => switchView(false)); $('operations-nav').addEventListener('click',() => switchView(true)); $('back-sales').addEventListener('click',() => switchView(false));
  $('upload-trigger').addEventListener('click',() => $('upload').click());
  $('upload').addEventListener('change',async (event) => {
    const files = [...event.target.files]; if (!files.length) return;
    const messages = []; let issues = false;
    for (const file of files) {
      try {
        const workbook = await D.readWorkbook(file);
        const parsed = D.normalize(workbook.headers,workbook.body,file.name,workbook.sheetName,workbook.rowNumbers);
        const merge = D.merge(rows, parsed.rows || []);
        const added = merge.rows.length - rows.length;
        rows = merge.rows;
        messages.push(`${file.name}: ${added}건 추가`);
        const errors = [...(parsed.errors || []),...(merge.errors || [])];
        if (errors.length) { issues = true; messages.push(...errors.map(errorText)); }
      } catch (error) { issues = true; messages.push(`${file.name}: ${error.message || error}`); }
    }
    populateFilters(); render(); notify(`${messages.join('\n')}\n현재 전체 ${fmt(rows.length)}건. 먼저 불러온 정상 자료는 유지됩니다.`,issues); event.target.value = '';
  });
  $('export').addEventListener('click',async () => {
    $('export').disabled = true;
    try {
      const blob = await D.createWorkbook(filtered,memos,reportHeaders);
      const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url;
      const filterName = [$('store').value,$('channel').value,$('week').value].filter(Boolean).join('_') || '전체';
      a.download = `모먼트룸_매출보고_${filterName}.xlsx`; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url),60000);
      notify(`선택한 ${fmt(filtered.length)}건의 엑셀 보고서를 새 파일로 내보냈습니다. 메모와 계산값을 포함하며 04 보고양식의 열 순서를 따릅니다. 파일은 브라우저의 다운로드 위치에서 확인해 주세요.`);
    } catch (error) { notify(`보고서 내보내기 실패: ${error.message || error}`,true); }
    finally { $('export').disabled = !filtered.length; }
  });
  async function init() {
    try { const saved = JSON.parse(localStorage.getItem(storageKey) || '{}'); if (saved && typeof saved === 'object' && !Array.isArray(saved)) for (const [id,value] of Object.entries(saved)) if (typeof value === 'string') memos[id] = value; }
    catch { $('memo-status').textContent = '저장된 메모를 읽지 못했습니다. 새 메모 저장 시 브라우저 권한을 확인해 주세요.'; }
    try {
      if (!D) throw new Error('데이터 처리 모듈을 불러오지 못했습니다.');
      const response = await fetch('data/sales.json'); if (!response.ok) throw new Error(`기본 데이터 응답 오류 (${response.status})`);
      const data = await response.json(); const merged = D.merge([],data.rows); rows = merged.rows; reportHeaders = data.reportHeaders; ready = true;
      populateFilters(); render();
      $('sources').innerHTML = Array.isArray(data.sources) ? data.sources.map(source => typeof source === 'string' ? `<p>${esc(source)}</p>` : `<p><strong>${esc(source.filename)}</strong> · ${esc(source.sheet || '원본')} 시트 · ${fmt(source.rows)}행<br>원본 열: ${esc((source.headers || []).join(' · '))}</p>`).join('') : `<p>${esc(JSON.stringify(data.sources || {}))}</p>`;
      notify(`교육용 가상 실적 ${fmt(rows.length)}건을 불러왔습니다. 3개 채널의 자료를 열 이름으로 통합한 사본입니다.${merged.errors?.length ? '\n' + merged.errors.map(errorText).join('\n') : ''}`,!!merged.errors?.length);
    } catch (error) { notify(`자료를 불러오지 못했습니다: ${error.message}. 로컬 실행 서버에서 열었는지 확인해 주세요.`,true); }
  }
  init();
})();
