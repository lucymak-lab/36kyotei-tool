/* 36協定届 作成ツール — app.js */

// ── フォーム値取得 ───────────────────────────────────────────
function f(id) { return document.getElementById(id)?.value?.trim() || ''; }

function getWorkRows() {
  return Array.from(document.querySelectorAll('.work-row')).map(row => ({
    work:       row.querySelector('.wr-work').value.trim(),
    workers:    row.querySelector('.wr-workers').value.trim(),
    dailyHours: row.querySelector('.wr-hours').value.trim(),
    reason:     row.querySelector('.wr-reason').value.trim(),
  }));
}

function formData() {
  const repType = document.querySelector('input[name="rep-type"]:checked')?.value || 'rep';
  return {
    biztype:      f('f-biztype') || '製造業',
    company:      f('f-company'),
    address:      f('f-address'),
    tel:          f('f-tel'),
    rep:          f('f-rep'),
    workRows:     getWorkRows(),
    limitDay:     f('f-limit-day'),
    limitMonth:   f('f-limit-month'),
    limitYear:    f('f-limit-year'),
    startDate:    f('f-start-date'),
    holiday:      f('f-holiday'),
    holidayStart: f('f-holiday-start'),
    holidayEnd:   f('f-holiday-end'),
    spReason:     f('f-sp-reason'),
    spMonth:      f('f-sp-month'),
    spYear:       f('f-sp-year'),
    spCount:      f('f-sp-count'),
    spRate:       f('f-sp-rate'),
    spProcedure:  f('f-sp-procedure'),
    guideMonth:   f('g-sp-month') || '100',
    guideYear:    f('g-sp-year')  || '720',
    guideAvg:     f('g-sp-avg')   || '80',
    guideCount:   f('g-sp-count') || '6',
    repType,
    unionName:    f('f-union-name'),
    repTitle:     f('f-rep-title'),
    repName:      f('f-rep-name'),
    repHow:       f('f-rep-how'),
    validFrom:    f('f-valid-from'),
    validTo:      f('f-valid-to'),
    submitDate:   f('f-submit-date'),
    office:       f('f-office'),
  };
}

// ── 業務行の追加・削除 ────────────────────────────────────────
function makeWorkRow() {
  const row = document.createElement('div');
  row.className = 'work-row';
  row.innerHTML = `
    <input type="text" class="wr-work" placeholder="例：製造業務全般" />
    <div class="wr-unit"><input type="number" class="wr-workers" placeholder="10" min="1" /><span class="unit">人</span></div>
    <div class="wr-unit"><input type="number" class="wr-hours" placeholder="8" min="1" max="8" /><span class="unit">時間</span></div>
    <input type="text" class="wr-reason" placeholder="業務量の増加・繁忙期対応等" />
    <button class="del-row-btn" title="削除">✕</button>
  `;
  row.querySelector('.del-row-btn').addEventListener('click', () => {
    if (document.querySelectorAll('.work-row').length > 1) row.remove();
  });
  return row;
}

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('workRows').appendChild(makeWorkRow());
  document.getElementById('addWorkBtn').addEventListener('click', () => {
    document.getElementById('workRows').appendChild(makeWorkRow());
  });
});

// ── テーマ切替 ───────────────────────────────────────────────
document.getElementById('themeBtn').addEventListener('click', () => {
  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  document.documentElement.setAttribute('data-theme', isDark ? '' : 'dark');
  document.getElementById('themeBtn').textContent = isDark ? '🌙' : '☀️';
});

// ── 協定当事者の種別切替 ─────────────────────────────────────
document.querySelectorAll('input[name="rep-type"]').forEach(radio => {
  radio.addEventListener('change', () => {
    const isUnion = radio.value === 'union' && radio.checked;
    document.getElementById('union-fields').style.display = isUnion ? '' : 'none';
    document.getElementById('rep-fields').style.display   = isUnion ? 'none' : '';
  });
});

// ── Word ユーティリティ ───────────────────────────────────────
function makeDoc(children) {
  const { Document, Packer } = docx;
  return new Document({ sections: [{ children }] });
}

function p(text, opts = {}) {
  const { Paragraph, TextRun, AlignmentType } = docx;
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : undefined,
    spacing:   { after: opts.after ?? 160 },
    children:  String(text || '').split('\n').flatMap((line, i) =>
      i === 0
        ? [new TextRun({ text: line, bold: !!opts.bold, size: opts.size || 22 })]
        : [new TextRun({ text: '', break: 1 }), new TextRun({ text: line, size: opts.size || 22 })]
    ),
  });
}

function blank(n = 1) {
  const { Paragraph } = docx;
  return Array.from({ length: n }, () => new Paragraph({ children: [] }));
}

function hline() {
  const { Paragraph, BorderStyle } = docx;
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'AAAAAA' } },
    spacing: { after: 120 },
    children: [],
  });
}

function tbl(rows, opts = {}) {
  const { Table, TableRow, TableCell, Paragraph, TextRun, WidthType, BorderStyle } = docx;
  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map(cols => new TableRow({
      tableHeader: !!cols._header,
      children: cols.filter(c => c !== undefined).map(({ text = '', width, bold, bg, center, size } = {}) =>
        new TableCell({
          width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
          shading: bg ? { fill: bg } : undefined,
          children: String(text || '').split('\n').map(line =>
            new Paragraph({
              alignment: center ? 'center' : undefined,
              children: [new TextRun({ text: line, bold: !!bold, size: size || 20 })],
            })
          ),
        })
      ),
    })),
  });
}

function fmtDate(str) {
  if (!str) return '　　　年　　月　　日';
  const d = new Date(str);
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
}

function fmtPeriod(from, to) {
  if (!from && !to) return '　　年　月　日　〜　　　年　月　日';
  return `${fmtDate(from)}　〜　${fmtDate(to)}`;
}

function today() { return new Date().toISOString().slice(0, 10).replace(/-/g, ''); }

// ── 協定当事者ブロック生成 ────────────────────────────────────
function repRows(d) {
  if (d.repType === 'union') {
    return [
      [{ text: '協定の当事者', width: 30, bold: true, bg: 'E8EEF7' }, { text: `労働組合の名称：${d.unionName || '　　　　　　　　'}` }],
    ];
  }
  return [
    [{ text: '協定の当事者', width: 30, bold: true, bg: 'E8EEF7' },
     { text: `職名：${d.repTitle || '　　　　'}　氏名：${d.repName || '　　　　　　'}　（選出方法：${d.repHow || '　　　　　　'}）` }],
  ];
}

// ── 様式第9号（一般条項） ─────────────────────────────────────
async function buildForm9() {
  const d = formData();

  const children = [
    ...blank(1),
    p('時間外労働・休日労働に関する協定届', { bold: true, center: true, size: 28, after: 60 }),
    p('（様式第9号）', { center: true, size: 20, after: 300 }),

    p('【事業場情報】', { bold: true, size: 22, after: 100 }),
    tbl([
      [{ text: '事業の種類', width: 25, bold: true, bg: 'E8EEF7' },
       { text: d.biztype, width: 25 },
       { text: '事業の名称', width: 25, bold: true, bg: 'E8EEF7' },
       { text: d.company }],
      [{ text: '事業の所在地', width: 25, bold: true, bg: 'E8EEF7' },
       { text: d.address },
       { text: '電話番号', width: 25, bold: true, bg: 'E8EEF7' },
       { text: d.tel }],
    ]),

    ...blank(1),
    p('【協定内容】', { bold: true, size: 22, after: 100 }),

    p('■ 業務・労働者', { bold: true, size: 20, after: 80 }),
    tbl([
      [
        { text: '業務の種類', width: 28, bold: true, bg: 'E8EEF7' },
        { text: '労働者数', width: 15, bold: true, bg: 'E8EEF7', center: true },
        { text: '所定労働時間/日', width: 17, bold: true, bg: 'E8EEF7', center: true },
        { text: '具体的事由', bold: true, bg: 'E8EEF7' },
      ],
      ...d.workRows.map(r => [
        { text: r.work },
        { text: r.workers ? `${r.workers}人` : '', center: true },
        { text: r.dailyHours ? `${r.dailyHours}時間` : '', center: true },
        { text: r.reason },
      ]),
    ]),

    ...blank(1),
    p('■ 延長できる時間数', { bold: true, size: 20, after: 80 }),
    tbl([
      [
        { text: '1日', width: 20, bold: true, bg: 'E8EEF7', center: true },
        { text: '1か月', width: 20, bold: true, bg: 'E8EEF7', center: true },
        { text: '1年間', width: 20, bold: true, bg: 'E8EEF7', center: true },
        { text: '起算日（1年単位）', width: 40, bold: true, bg: 'E8EEF7', center: true },
      ],
      [
        { text: d.limitDay ? `${d.limitDay}時間` : '', center: true },
        { text: d.limitMonth ? `${d.limitMonth}時間` : '', center: true },
        { text: d.limitYear ? `${d.limitYear}時間` : '', center: true },
        { text: fmtDate(d.startDate), center: true },
      ],
    ]),

    ...blank(1),
    p('■ 休日労働', { bold: true, size: 20, after: 80 }),
    tbl([
      [{ text: '対象となる休日', width: 40, bold: true, bg: 'E8EEF7' },
       { text: '始業時刻', width: 20, bold: true, bg: 'E8EEF7', center: true },
       { text: '終業時刻', width: 20, bold: true, bg: 'E8EEF7', center: true }],
      [{ text: d.holiday },
       { text: d.holidayStart, center: true },
       { text: d.holidayEnd, center: true }],
    ]),

    ...blank(1),
    p('【協定の有効期間・当事者】', { bold: true, size: 22, after: 100 }),
    tbl([
      [{ text: '協定の有効期間', width: 30, bold: true, bg: 'E8EEF7' },
       { text: fmtPeriod(d.validFrom, d.validTo) }],
      ...repRows(d),
    ]),

    ...blank(2),
    hline(),
    p(`${fmtDate(d.submitDate)}`, { after: 160 }),
    p(`${d.office || '　　　　　　　'}労働基準監督署長　殿`, { after: 300 }),
    p(`使用者職氏名：${d.rep || '　　　　　　　　　　'}　　（印）`, { after: 120 }),
  ];

  return makeDoc(children);
}

// ── 様式第9号の2（特別条項） ──────────────────────────────────
async function buildForm9no2() {
  const d = formData();

  const children = [
    ...blank(1),
    p('時間外労働・休日労働に関する協定届（特別条項）', { bold: true, center: true, size: 28, after: 60 }),
    p('（様式第9号の2）', { center: true, size: 20, after: 300 }),

    p('【事業場情報】', { bold: true, size: 22, after: 100 }),
    tbl([
      [{ text: '事業の種類', width: 25, bold: true, bg: 'E8EEF7' },
       { text: d.biztype, width: 25 },
       { text: '事業の名称', width: 25, bold: true, bg: 'E8EEF7' },
       { text: d.company }],
      [{ text: '事業の所在地', width: 25, bold: true, bg: 'E8EEF7' },
       { text: d.address },
       { text: '電話番号', width: 25, bold: true, bg: 'E8EEF7' },
       { text: d.tel }],
    ]),

    ...blank(1),
    p('【一般条項】', { bold: true, size: 22, after: 100 }),

    p('■ 業務・労働者', { bold: true, size: 20, after: 80 }),
    tbl([
      [
        { text: '業務の種類', width: 28, bold: true, bg: 'E8EEF7' },
        { text: '労働者数', width: 15, bold: true, bg: 'E8EEF7', center: true },
        { text: '所定労働時間/日', width: 17, bold: true, bg: 'E8EEF7', center: true },
        { text: '具体的事由', bold: true, bg: 'E8EEF7' },
      ],
      ...d.workRows.map(r => [
        { text: r.work },
        { text: r.workers ? `${r.workers}人` : '', center: true },
        { text: r.dailyHours ? `${r.dailyHours}時間` : '', center: true },
        { text: r.reason },
      ]),
    ]),

    ...blank(1),
    p('■ 延長できる時間数（限度時間）', { bold: true, size: 20, after: 80 }),
    tbl([
      [
        { text: '1日', width: 20, bold: true, bg: 'E8EEF7', center: true },
        { text: '1か月', width: 20, bold: true, bg: 'E8EEF7', center: true },
        { text: '1年間', width: 20, bold: true, bg: 'E8EEF7', center: true },
        { text: '起算日（1年単位）', width: 40, bold: true, bg: 'E8EEF7', center: true },
      ],
      [
        { text: d.limitDay ? `${d.limitDay}時間` : '', center: true },
        { text: d.limitMonth ? `${d.limitMonth}時間` : '', center: true },
        { text: d.limitYear ? `${d.limitYear}時間` : '', center: true },
        { text: fmtDate(d.startDate), center: true },
      ],
    ]),

    ...blank(1),
    p('■ 休日労働', { bold: true, size: 20, after: 80 }),
    tbl([
      [{ text: '対象となる休日', width: 40, bold: true, bg: 'E8EEF7' },
       { text: '始業時刻', width: 20, bold: true, bg: 'E8EEF7', center: true },
       { text: '終業時刻', width: 20, bold: true, bg: 'E8EEF7', center: true }],
      [{ text: d.holiday },
       { text: d.holidayStart, center: true },
       { text: d.holidayEnd, center: true }],
    ]),

    ...blank(1),
    p('【特別条項】', { bold: true, size: 22, after: 100 }),
    tbl([
      [{ text: '臨時的な特別の事情', width: 30, bold: true, bg: 'FFF0D0' },
       { text: d.spReason }],
      [{ text: '限度時間を超えて延長できる時間数（1か月）', width: 30, bold: true, bg: 'FFF0D0' },
       { text: d.spMonth ? `${d.spMonth}時間未満` : '' }],
      [{ text: '限度時間を超えて延長できる時間数（1年間）', width: 30, bold: true, bg: 'FFF0D0' },
       { text: d.spYear ? `${d.spYear}時間以下` : '' }],
      [{ text: '社内上限（1か月）', width: 30, bold: true, bg: 'FFF0D0' },
       { text: `${d.guideMonth}時間未満` }],
      [{ text: '社内上限（1年間）', width: 30, bold: true, bg: 'FFF0D0' },
       { text: `${d.guideYear}時間以下` }],
      [{ text: '社内上限（複数月平均）', width: 30, bold: true, bg: 'FFF0D0' },
       { text: `${d.guideAvg}時間以下` }],
      [{ text: '発動回数上限（1年）', width: 30, bold: true, bg: 'FFF0D0' },
       { text: `${d.guideCount}回以内` }],
      [{ text: '限度時間を超える回数（1年）', width: 30, bold: true, bg: 'FFF0D0' },
       { text: d.spCount ? `${d.spCount}回以内` : '' }],
      [{ text: '割増賃金率', width: 30, bold: true, bg: 'FFF0D0' },
       { text: d.spRate ? `${d.spRate}%以上` : '' }],
      [{ text: '限度時間を超えた場合の手続き', width: 30, bold: true, bg: 'FFF0D0' },
       { text: d.spProcedure }],
    ]),

    ...blank(1),
    p('【協定の有効期間・当事者】', { bold: true, size: 22, after: 100 }),
    tbl([
      [{ text: '協定の有効期間', width: 30, bold: true, bg: 'E8EEF7' },
       { text: fmtPeriod(d.validFrom, d.validTo) }],
      ...repRows(d),
    ]),

    ...blank(2),
    hline(),
    p(`${fmtDate(d.submitDate)}`, { after: 160 }),
    p(`${d.office || '　　　　　　　'}労働基準監督署長　殿`, { after: 300 }),
    p(`使用者職氏名：${d.rep || '　　　　　　　　　　'}　　（印）`, { after: 120 }),
  ];

  return makeDoc(children);
}

// ── ダウンロードヘルパー ─────────────────────────────────────
async function download(doc, filename) {
  const blob = await docx.Packer.toBlob(doc);
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

// ── JSON 保存・読み込み ──────────────────────────────────────
function saveJSON() {
  const blob = new Blob([JSON.stringify(formData(), null, 2)], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = `36協定_${today()}.json`; a.click();
  URL.revokeObjectURL(url);
}

function loadJSON(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const d = JSON.parse(e.target.result);
      const set = (id, val) => { const el = document.getElementById(id); if (el && val != null) el.value = val; };
      set('f-biztype',       d.biztype);
      set('f-company',       d.company);
      set('f-address',       d.address);
      set('f-tel',           d.tel);
      set('f-rep',           d.rep);
      set('f-limit-day',     d.limitDay);
      set('f-limit-month',   d.limitMonth);
      set('f-limit-year',    d.limitYear);
      set('f-start-date',    d.startDate);
      set('f-holiday',       d.holiday);
      set('f-holiday-start', d.holidayStart);
      set('f-holiday-end',   d.holidayEnd);
      set('f-sp-reason',     d.spReason);
      set('f-sp-month',      d.spMonth);
      set('f-sp-year',       d.spYear);
      set('f-sp-count',      d.spCount);
      set('f-sp-rate',       d.spRate);
      set('f-sp-procedure',  d.spProcedure);
      set('g-sp-month',      d.guideMonth);
      set('g-sp-year',       d.guideYear);
      set('g-sp-avg',        d.guideAvg);
      set('g-sp-count',      d.guideCount);
      set('f-union-name',    d.unionName);
      set('f-rep-title',     d.repTitle);
      set('f-rep-name',      d.repName);
      set('f-rep-how',       d.repHow);
      set('f-valid-from',    d.validFrom);
      set('f-valid-to',      d.validTo);
      set('f-submit-date',   d.submitDate);
      set('f-office',        d.office);
      if (d.repType) {
        const radio = document.querySelector(`input[name="rep-type"][value="${d.repType}"]`);
        if (radio) { radio.checked = true; radio.dispatchEvent(new Event('change')); }
      }
      if (Array.isArray(d.workRows)) {
        const container = document.getElementById('workRows');
        container.innerHTML = '';
        (d.workRows.length ? d.workRows : [{}]).forEach(r => {
          const row = makeWorkRow();
          row.querySelector('.wr-work').value    = r.work       || '';
          row.querySelector('.wr-workers').value = r.workers    || '';
          row.querySelector('.wr-hours').value   = r.dailyHours || '';
          row.querySelector('.wr-reason').value  = r.reason     || '';
          container.appendChild(row);
        });
      }
    } catch { alert('JSONファイルの読み込みに失敗しました。'); }
  };
  reader.readAsText(file);
}

// ── エクスポートボタン ────────────────────────────────────────
document.getElementById('btnForm9').addEventListener('click', async () => {
  download(await buildForm9(), `36協定届_様式第9号_${today()}.docx`);
});

document.getElementById('btnForm9no2').addEventListener('click', async () => {
  download(await buildForm9no2(), `36協定届_様式第9号の2_${today()}.docx`);
});

document.getElementById('btnZip').addEventListener('click', async () => {
  const btn = document.getElementById('btnZip');
  btn.disabled = true; btn.textContent = '⏳ 生成中...';
  try {
    const zip = new JSZip();
    const toBlob = async doc => docx.Packer.toBlob(doc);
    zip.file('①36協定届_様式第9号.docx',      await toBlob(await buildForm9()));
    zip.file('②36協定届_様式第9号の2.docx',    await toBlob(await buildForm9no2()));
    const blob = await zip.generateAsync({ type: 'blob' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url;
    a.download = `36協定届一式_${today()}.zip`;
    a.click();
    URL.revokeObjectURL(url);
  } finally {
    btn.disabled = false; btn.textContent = '⬇ ZIP一括';
  }
});
