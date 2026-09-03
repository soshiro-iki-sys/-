const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType,
  VerticalAlign, PageBreak, PageOrientation,
} = require('docx');

const FONT = 'MS PGothic';
const COLS = [800, 5238, 1600, 2000];                 // 合計 9638 DXA
const TOTAL = COLS.reduce((a, b) => a + b, 0);

const morning = [
  { name: '荘司　和弘',        n: '1名' },
  { name: '遠藤　清一',        n: '1名' },
  { name: '福島　正',          n: '1名' },
  { name: '廣川　克仁',        n: '1名' },
  { name: 'タケダ　ツヨシ',    n: '1名' },
  { name: '廣川　克仁',        n: '1名' },
  { name: '平野　和則',        n: '1名' },
  { name: '村上　聖一',        n: '1名' },
];

const afternoon = [
  { name: '棚村　功',          n: '1名' },
  { name: '高石　俊一',        n: '1名' },
  { name: '永井　清',          n: '1名' },
  { name: 'タカノ　ノブヤス',  n: '1名' },
];

const run = (text, o = {}) => new TextRun({
  text, font: FONT, size: o.size || 21, bold: !!o.bold, color: o.color || '000000',
});

const p = (text, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.LEFT,
  spacing: { before: o.before || 0, after: o.after || 0, line: 260 },
  children: [run(text, o)],
});

const cell = (text, i, o = {}) => new TableCell({
  width: { size: COLS[i], type: WidthType.DXA },
  margins: { top: 90, bottom: 90, left: 100, right: 100 },
  verticalAlign: VerticalAlign.CENTER,
  shading: o.shade ? { type: ShadingType.CLEAR, fill: o.shade, color: 'auto' } : undefined,
  children: [p(text, { align: o.align || AlignmentType.CENTER, bold: o.bold, size: o.size, color: o.color })],
});

function sheet(title, time, rows, blanks) {
  const head = ['No.', 'お名前', '人数', '出席'];

  const tableRows = [
    new TableRow({
      tableHeader: true,
      children: head.map((h, i) => cell(h, i, { bold: true, shade: 'E8EEF7' })),
    }),
  ];

  rows.forEach((r, idx) => {
    tableRows.push(new TableRow({
      children: [
        cell(String(idx + 1), 0),
        cell(r.name, 1, { align: AlignmentType.LEFT, size: 22 }),
        cell(r.n, 2),
        cell('□', 3, { size: 28 }),
      ],
    }));
  });

  for (let i = 0; i < blanks; i++) {
    tableRows.push(new TableRow({
      children: [
        cell(String(rows.length + i + 1), 0, { shade: 'FAFAFA' }),
        cell('', 1, { shade: 'FAFAFA' }),
        cell('', 2, { shade: 'FAFAFA' }),
        cell('□', 3, { size: 28, shade: 'FAFAFA' }),
      ],
    }));
  }

  const border = { style: BorderStyle.SINGLE, size: 4, color: '7A7A7A' };

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [run('出席確認シート', { bold: true, size: 32 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 260 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '2E5C8A', space: 6 } },
      children: [run('太陽光・蓄電池セミナー', { size: 22 })],
    }),
    p('開催日：2026年　　　月　　　日（　　）　　　会場：　　　　　　　　　　　　　　　　', { size: 20, after: 200 }),
    new Paragraph({
      spacing: { after: 140 },
      children: [run(`【${title}】　${time}`, { bold: true, size: 26 })],
    }),
    new Table({
      columnWidths: COLS,
      width: { size: TOTAL, type: WidthType.DXA },
      borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
      rows: tableRows,
    }),
    p(`申込者数：${rows.length}名　／　当日出席者数：　　　　名　（うち飛び入り参加：　　　　名）`, { size: 20, before: 260 }),
    p('受付担当者：　　　　　　　　　　　　　　　', { size: 20, before: 160 }),
  ];
}

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  sections: [{
    properties: {
      page: {
        size: { orientation: PageOrientation.PORTRAIT },
        margin: { top: 1000, bottom: 1000, left: 1134, right: 1134 },
      },
    },
    children: [
      ...sheet('午前の部', '受付 10:00～ ／ 講座 10:30～12:00', morning, 5),
      new Paragraph({ children: [new PageBreak()] }),
      ...sheet('午後の部', '受付 13:00～ ／ 講座 13:30～15:00', afternoon, 5),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('出席確認シート.docx', b);
  console.log('written');
});
