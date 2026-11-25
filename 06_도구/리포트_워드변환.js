// 사용법: node 리포트_워드변환.js <입력.md> <출력.docx>
// 일일 리포트(마크다운)를 A4 가로 워드 문서로 변환한다. 표·제목·인용·목록·링크·굵게를 지원한다.
const fs = require('fs');
const D = require('docx');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, WidthType,
  ShadingType, BorderStyle, LevelFormat, ExternalHyperlink, PageOrientation, AlignmentType, Footer, PageNumber } = D;
const [,, inp, out] = process.argv;
if (!inp || !out) { console.error('사용법: node 리포트_워드변환.js 입력.md 출력.docx'); process.exit(1); }
const FONT = '맑은 고딕';
const lines = fs.readFileSync(inp, 'utf8').replace(/\r/g, '').split('\n');
const W = 16838 - 2 * 1134; // A4 가로 폭 - 여백(2cm)
const bd = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const borders = { top: bd, bottom: bd, left: bd, right: bd };

function runs(text, base = {}) {
  text = text.replace(/🆕/g, '[신규]').replace(/⏰/g, '[임박]').replace(/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE0F}]/gu, '');
  const out = [];
  const re = /(\[([^\]]+)\]\((https?:[^)]+)\))|(\*\*([^*]+)\*\*)|(`([^`]+)`)/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, ...base }));
    if (m[1]) out.push(new ExternalHyperlink({ link: m[3], children: [new TextRun({ text: m[2], font: FONT, style: 'Hyperlink', ...base })] }));
    else if (m[4]) out.push(new TextRun({ text: m[5], bold: true, font: FONT, ...base }));
    else out.push(new TextRun({ text: m[7], font: 'Consolas', ...base }));
    last = re.lastIndex;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, ...base }));
  return out;
}
const isSep = l => /^\|?\s*:?-{2,}/.test(l.trim()) && l.includes('-') && !/[가-힣A-Za-z0-9]/.test(l);
const cells = l => l.trim().replace(/^\||\|$/g, '').split('|').map(s => s.trim());

function table(rows) {
  const n = rows[0].length;
  // 열 폭: 글자 수에 비례(최소 비율 보장)
  const len = Array.from({ length: n }, (_, i) => Math.max(9, Math.min(40, Math.max(...rows.map(r => (r[i] || '').length)))));
  const sum = len.reduce((a, b) => a + b, 0);
  let cw = len.map(x => Math.floor(W * x / sum));
  cw[n - 1] += W - cw.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: cw,
    rows: rows.map((r, ri) => new TableRow({
      tableHeader: ri === 0, cantSplit: true,
      children: cw.map((w, ci) => new TableCell({
        width: { size: w, type: WidthType.DXA }, borders,
        margins: { top: 60, bottom: 60, left: 90, right: 90 },
        shading: ri === 0 ? { fill: 'DCE6F1', type: ShadingType.CLEAR, color: 'auto' } : undefined,
        children: [new Paragraph({ children: runs(r[ci] || '', { size: 18, bold: ri === 0 }) })]
      }))
    }))
  });
}

const kids = [];
let i = 0;
while (i < lines.length) {
  const l = lines[i];
  if (!l.trim()) { i++; continue; }
  if (l.startsWith('|')) {
    const block = [];
    while (i < lines.length && lines[i].startsWith('|')) { if (!isSep(lines[i])) block.push(cells(lines[i])); i++; }
    const n = block[0].length;
    kids.push(table(block.map(r => { while (r.length < n) r.push(''); return r.slice(0, n); })));
    kids.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
    continue;
  }
  if (l.startsWith('# ')) kids.push(new Paragraph({ heading: HeadingLevel.TITLE, children: runs(l.slice(2)) }));
  else if (l.startsWith('## ')) kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs(l.slice(3)) }));
  else if (l.startsWith('### ')) kids.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs(l.slice(4)) }));
  else if (l.startsWith('>')) kids.push(new Paragraph({
    spacing: { after: 60 }, indent: { left: 200 },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: '4F81BD', space: 8 } },
    children: runs(l.replace(/^>\s?/, ''), { size: 20, color: '444444' }) }));
  else if (/^\s*[-*] /.test(l)) kids.push(new Paragraph({ numbering: { reference: 'b', level: /^\s{2,}/.test(l) ? 1 : 0 }, spacing: { after: 40 }, children: runs(l.replace(/^\s*[-*] /, ''), { size: 21 }) }));
  else kids.push(new Paragraph({ spacing: { after: 80 }, children: runs(l, { size: 21 }) }));
  i++;
}

const doc = new Document({
  creator: '지원사업 서칭 에이전트', title: '지원사업 일일 리포트',
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', run: { font: FONT, size: 36, bold: true, color: '1F3864' }, paragraph: { spacing: { before: 0, after: 160 }, keepNext: true } },
      { id: 'Heading1', name: 'heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 28, bold: true, color: '1F3864' }, paragraph: { spacing: { before: 280, after: 120 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 24, bold: true, color: '2F5496' }, paragraph: { spacing: { before: 200, after: 80 }, keepNext: true, outlineLevel: 1 } },
    ]
  },
  numbering: { config: [{ reference: 'b', levels: [
    { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } },
    { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1000, hanging: 270 } } } } ] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18 })] })] }) },
    children: kids
  }]
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(out, b); console.log('저장:', out); });
