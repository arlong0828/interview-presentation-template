const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const input = process.env.PRESENTATION_PROFILE || 'profile.json';
const output = process.env.PRESENTATION_OUTPUT || 'interview-presentation.pptx';
if (!fs.existsSync(input)) {
  throw new Error(`Missing ${input}. Copy profile.example.json to profile.json and edit it.`);
}
const data = JSON.parse(fs.readFileSync(input, 'utf8'));
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = data.candidate.name;
pptx.subject = 'Interview presentation';
pptx.title = `${data.candidate.name} — ${data.candidate.title}`;
pptx.company = '';
pptx.lang = 'zh-TW';
pptx.theme = {
  headFontFace: 'Microsoft JhengHei',
  bodyFontFace: 'Microsoft JhengHei',
  lang: 'zh-TW',
};

const C = { navy: '14213D', gold: 'FCA311', paper: 'F7F7F5', ink: '202124', white: 'FFFFFF', muted: '64748B' };
function base(slide, section) {
  slide.background = { color: C.paper };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.18, fill: { color: C.gold }, line: { color: C.gold } });
  slide.addText(section, { x: 0.65, y: 7.05, w: 12, h: 0.2, fontSize: 9, color: C.muted, margin: 0 });
}
function title(slide, text, subtitle = '') {
  slide.addText(text, { x: 0.7, y: 0.55, w: 11.9, h: 0.65, fontSize: 30, bold: true, color: C.navy, margin: 0 });
  if (subtitle) slide.addText(subtitle, { x: 0.72, y: 1.28, w: 11.6, h: 0.4, fontSize: 15, color: C.muted, margin: 0 });
}
function bullets(slide, items, x, y, w, h) {
  slide.addText(items.map(text => ({ text, options: { bullet: { indent: 18 }, hanging: 4, breakLine: true } })), {
    x, y, w, h, fontSize: 20, color: C.ink, breakLine: false, paraSpaceAfterPt: 12, margin: 0.08,
  });
}

{
  const s = pptx.addSlide(); base(s, 'INTRODUCTION');
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0.18, w: 13.333, h: 6.82, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText(data.candidate.name, { x: 0.8, y: 1.65, w: 11.7, h: 0.8, fontSize: 42, bold: true, color: C.white, align: 'center', margin: 0 });
  s.addText(data.candidate.title, { x: 1.2, y: 2.6, w: 10.9, h: 0.55, fontSize: 24, color: C.gold, align: 'center', margin: 0 });
  s.addText(data.candidate.summary, { x: 1.45, y: 3.5, w: 10.4, h: 1.25, fontSize: 19, color: C.white, align: 'center', valign: 'mid', margin: 0.05 });
  s.addText(data.candidate.contact, { x: 1.2, y: 5.7, w: 10.9, h: 0.35, fontSize: 14, color: 'CBD5E1', align: 'center', margin: 0 });
}
{
  const s = pptx.addSlide(); base(s, 'CORE STRENGTHS'); title(s, 'Core strengths');
  bullets(s, data.strengths, 1.05, 1.9, 11.2, 4.5);
}
{
  const s = pptx.addSlide(); base(s, 'EXPERIENCE'); title(s, 'Experience');
  let y = 1.75;
  for (const item of data.experience) {
    s.addText(`${item.role} · ${item.company}`, { x: 0.8, y, w: 8.7, h: 0.42, fontSize: 20, bold: true, color: C.navy, margin: 0 });
    s.addText(item.period, { x: 9.8, y, w: 2.65, h: 0.35, fontSize: 13, color: C.muted, align: 'right', margin: 0 });
    bullets(s, item.highlights, 1.05, y + 0.55, 11.2, 1.25);
    y += 1.85;
  }
}
for (const project of data.projects) {
  const s = pptx.addSlide(); base(s, 'PROJECT'); title(s, project.name, 'Problem → Action → Result');
  const cards = [['Problem', project.problem], ['Action', project.action], ['Result', project.result]];
  cards.forEach(([label, text], i) => {
    const x = 0.65 + i * 4.18;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 2, w: 3.75, h: 3.7, rectRadius: 0.08, fill: { color: C.white }, line: { color: i === 2 ? C.gold : 'D7DCE2', width: 1.2 } });
    s.addText(label, { x: x + 0.28, y: 2.3, w: 3.15, h: 0.45, fontSize: 20, bold: true, color: C.navy, margin: 0 });
    s.addText(text, { x: x + 0.28, y: 3.0, w: 3.15, h: 2.2, fontSize: 16, color: C.ink, valign: 'mid', margin: 0.04, breakLine: false });
  });
}
{
  const s = pptx.addSlide(); base(s, 'CLOSING');
  s.addText(data.closing || 'Thank you', { x: 0.8, y: 2.45, w: 11.7, h: 1, fontSize: 44, bold: true, color: C.navy, align: 'center', margin: 0 });
  s.addText(data.candidate.contact, { x: 1.2, y: 4, w: 10.9, h: 0.45, fontSize: 18, color: C.muted, align: 'center', margin: 0 });
}

pptx.writeFile({ fileName: output });
