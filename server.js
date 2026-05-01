const express = require('express');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak, Header, Footer } = require('docx');

const app = express();
app.use(express.json());

const C = {
  navy: '1A2B5F', navyDk: '0F1A3D', green: '2ECC71', greenDk: '27AE60',
  red: 'E74C3C', orange: 'F39C12', gray: 'F8F9FA', grayMid: 'ECF0F1',
  grayTxt: '7F8C8D', white: 'FFFFFF', black: '2C3E50', blue: '3498DB',
};

const bdr = (color='CCCCCC', size=6) => ({ style: BorderStyle.SINGLE, size, color });
const noBdr = () => ({ style: BorderStyle.NONE, size: 0, color: 'FFFFFF' });
const allBorders = (color='CCCCCC') => ({ top: bdr(color), bottom: bdr(color), left: bdr(color), right: bdr(color) });
const noBorders = () => ({ top: noBdr(), bottom: noBdr(), left: noBdr(), right: noBdr() });

function t(text, opts={}) {
  return new TextRun({ text: String(text), font: 'Arial', size: (opts.size||11)*2, bold: opts.bold||false, italics: opts.italic||false, color: opts.color||C.black });
}

function p(children, opts={}) {
  if (!Array.isArray(children)) children = [children];
  return new Paragraph({ alignment: opts.align||AlignmentType.LEFT, spacing: { before: opts.before||0, after: opts.after||80 }, children });
}

function h(text, size=14, color=C.navy) {
  return p([t(text, { bold: true, size, color })], { before: 200, after: 120 });
}

function divider(color=C.navy) {
  return new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 1 } }, spacing: { before: 100, after: 100 }, children: [t('')] });
}

function sp() { return p([t('')]); }

function bar(pct, width=30) {
  const filled = Math.round(Math.min(100,pct)/100*width);
  return '\u2588'.repeat(filled) + '\u2591'.repeat(width-filled);
}

function miniBar(label, value, max, unit='') {
  const pct = Math.min(100,(value/max)*100);
  const color = pct>=70 ? C.green : pct>=40 ? C.orange : C.red;
  return p([
    t(label.padEnd(20), { size:10, color:C.grayTxt }),
    t(bar(pct,25), { size:10, color, font:'Courier New' }),
    t(' '+value+unit, { size:10, bold:true, color:C.navy }),
  ]);
}

function cell(text, fill=C.white, textColor=C.black, width=3120, align=AlignmentType.CENTER, bold=false) {
  return new TableCell({
    width: { size:width, type:WidthType.DXA },
    shading: { fill, type:ShadingType.CLEAR },
    borders: allBorders(fill===C.white?'DDDDDD':fill),
    margins: { top:100, bottom:100, left:150, right:150 },
    children: [p([t(text, { color:textColor, size:11, bold })], { align })]
  });
}

function kpiCard(label, value, sub, fill=C.gray, w=2409) {
  return new TableCell({
    width: { size:w, type:WidthType.DXA },
    shading: { fill, type:ShadingType.CLEAR },
    borders: allBorders('E0E0E0'),
    margins: { top:150, bottom:150, left:200, right:200 },
    children: [
      p([t(label, { size:9, color:fill===C.gray?C.grayTxt:C.white, bold:true })]),
      p([t(value, { size:20, bold:true, color:fill===C.gray?C.navy:C.white })]),
      p([t(sub, { size:10, color:fill===C.gray?C.grayTxt:C.white })]),
    ]
  });
}

function simpleTable(rows, widths) {
  return new Table({
    width: { size:widths.reduce((a,b)=>a+b,0), type:WidthType.DXA },
    columnWidths: widths,
    rows: rows.map(row => new TableRow({ children: row }))
  });
}

function pageBreak() { return new Paragraph({ children: [new PageBreak()] }); }

function generate(d) {
  const pesoN = parseFloat(d.peso||107.7);
  const pesoObN = parseFloat(d.pesoObiettivo||85);
  const pesoInN = parseFloat(d.pesoIniziale||107.7);
  const kgPersi = Math.max(0, pesoInN-pesoN).toFixed(1);
  const kgMancanti = Math.max(0, pesoN-pesoObN).toFixed(1);
  const progresso = pesoInN>pesoObN ? Math.min(100,((pesoInN-pesoN)/(pesoInN-pesoObN))*100).toFixed(1) : '0';
  const calorieGiorni = d.calorieGiorni||[1950,1800,2000,1900,2100,1850,1950];
  const calMedia = Math.round(calorieGiorni.reduce((a,b)=>a+b,0)/calorieGiorni.length);
  const aderenza = Math.round(calorieGiorni.filter(c=>Math.abs(c-2000)<250).length/7*100);
  const giorniPalestra = d.giorniPalestra||3;
  const acquaMedia = d.acquaMedia||2.5;
  const streak = d.streak||7;
  const umoreScore = d.umoreScore||75;
  const stressScore = d.stressScore||35;
  const sonnoScore = d.sonnoScore||70;
  const scoreFinale = Math.round([aderenza, Math.min(100,giorniPalestra*33), Math.min(100,(acquaMedia/3.5)*100), Math.min(100,streak*10), umoreScore].reduce((a,b)=>a+b,0)/5);
  let gradeLabel, gradeColor;
  if (scoreFinale>=85) { gradeLabel='ECCELLENTE'; gradeColor=C.green; }
  else if (scoreFinale>=70) { gradeLabel='OTTIMO'; gradeColor=C.blue; }
  else if (scoreFinale>=55) { gradeLabel='BUONO'; gradeColor=C.orange; }
  else { gradeLabel='DA MIGLIORARE'; gradeColor=C.red; }

  const days = ['LUN','MAR','MER','GIO','VEN','SAB','DOM'];

  const doc = new Document({
    styles: { default: { document: { run: { font:'Arial', size:24, color:C.black } } } },
    sections: [{
      properties: { page: { size: { width:11906, height:16838 }, margin: { top:1134, right:1134, bottom:1134, left:1134 } } },
      headers: {
        default: new Header({ children: [
          simpleTable([[
            new TableCell({ width:{size:6000,type:WidthType.DXA}, borders:noBorders(), children:[p([t('OMAR COACH', {bold:true,size:12,color:C.navy}), t('  |  Weekly Performance Report', {size:10,color:C.grayTxt})])] }),
            new TableCell({ width:{size:3638,type:WidthType.DXA}, borders:noBorders(), children:[p([t('Settimana '+(d.settimana||1)+'  -  '+(d.dataReport||new Date().toLocaleDateString('it-IT')), {size:10,color:C.grayTxt})], {align:AlignmentType.RIGHT})] }),
          ]], [6000,3638])
        ]})
      },
      footers: {
        default: new Footer({ children: [
          new
