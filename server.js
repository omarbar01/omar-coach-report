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
          new Paragraph({ border:{top:{style:BorderStyle.SINGLE,size:4,color:C.grayMid,space:1}}, spacing:{before:80}, alignment:AlignmentType.CENTER, children:[t('Omar Coach  -  Report Settimanale  -  '+(d.dataReport||''), {size:9,color:C.grayTxt})] })
        ]})
      },
      children: [
        sp(), sp(), sp(),
        simpleTable([[
          new TableCell({ width:{size:1400,type:WidthType.DXA}, shading:{fill:C.navy,type:ShadingType.CLEAR}, borders:allBorders(C.navy), margins:{top:180,bottom:180,left:100,right:100}, children:[p([t('OC',{bold:true,size:28,color:C.white})],{align:AlignmentType.CENTER})] }),
        ]], [1400]),
        sp(),
        p([t('OMAR COACH', {bold:true,size:30,color:C.navy})]),
        p([t('Weekly Performance Report', {size:15,color:C.grayTxt})]),
        divider(C.navy), sp(),
        simpleTable([[
          kpiCard('ATLETA', 'Omar Barhami', '24 anni - 180 cm - Settimana '+(d.settimana||1), C.gray, 4819),
          kpiCard('PERFORMANCE SCORE', scoreFinale+'/100', gradeLabel, gradeColor, 4819),
        ]], [4819,4819]),
        sp(), sp(),
        simpleTable([[
          new TableCell({ width:{size:9638,type:WidthType.DXA}, shading:{fill:C.navy,type:ShadingType.CLEAR}, borders:allBorders(C.navy), margins:{top:200,bottom:200,left:300,right:300}, children:[
            p([t('OBIETTIVO TRASFORMAZIONE', {size:9,color:'AABBCC',bold:true})]),
            p([t('Peso attuale '+(d.peso||'107.7')+' kg  >  Obiettivo '+(d.pesoObiettivo||'85')+' kg  (-'+kgMancanti+' kg rimasti)', {size:13,color:C.white,bold:true})]),
            p([t(bar(parseFloat(progresso),45)+'  '+progresso+'%', {size:10,color:C.green,font:'Courier New'})]),
          ]})
        ]], [9638]),
        pageBreak(),

        h('DASHBOARD CORPOREO', 16), divider(), sp(),
        simpleTable([[
          kpiCard('PESO ATTUALE', (d.peso||'107.7')+' kg', parseFloat(kgPersi)>0?'Calati '+kgPersi+' kg':'Inizio percorso', C.gray, 2409),
          kpiCard('OBIETTIVO', (d.pesoObiettivo||'85')+' kg', kgMancanti+' kg rimasti', C.gray, 2409),
          kpiCard('PROGRESSI', progresso+'%', 'verso obiettivo', gradeColor, 2410),
          kpiCard('STREAK', streak+' giorni', 'consecutivi', C.gray, 2410),
        ]], [2409,2409,2410,2410]),
        sp(), h('Misurazioni Corporee', 13, C.black),
        simpleTable([
          [cell('MISURAZIONE',C.navy,C.white,3000,AlignmentType.LEFT,true), cell('VALORE',C.navy,C.white,2319,AlignmentType.CENTER,true), cell('OBIETTIVO',C.navy,C.white,2319,AlignmentType.CENTER,true), cell('TREND',C.navy,C.white,2000,AlignmentType.CENTER,true)],
          [cell('Peso',C.white,C.black,3000,AlignmentType.LEFT), cell((d.peso||'107.7')+' kg',C.white,C.navy,2319), cell((d.pesoObiettivo||'85')+' kg',C.white,C.grayTxt,2319), cell(parseFloat(kgPersi)>0?'In calo':'Stabile',C.white,C.green,2000)],
          [cell('Giro Vita',C.gray,C.black,3000,AlignmentType.LEFT), cell((d.vita||'-')+' cm',C.gray,C.navy,2319), cell('Ridurre',C.gray,C.grayTxt,2319), cell('-',C.gray,C.grayTxt,2000)],
          [cell('Giro Collo',C.white,C.black,3000,AlignmentType.LEFT), cell((d.collo||'-')+' cm',C.white,C.navy,2319), cell('Ridurre',C.white,C.grayTxt,2319), cell('-',C.white,C.grayTxt,2000)],
          [cell('Giro Braccia',C.gray,C.black,3000,AlignmentType.LEFT), cell((d.braccia||'-')+' cm',C.gray,C.navy,2319), cell('Tonificare',C.gray,C.grayTxt,2319), cell('-',C.gray,C.grayTxt,2000)],
          [cell('Giro Petto',C.white,C.black,3000,AlignmentType.LEFT), cell((d.petto||'-')+' cm',C.white,C.navy,2319), cell('Tonificare',C.white,C.grayTxt,2319), cell('-',C.white,C.grayTxt,2000)],
        ], [3000,2319,2319,2000]),
        sp(),
        h('Progressi verso Obiettivo', 13, C.black),
        p([t(bar(parseFloat(progresso),50), {font:'Courier New',size:11,color:C.navy})]),
        p([t(progresso+'% completato  -  Mancano '+kgMancanti+' kg per raggiungere '+( d.pesoObiettivo||'85')+' kg', {size:10,color:C.grayTxt})]),
        pageBreak(),

        h('NUTRIZIONE SETTIMANALE', 16), divider(), sp(),
        simpleTable([[
          kpiCard('MEDIA CALORIE/GIORNO', calMedia+' kcal', 'Target: 2.000 kcal', calMedia<=2100?C.green:C.red, 3212),
          kpiCard('ADERENZA AL PIANO', aderenza+'%', 'giorni rispettati', C.gray, 3213),
          kpiCard('DEFICIT TOTALE', Math.round((2000-calMedia)*7)+' kcal', 'questa settimana', C.gray, 3213),
        ]], [3212,3213,3213]),
        sp(), h('Calorie Giornaliere', 13, C.black),
        ...calorieGiorni.map((cal,i) => miniBar(days[i]+'  ', cal, 2500, ' kcal')),
        sp(), h('Macronutrienti Target', 13, C.black),
        simpleTable([
          [cell('PROTEINE',C.blue,C.white,2409,AlignmentType.CENTER,true), cell('CARBOIDRATI',C.orange,C.white,2409,AlignmentType.CENTER,true), cell('GRASSI',C.navy,C.white,2410,AlignmentType.CENTER,true), cell('FIBRE',C.greenDk,C.white,2410,AlignmentType.CENTER,true)],
          [cell('160g / giorno',C.white,C.blue,2409), cell('200g / giorno',C.white,C.orange,2409), cell('62g / giorno',C.white,C.navy,2410), cell('35g / giorno',C.white,C.greenDk,2410)],
        ], [2409,2409,2410,2410]),
        pageBreak(),

        h('ATTIVITA FISICA', 16), divider(), sp(),
        simpleTable([[
          kpiCard('SESSIONI PALESTRA', giorniPalestra+'/7', giorniPalestra>=3?'Obiettivo raggiunto':'Continua cosi!', giorniPalestra>=3?C.green:C.orange, 3212),
          kpiCard('ACQUA MEDIA', acquaMedia+'L', 'Target: 3.5L/giorno', C.gray, 3213),
          kpiCard('CALORIE BRUCIATE', '~'+(giorniPalestra*450)+' kcal', 'stimate palestra', C.gray, 3213),
        ]], [3212,3213,3213]),
        sp(), h('Piano Allenamento', 13, C.black),
        simpleTable([
          days.map((day,i) => new TableCell({ width:{size:i<6?1371:1412,type:WidthType.DXA}, shading:{fill:C.navy,type:ShadingType.CLEAR}, borders:allBorders(C.navy), margins:{top:80,bottom:80,left:60,right:60}, children:[p([t(day,{bold:true,color:C.white,size:10})],{align:AlignmentType.CENTER})] })),
          [{label:'Petto/Tricipiti',pal:true},{label:'Riposo Attivo',pal:false},{label:'Schiena/Bicipiti',pal:true},{label:'Riposo Attivo',pal:false},{label:'Gambe/Spalle',pal:true},{label:'Cardio 30min',pal:false},{label:'Riposo',pal:false}]
            .map((item,i) => new TableCell({ width:{size:i<6?1371:1412,type:WidthType.DXA}, shading:{fill:item.pal?C.blue:C.gray,type:ShadingType.CLEAR}, borders:allBorders('CCCCCC'), margins:{top:80,bottom:80,left:60,right:60}, children:[p([t(item.label,{color:item.pal?C.white:C.grayTxt,size:9})],{align:AlignmentType.CENTER})] })),
        ], [1371,1371,1371,1371,1371,1371,1412]),
        sp(), h('Idratazione', 13, C.black),
        miniBar('Acqua media   ', acquaMedia, 3.5, 'L'),
        pageBreak(),

        h('BENESSERE E RECUPERO', 16), divider(), sp(),
        simpleTable([[
          kpiCard('UMORE MEDIO', umoreScore+'/100', umoreScore>=70?'Ottimo':'Da migliorare', C.gray, 3212),
          kpiCard('QUALITA SONNO', sonnoScore+'/100', sonnoScore>=70?'Buon recupero':'Dormi di piu', C.gray, 3213),
          kpiCard('LIVELLO STRESS', stressScore+'/100', stressScore<=40?'Gestito bene':'Alto - attenzione', C.gray, 3213),
        ]], [3212,3213,3213]),
        sp(), h('Score Componenti', 13, C.black),
        miniBar('Disciplina Dieta     ', aderenza, 100, '%'),
        miniBar('Attivita Fisica      ', Math.min(100,giorniPalestra*33), 100, '%'),
        miniBar('Idratazione          ', Math.min(100,(acquaMedia/3.5)*100), 100, '%'),
        miniBar('Costanza Streak      ', Math.min(100,streak*10), 100, '%'),
        miniBar('Benessere Mentale    ', umoreScore, 100, '%'),
        sp(),
        simpleTable([[
          new TableCell({ width:{size:9638,type:WidthType.DXA}, shading:{fill:gradeColor,type:ShadingType.CLEAR}, borders:allBorders(gradeColor), margins:{top:200,bottom:200,left:300,right:300}, children:[
            p([t('SCORE FINALE SETTIMANA', {size:10,color:C.white,bold:true})]),
            p([t(scoreFinale+'/100  -  '+gradeLabel, {size:20,bold:true,color:C.white})]),
            p([t(bar(scoreFinale,50), {font:'Courier New',size:10,color:C.white})]),
          ]})
        ]], [9638]),
        pageBreak(),

        h('ANALISI DEL COACH', 16), divider(), sp(),
        simpleTable([[
          new TableCell({ width:{size:4719,type:WidthType.DXA}, shading:{fill:'F0FFF4',type:ShadingType.CLEAR}, borders:{top:bdr(C.greenDk,8),bottom:bdr(C.greenDk,8),left:bdr(C.greenDk,8),right:bdr(C.greenDk,8)}, margins:{top:200,bottom:200,left:300,right:300}, children:[
            p([t('PUNTI DI FORZA', {size:11,color:C.greenDk,bold:true})]), sp(),
            p([t(d.costoFatto||'Hai dimostrato costanza questa settimana.', {size:11,color:C.black})]),
          ]}),
          new TableCell({ width:{size:4919,type:WidthType.DXA}, shading:{fill:'FFFBF0',type:ShadingType.CLEAR}, borders:{top:bdr(C.orange,8),bottom:bdr(C.orange,8),left:bdr(C.orange,8),right:bdr(C.orange,8)}, margins:{top:200,bottom:200,left:300,right:300}, children:[
            p([t('AREE DI MIGLIORAMENTO', {size:11,color:C.orange,bold:true})]), sp(),
            p([t(d.cosaDaMigliorare||'Aumenta l idratazione giornaliera.', {size:11,color:C.black})]),
          ]}),
        ]], [4719,4919]),
        sp(), h('Piano Settimana Prossima', 13, C.black),
        simpleTable([[
          new TableCell({ width:{size:9638,type:WidthType.DXA}, shading:{fill:C.gray,type:ShadingType.CLEAR}, borders:allBorders('E0E0E0'), margins:{top:200,bottom:200,left:300,right:300}, children:[
            p([t('OBIETTIVO', {size:10,color:C.navy,bold:true})]),
            p([t(d.obiettivo||'Palestra 3 volte + Dieta + Acqua 3L', {size:12,color:C.black})]),
          ]})
        ]], [9638]),
        sp(),
        simpleTable([[
          new TableCell({ width:{size:9638,type:WidthType.DXA}, shading:{fill:C.navy,type:ShadingType.CLEAR}, borders:allBorders(C.navy), margins:{top:250,bottom:250,left:400,right:400}, children:[
            p([t('IL TUO COACH DICE:', {size:10,color:'AABBCC',bold:true})]), sp(),
            p([t('"'+(d.messaggioCoach||'Continua cosi, ogni giorno conta!')+'"', {size:13,color:C.white,italic:true})]),
            sp(), p([t('- Omar Coach', {size:10,color:'AABBCC',bold:true})], {align:AlignmentType.RIGHT}),
          ]})
        ]], [9638]),
        pageBreak(),

        h('ACHIEVEMENT E TROFEI', 16), divider(), sp(),
        h('Trofei Sbloccati Questa Settimana', 13, C.black),
        ...(() => {
          const trofei = [];
          if (giorniPalestra>=3) trofei.push(['🏆','HAT-TRICK','3+ sessioni di palestra questa settimana',C.green]);
          if (streak>=7) trofei.push(['🥇','SETTIMANA PERFETTA','7 giorni consecutivi di streak',C.blue]);
          if (parseFloat(kgPersi)>0) trofei.push(['⚽','PRIMO PASSO',kgPersi+' kg persi dall inizio',C.orange]);
          if (acquaMedia>=3.5) trofei.push(['💧','IDRATAZIONE OK','Media 3.5L/giorno raggiunta',C.blue]);
          if (aderenza>=80) trofei.push(['🍽','DISCIPLINA','Dieta rispettata oltre l 80%',C.greenDk]);
          if (trofei.length===0) trofei.push(['💪','IN CAMMINO','Continua cosi - i trofei arrivano!',C.grayTxt]);
          return trofei.map(([icon,nome,desc,color]) =>
            simpleTable([[
              new TableCell({ width:{size:800,type:WidthType.DXA}, shading:{fill:color,type:ShadingType.CLEAR}, borders:allBorders(color), margins:{top:120,bottom:120,left:100,right:100}, children:[p([t(icon,{size:16})],{align:AlignmentType.CENTER})] }),
              new TableCell({ width:{size:8838,type:WidthType.DXA}, borders:allBorders('E0E0E0'), margins:{top:120,bottom:120,left:200,right:200}, children:[p([t(nome,{bold:true,size:12,color:C.navy}), t('  -  '+desc,{size:11,color:C.grayTxt})])] }),
            ]], [800,8838])
          );
        })(),
        sp(), sp(),
        simpleTable([[
          new TableCell({ width:{size:9638,type:WidthType.DXA}, shading:{fill:C.gray,type:ShadingType.CLEAR}, borders:{top:bdr(C.navy,12),bottom:bdr(C.navy,12),left:bdr(C.navy,12),right:bdr(C.navy,12)}, margins:{top:300,bottom:300,left:400,right:400}, children:[
            p([t('OMAR COACH', {bold:true,size:14,color:C.navy})], {align:AlignmentType.CENTER}),
            p([t('Il tuo percorso di trasformazione, un giorno alla volta.', {size:12,italic:true,color:C.grayTxt})], {align:AlignmentType.CENTER}),
            p([t('Settimana '+(d.settimana||1)+'  -  '+(d.dataReport||'')+'  -  Score: '+scoreFinale+'/100', {size:10,color:C.grayTxt})], {align:AlignmentType.CENTER}),
          ]})
        ]], [9638]),
      ]
    }]
  });
  return doc;
}

app.post('/generate', async (req, res) => {
  try {
    const doc = generate(req.body);
    const buffer = await Packer.toBuffer(doc);
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': 'attachment; filename="Omar_Coach_Report.docx"',
      'Content-Length': buffer.length
    });
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/', (req, res) => res.json({ status: 'Omar Coach Report Server running!' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Server running on port ' + PORT));
