// WB の図 試作：本物の Excel に、アプリと同じ図を描かせられるかを確かめる。
// 数は作り物（実験のデータではない）。機能ごとに try し、できた・できなかったを下に 1 行ずつ出す。
'use strict';

const ROWS = [
  ['Control', '', '0 h', '', 1.0, 0, ''],
  ['Control', '', '0.5 h', 'RSL3 0.75 µM', 0.92, 0.05, ''],
  ['Control', '', '1 h', 'RSL3 0.75 µM', 0.77, 0.05, ''],
  ['Control', '', '2 h', 'RSL3 0.75 µM', 0.71, 0.07, '*'],
  ['selenite', '500 nM', '0 h', '', 1.67, 0.18, '***'],
  ['selenite', '500 nM', '0.5 h', 'RSL3 0.75 µM', 1.42, 0.09, '**'],
  ['selenite', '500 nM', '1 h', 'RSL3 0.75 µM', 1.22, 0.05, ''],
  ['selenite', '500 nM', '2 h', 'RSL3 0.75 µM', 1.08, 0.02, ''],
  ['HSeGalNAc', '5 µM', '0 h', '', 2.18, 0.11, '***'],
  ['HSeGalNAc', '5 µM', '0.5 h', 'RSL3 0.75 µM', 2.06, 0.26, '***'],
  ['HSeGalNAc', '5 µM', '1 h', 'RSL3 0.75 µM', 1.68, 0.10, '***'],
  ['HSeGalNAc', '5 µM', '2 h', 'RSL3 0.75 µM', 1.38, 0.07, '**']
];
const FILL = { Control: '#D9D9D9', selenite: '#7F7F7F', HSeGalNAc: '#0072B2' };
const N = ROWS.length, LAST = N + 1;

const log = (cls, text) => { const li = document.createElement('li'); li.className = cls; li.textContent = text; document.getElementById('log').appendChild(li); };
const api = v => Office.context.requirements.isSetSupported('ExcelApi', v);

// 1 つの機能を試す。失敗しても次へ進む（どこまでできるかを見るため）
async function step(context, name, fn) {
  try { await fn(); await context.sync(); log('ok', name); return true; }
  catch (e) { log('ng', name + '：' + (e.message || e)); return false; }
}

async function freshSheet(context, name) {
  const old = context.workbook.worksheets.getItemOrNullObject(name);
  await context.sync();
  if (!old.isNullObject) { old.delete(); await context.sync(); }
  const ws = context.workbook.worksheets.add(name); ws.activate(); return ws;
}

async function buildA() {
  document.getElementById('log').innerHTML = '';
  log('info', 'Excel API：' + ['1.7', '1.8', '1.9', '1.13', '1.17'].map(v => v + (api(v) ? ' ○' : ' ×')).join('　'));
  await Excel.run(async context => {
    const ws = await freshSheet(context, '試作A');
    await step(context, '値を書く（記号の高さは 平均＋SD の式）', async () => {
      ws.getRange('A1:H1').values = [['化合物', '濃度', '時間', '処理', '平均', 'SD', '記号', '記号の高さ']];
      ws.getRange('A2:G' + LAST).values = ROWS;
      ws.getRange('H2:H' + LAST).formulas = ROWS.map((r, i) => ['=E' + (i + 2) + '+F' + (i + 2)]);
    });
    const chart = ws.charts.add(Excel.ChartType.columnClustered, ws.getRange('E2:E' + LAST), Excel.ChartSeriesBy.columns);
    chart.top = 10; chart.left = 480; chart.width = 480; chart.height = 300;
    const bars = chart.series.getItemAt(0);
    await step(context, '図表を作る（棒は平均の列につながる）', async () => { bars.setXAxisValues(ws.getRange('C2:C' + LAST)); });
    await step(context, '題・凡例・字（Arial）', async () => {
      chart.title.text = 'SELENOP'; chart.title.format.font.name = 'Arial'; chart.title.format.font.size = 14; chart.title.format.font.bold = true;
      chart.legend.visible = false; chart.format.font.name = 'Arial'; chart.format.font.size = 9; chart.format.font.color = '#000000';
    });
    await step(context, '縦軸：0〜3.0、刻み 0.5、「0.0」、格子線なし、目盛は外向き', async () => {
      const y = chart.axes.valueAxis;
      y.minimum = 0; y.maximum = 3; y.majorUnit = 0.5; y.numberFormat = '0.0';
      y.majorGridlines.visible = false; y.majorTickMark = 'Outside';
      y.format.line.color = '#000000'; y.format.line.weight = 0.75;
    });
    await step(context, '横軸：目盛は外向きの短い線', async () => {
      const x = chart.axes.categoryAxis; x.majorTickMark = 'Outside'; x.format.line.color = '#000000'; x.format.line.weight = 0.75;
    });
    await step(context, '棒の間（要素の間隔 50）', async () => { bars.gapWidth = 50; });
    await step(context, '棒ごとの色と黒い枠', async () => {
      ROWS.forEach((r, i) => {
        const p = bars.points.getItemAt(i);
        p.format.fill.setSolidColor(FILL[r[0]]);
        p.format.border.color = '#000000'; p.format.border.weight = 0.75; p.format.border.lineStyle = 'Continuous';
      });
    });
    // 誤差範囲：SD のセルにつなげるか（ユーザー設定の値）。API で範囲を渡せるかを見る
    await step(context, '誤差範囲：両方向・キャップあり', async () => {
      const e = bars.yErrorBars; e.visible = true; e.include = 'Both'; e.endStyleCap = true;
    });
    await step(context, '誤差範囲：SD のセルを値にする（ユーザー設定）', async () => {
      const e = bars.yErrorBars; e.type = 'Custom';
      if (typeof e.setCustomValues !== 'function') throw Error('API に「ユーザー設定の値を範囲で渡す」方法が無い');
      e.setCustomValues(ws.getRange('F2:F' + LAST), ws.getRange('F2:F' + LAST));
    });
    // 記号：記号の高さの列を見えない折れ線にし、点ごとのラベルを「記号」のセルにつなぐ
    let marks = null;
    await step(context, '記号の系列（記号の高さ＝平均＋SD）', async () => {
      marks = chart.series.add('記号');
      marks.setValues(ws.getRange('H2:H' + LAST));
      marks.chartType = 'Line'; marks.markerStyle = 'None'; marks.format.line.lineStyle = 'None';
    });
    await step(context, '記号のラベルを「記号」のセルにつなぐ（点ごとの式）', async () => {
      if (!marks) throw Error('記号の系列が無い');
      marks.hasDataLabels = true;
      marks.dataLabels.showValue = false; marks.dataLabels.position = 'Top';
      ROWS.forEach((r, i) => { marks.points.getItemAt(i).dataLabel.formula = "='試作A'!$G$" + (i + 2); });
    });
    // 棒の下の表（RSL3 の＋／−・濃度・化合物）：図表の中に文字の図形を置く API があるか
    await step(context, '棒の下の表を図表の中の文字で置く', async () => {
      if (!chart.shapes) throw Error('図表の中に図形を足す API が無い（シートの上に置くしかない）');
    });
  });
  log('info', 'A：おわり。右の図をアプリの図と見比べてください。');
}

// B：アプリ（MTSとWB.html）が書いた Excel の「図の値」「論文」シートを、図表ごとこのブックへ差し込む
async function buildB() {
  document.getElementById('log').innerHTML = '';
  if (!api('1.13')) { log('ng', 'この Excel は ExcelApi 1.13（シートの差し込み）に対応していない'); return; }
  let b64;
  try {
    const buf = await (await fetch('sample.xlsx')).arrayBuffer();
    let s = ''; const u = new Uint8Array(buf); for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
    b64 = btoa(s); log('ok', '見本の Excel を読んだ（' + Math.round(u.length / 1024) + ' KB）');
  } catch (e) { log('ng', '見本の Excel を読めない：' + e.message); return; }
  await Excel.run(async context => {
    await step(context, '「図の値」「論文」を差し込む', async () => {
      context.workbook.insertWorksheetsFromBase64(b64, { sheetNamesToInsert: ['図の値', '論文'], positionType: Excel.WorksheetPositionType.end });
    });
    for (const name of ['図の値', '論文']) {
      const ws = context.workbook.worksheets.getItemOrNullObject(name), n = ws.charts.getCount();
      await context.sync();
      if (ws.isNullObject) { log('ng', name + '：シートが無い'); continue; }
      log(n.value ? 'ok' : 'ng', name + '：図表 ' + n.value + ' 個');
    }
  });
  log('info', 'B：おわり。差し込んだシートの図表で、記号・誤差範囲・棒の下の表が残っているかを見てください。');
}

Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) { log('ng', 'Excel の中で開いてください'); return; }
  document.getElementById('a').onclick = () => buildA().catch(e => log('ng', e.message));
  document.getElementById('b').onclick = () => buildB().catch(e => log('ng', e.message));
  log('info', '準備ができました');
});
