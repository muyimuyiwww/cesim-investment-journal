import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const [inputPath, outlookPath, macroPath, outputPath] = process.argv.slice(2);
if (!inputPath || !outlookPath || !macroPath || !outputPath) {
  throw new Error('用法: node tools/enrich_results.mjs 官网结果.xlsx 经济展望.json 宏观情景.json 输出.xlsx');
}

const outlook = JSON.parse(await fs.readFile(outlookPath, 'utf8'));
const macro = JSON.parse(await fs.readFile(macroPath, 'utf8'));
if (!Number.isInteger(outlook.round) || !Array.isArray(outlook.rows) || outlook.rows.length === 0) {
  throw new Error('经济展望缺少回合编号或逐项记录。');
}
if (outlook.rows.some(r => r.length < 5 || !Number.isFinite(r[3]) || !Number.isFinite(r[4]))) {
  throw new Error('经济展望的估计值或实现值缺失。');
}
if (!Array.isArray(macro.rows) || macro.rows.some(r => r.length < 4 || !Number.isFinite(r[3]))) {
  throw new Error('宏观情景数据缺项。');
}

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const originalSheets = [wb.worksheets.getItemAt(0).name];
if (originalSheets.includes('经济展望预测')) throw new Error('原文件已经有“经济展望预测”工作表，请从原始导出重新生成。');
const sheet = wb.worksheets.add('经济展望预测');
sheet.showGridLines = false;
sheet.tabColor = '#1976A3';
sheet.freezePanes.freezeRows(6);

sheet.getRange('A1:I1').values = [[`Cesim Invest｜回合 ${outlook.round} 经济展望预测与实现`, '', '', '', '', '', '', '', '']];
sheet.getRange('A2:I2').values = [[`期间：${outlook.period ?? '以官网结果为准'}　补录：${outlook.captured_at ?? '日期未记录'}　原始 Results 工作表来自 Cesim 下载文件`, '', '', '', '', '', '', '', '']];
sheet.getRange('A4:I4').values = [['市场板块表现：个人决策区的上回合估计值与上回合实现值', '', '', '', '', '', '', '', '']];
sheet.getRange('A6:I6').values = [['类别', '币种', '指标', '当时估计值', '本回合实现值', '误差：实现减估计', '绝对误差', '单位', '数据来源']];
const forecastRows = outlook.rows.map(r => [r[0], r[1], r[2], r[3], r[4], +(r[4] - r[3]).toFixed(2), +Math.abs(r[4] - r[3]).toFixed(2), '%', 'Cesim 经济展望：上回合估计／实现']);
sheet.getRange(`A7:I${6 + forecastRows.length}`).values = forecastRows;

const macroTitleRow = 9 + forecastRows.length;
const macroHeaderRow = macroTitleRow + 2;
const macroStartRow = macroHeaderRow + 1;
sheet.getRange(`A${macroTitleRow}:I${macroTitleRow}`).values = [['经济基本面：两个回合的网站预期并列，均非实际结果', '', '', '', '', '', '', '', '']];
sheet.getRange(`A${macroHeaderRow}:I${macroHeaderRow}`).values = [['指标', '地区或币种', '单位', `回合${outlook.round}网站预期`, `回合${outlook.round + 1}网站预期`, '预期变化', '口径', '补录日', '数据来源']];
const macroRows = macro.rows.map(r => [r[0], r[1], r[2], r[3], Number.isFinite(r[4]) ? r[4] : null, Number.isFinite(r[4]) ? +(r[4] - r[3]).toFixed(2) : null, '跨回合情景变化；非实现误差', macro.captured_at ?? outlook.captured_at ?? '', 'Cesim 经济展望：经济基本面']);
sheet.getRange(`A${macroStartRow}:I${macroStartRow + macroRows.length - 1}`).values = macroRows;

const noteRow = macroStartRow + macroRows.length + 2;
sheet.getRange(`A${noteRow}:I${noteRow}`).values = [['口径说明：债券比较期末收益率／信用利差，股票比较因子与行业回报；误差单位为百分点。不同类别的平均误差不代表投资收益。', '', '', '', '', '', '', '', '']];
sheet.getRange(`A${noteRow + 1}:I${noteRow + 1}`).values = [[`预测归属：${outlook.forecast_owner ?? '个人决策区输入，尚未区分分析师建议与人工调整'}。`, '', '', '', '', '', '', '', '']];
sheet.getRange(`A${noteRow + 2}:I${noteRow + 2}`).values = [[`来源：${outlook.source ?? 'Cesim 经济展望'}；宏观情景：${macro.source ?? 'Cesim 经济基本面'}。`, '', '', '', '', '', '', '', '']];

sheet.getRange(`A2:I${noteRow + 2}`).format.font = { name: 'Arial', size: 10 };
for (const range of [`A1:I1`, `A4:I4`, `A${macroTitleRow}:I${macroTitleRow}`]) {
  sheet.getRange(range).format.fill = '#17324D';
  sheet.getRange(range).format.font = { name: 'Arial', size: 11, color: '#FFFFFF', bold: true };
  sheet.getRange(range).format.rowHeight = 25;
}
for (const range of [`A6:I6`, `A${macroHeaderRow}:I${macroHeaderRow}`]) {
  sheet.getRange(range).format.fill = '#DCEAF2';
  sheet.getRange(range).format.font = { name: 'Arial', size: 10, color: '#17324D', bold: true };
  sheet.getRange(range).format.rowHeight = 24;
}
sheet.getRange(`A1:A${noteRow + 2}`).format.columnWidth = 25;
sheet.getRange(`B1:B${noteRow + 2}`).format.columnWidth = 15;
sheet.getRange(`C1:C${noteRow + 2}`).format.columnWidth = 20;
sheet.getRange(`D1:G${noteRow + 2}`).format.columnWidth = 20;
sheet.getRange(`H1:H${noteRow + 2}`).format.columnWidth = 29;
sheet.getRange(`I1:I${noteRow + 2}`).format.columnWidth = 36;
sheet.getRange(`D7:G${6 + forecastRows.length}`).setNumberFormat('0.00');
sheet.getRange(`D${macroStartRow}:F${macroStartRow + macroRows.length - 1}`).setNumberFormat('0.00');
sheet.getRange(`A${noteRow}:I${noteRow + 2}`).format.rowHeight = 28;

wb.recalculate();
await fs.mkdir(path.dirname(outputPath), { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
const preview = await wb.render({ sheetName: '经济展望预测', range: 'A1:I13', scale: 1.2, format: 'png' });
await fs.mkdir('private/qa', { recursive: true });
await fs.writeFile('private/qa/forecast-preview.png', new Uint8Array(await preview.arrayBuffer()));
console.log(JSON.stringify({ outputPath, originalSheets, forecasts: forecastRows.length, macroScenarios: macroRows.length }));
