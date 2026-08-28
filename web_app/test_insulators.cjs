const assert = require('node:assert/strict');
const { insulatorRate, insulatorQuantity, summarizeInsulators } = require('./static/insulators.js');
for (const [head, expected] of Object.entries({
  BA: [4, 12], DE: [0, 12], DDE: [6, 24], 'DDE.BL': [0, 24],
  SP: [3, 0], DP: [6, 0], 'CCB บน': [3, 0], 'CCB ล่าง': [3, 0],
  'CCB ประกบบน': [6, 0], 'CCB ประกบล่าง': [6, 0], 'CCB,CCB': [6, 0],
  'DDE.BL.st 2.5m': [0, 24], 'DDE.st 4.5m': [6, 24],
  '2BA.st 4.5 m': [8, 24], '2DE.st4.5': [0, 24], '2DDE.st 3.0m': [12, 48],
  '2SP บน, ล่าง': [6, 0], '2DP.st 3m': [12, 0], 'SP บน,ล่าง': [3, 0],
  'SP บน, ล่าง': [3, 0], 'SP 1-P': [2, 0], 'BA 1-P': [2, 8], 'BA.AL 1P': [2, 8],
  'DDE 1-P': [4, 16], 'DDE.BL 1-P': [0, 16], 'DE.CON 1-P': [4, 8],
  'CTB บน': [0, 0], 'CTB ประกบบน': [0, 0], 'CSC ปลายคอน DDE': [0, 0],
  'BA.AL': [4, 12], 'DDE.AL': [6, 24], 'DE.AL': [0, 12], 'BA.SLK บน': [6, 0],
  'LAT.SLK บน': [6, 12], 'LAT.SLK ล่าง': [6, 12], 'DDE.st 3m, LAT.SLK': [12, 36],
  'DDE,DP.st 3.0m': [12, 24], '2BA.st 4.5m+DE.CON': [12, 36], '2DE.st4.5 + DE.CON': [6, 36],
})) assert.deepEqual(insulatorRate(head), expected, head);
for (const head of ['Unknown', 'DE+Unknown', '2DE, 2LAT.SLK Steel 4.50m']) {
  assert.equal(insulatorRate(head), null, head);
}
for (const [expression, expected] of [['4+4+5+6', 19], ['6-2', 4], ['(4+5)-2', 7], ['-2+3', 1], ['1.5+0.5', 2], ['1e2-1', 99]]) {
  assert.equal(insulatorQuantity(expression), expected);
}
for (const expression of ['', '1+', '1 2', '3*2', 'alert(1)', '(2+3', 'Infinity']) {
  assert.equal(insulatorQuantity(expression), null);
}
const row = (head, count) => ({ size: '14.3', head, count });
const pages = [[row('BA', '1+1'), row('DDE', '1')], [row('CCB ประกบบน', '2'), row('DDE.BL', '1')]];
const before = JSON.stringify(pages);
assert.deepEqual(summarizeInsulators(pages), { upright: 26, horizontal: 72, warnings: [] });
assert.equal(JSON.stringify(pages), before, 'Must not mutate saved project data');
assert.deepEqual(summarizeInsulators([[]]), { upright: 0, horizontal: 0, warnings: [] });
assert.equal(summarizeInsulators([[row('BA', '1+'), row('Unknown', '1'), row('BA', '-1')]]).warnings.length, 3);
console.log('Insulator count tests passed');
