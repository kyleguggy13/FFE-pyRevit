// Run: node --test tests/revit_data_table.test.cjs
const { test } = require('node:test');
const assert = require('node:assert/strict');
const model = require('../support/table_model.js');

const rows = [
  { id: '9007199254740993', phaseId: '1', levelId: '10', level: 'Ground', levelElevation: 0,
    number: '002', name: 'Office', areaValue: 100, heightValue: 9 },
  { id: '9007199254740994', phaseId: '2', levelId: '10', level: 'Ground', levelElevation: 0,
    number: '10', name: '<script>literal text</script>', areaValue: 20, heightValue: null },
  { id: '3', phaseId: '2', levelId: '11', level: 'Basement', levelElevation: -10,
    number: '1', name: 'Storage', areaValue: 0, heightValue: 0 },
];
const snapshot = { rows, defaultPhaseId: '2', phases: [{ id: '1' }, { id: '2' }], levels: [{ id: '10' }, { id: '11' }] };

test('phase, level, and case-insensitive search intersect without altering rows', () => {
  const filters = { phase: '2', level: '10', query: ' LITERAL ' };
  assert.deepEqual(rows.filter(row => model.matches(row, filters)).map(row => row.id), ['9007199254740994']);
  assert.equal(rows.length, 3);
  assert.equal(model.matches(rows[0], { phase: 'all', level: 'all', query: 'ground' }), true);
});

test('natural number order preserves leading zeros and alphanumeric identifiers', () => {
  assert.deepEqual(['10', '002', 'A10', 'A2'].sort(model.compareNatural), ['002', '10', 'A2', 'A10']);
});

test('default level ordering uses elevation and numeric sorting distinguishes zero and missing', () => {
  assert.equal([...rows].sort(model.compareLevel)[0].id, '3');
  assert.deepEqual([null, 100, 0, 20].sort(model.compareNumeric), [0, 20, 100, null]);
});

test('refresh preserves valid filters and falls back when a phase or level disappears', () => {
  const filters = { phase: '1', level: '10', query: 'Office' };
  assert.deepEqual(model.reconcileFilters(snapshot, filters), filters);
  assert.deepEqual(model.reconcileFilters({ ...snapshot, phases: [{ id: '2' }], levels: [] }, filters),
    { phase: '2', level: 'all', query: 'Office' });
  assert.equal(model.reconcileFilters(snapshot, { phase: 'all', level: 'all' }).phase, 'all');
  assert.equal(model.reconcileFilters(snapshot, null).phase, '2');
});

test('navigation and refresh selection exclude deleted and filtered rows without rounding IDs', () => {
  assert.deepEqual(model.selectedVisibleIds(rows, ['9007199254740993', '9007199254740994', 'deleted'],
    { phase: '2', level: 'all', query: '' }), ['9007199254740994']);
});
