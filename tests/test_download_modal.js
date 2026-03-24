const test = require('node:test');
const assert = require('node:assert/strict');

const helpers = require('../web_app/static/download_modal.js');

test('initial state requires confirmation before download', () => {
  const state = helpers.buildInitialState();

  assert.equal(state.confirmedFilename, '');
  assert.equal(state.canDownload, false);
  assert.match(state.statusText, /请先输入文件名/);
});

test('confirmed state enables download with chosen filename', () => {
  const state = helpers.buildConfirmedState('sales.v2.final', '.csv');

  assert.equal(state.confirmedFilename, 'sales.v2.final');
  assert.equal(state.canDownload, true);
  assert.equal(state.statusText, '已确认文件名：sales.v2.final.csv');
});

test('editing after confirmation invalidates the download state', () => {
  const state = helpers.buildEditedState('sales.v3.final');

  assert.equal(state.confirmedFilename, '');
  assert.equal(state.canDownload, false);
  assert.equal(state.statusText, '文件名已变更，请重新点击“确定”。');
});

test('empty input keeps download disabled', () => {
  const state = helpers.buildEditedState('   ');

  assert.equal(state.confirmedFilename, '');
  assert.equal(state.canDownload, false);
  assert.equal(state.statusText, '文件名不能为空。');
});
