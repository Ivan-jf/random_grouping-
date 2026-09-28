const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { test } = require('node:test');

const html = fs.readFileSync('templates/index.html', 'utf8');
const script = html.slice(html.indexOf('<script>') + 8, html.indexOf('</script>'));

test('preview shows consecutive positions and preserves source sequence numbers', () => {
  const elements = new Map();
  const document = {
    addEventListener() {},
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, { addEventListener() {}, innerHTML: '' });
      return elements.get(id);
    },
  };
  const context = vm.createContext({ document });
  vm.runInContext(script, context);
  vm.runInContext(`
    previewRows = [
      {'体积': 130.54, '宽': 5.9, '长': 7.5, '序号': 20, '耳标号': 93},
      {'体积': 130.54, '宽': 5.9, '长': 7.5, '序号': 29, '耳标号': 90},
      {'体积': 129.51, '宽': 5.8, '长': 7.7, '序号': 18, '耳标号': 73}
    ];
    previewColumns = ['序号', '耳标号', '长', '宽', '体积'];
    requiredCount = 2;
    selectedStart = 0;
    renderPreviewTable();
  `, context);
  const table = elements.get('preview-table').innerHTML;
  assert.match(table, /<thead><tr><th>序号<\/th><th>原始序号<\/th><th>耳标号<\/th><th>长<\/th><th>宽<\/th><th>体积<\/th><\/tr>/);
  assert.deepEqual([...table.matchAll(/<tr[^>]*><td>(\d+)<\/td><td>(\d+)<\/td>/g)].map(m => [Number(m[1]), Number(m[2])]), [[1, 20], [2, 29], [3, 18]]);
});

test('preview keeps Excel columns in order without a source sequence', () => {
  const elements = new Map();
  const document = {
    addEventListener() {},
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, { addEventListener() {}, innerHTML: '' });
      return elements.get(id);
    },
  };
  const context = vm.createContext({ document });
  vm.runInContext(script, context);
  vm.runInContext(`
    previewRows = [{'体积': 130.54, '宽': 5.9, '长': 7.5, '耳标号': 93}];
    previewColumns = ['耳标号', '长', '宽', '体积'];
    requiredCount = 1;
    renderPreviewTable();
  `, context);
  assert.match(elements.get('preview-table').innerHTML, /<thead><tr><th>序号<\/th><th>耳标号<\/th><th>长<\/th><th>宽<\/th><th>体积<\/th><\/tr>/);
});

test('group presets offer two through eight letter groups and fill the name input', () => {
  const presets = [...html.matchAll(/data-group-preset="([A-H,]+)"/g)].map(match => match[1]);
  assert.deepEqual(presets, Array.from({length: 7}, (_, index) => 'ABCDEFGH'.slice(0, index + 2).split('').join(',')));

  const elements = new Map();
  const document = {
    addEventListener() {},
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, { addEventListener() {}, value: '', hidden: true, style: {}, innerHTML: '', querySelectorAll() { return []; }, setAttribute() {} });
      return elements.get(id);
    },
  };
  const context = vm.createContext({ document });
  vm.runInContext(script, context);
  vm.runInContext("chooseGroupPreset('A,B,C,D')", context);
  assert.equal(elements.get('group-names').value, 'A,B,C,D');
  assert.equal(elements.get('group-preset-menu').hidden, true);
  assert.equal(elements.get('group-size-section').style.display, 'block');
  assert.match(elements.get('group-size-grid').innerHTML, /data-group-name="D"/);
});

test('custom counts override the default while blank fields keep it', () => {
  const elements = new Map();
  const inputs = [
    {dataset: {groupName: 'A'}, value: '4'},
    {dataset: {groupName: 'B'}, value: ''},
    {dataset: {groupName: 'C'}, value: '3'},
  ];
  const document = {
    addEventListener() {},
    querySelectorAll() { return inputs; },
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, {addEventListener() {}, value: '', setAttribute() {}});
      return elements.get(id);
    },
  };
  const context = vm.createContext({document});
  vm.runInContext(script, context);
  document.getElementById('group-num').value = '2';
  document.getElementById('group-names').value = 'A,B,C';
  assert.equal(JSON.stringify(vm.runInContext('collectGroupSizes()', context)), '{"A":4,"C":3}');
  inputs[0].value = '0';
  assert.throws(() => vm.runInContext('collectGroupSizes()', context), /正整数/);
});
