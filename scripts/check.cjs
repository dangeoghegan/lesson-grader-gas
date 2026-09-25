const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
const files = fs.readdirSync(root);
let doGet = 0;
for (const file of files) {
  if (file.endsWith('.gs')) {
    const content = fs.readFileSync(path.join(root, file), 'utf8');
    new vm.Script(content, { filename: file });
    doGet += (content.match(/^function doGet\s*\(/gm) || []).length;
    if (file === 'Code.gs') {
      for (const name of ['Config', 'Sheets', 'AssessmentService', 'ClassroomService', 'GeminiService']) {
        if ((content.match(new RegExp(`^var ${name}\\s*=`, 'gm')) || []).length !== 1) {
          throw new Error(`${name} must be declared exactly once; the duplicated half of Code.gs may have returned.`);
        }
      }
    }
    console.log('✓ ' + file + ' parses');
  } else if (file.endsWith('.html')) {
    const content = fs.readFileSync(path.join(root, file), 'utf8');
    const scripts = content.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi);
    for (const script of scripts) {
      // GAS template tags in legacy dialogs are HTML, not plain JS. WebApp uses none.
      if (!script[1].includes('<?')) new vm.Script(script[1], { filename: file });
    }
    console.log('✓ ' + file + ' scripts parse');
  }
}
if (doGet !== 1) throw new Error(`Expected one doGet, found ${doGet}`);
console.log('✓ one doGet');
const manifest = JSON.parse(fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8'));
if (manifest.runtimeVersion !== 'V8' || !(manifest.dependencies?.enabledAdvancedServices || []).some(service => service.userSymbol === 'Classroom')) {
  throw new Error('Apps Script manifest must declare V8 and the Classroom advanced service.');
}
console.log('✓ manifest parses with Classroom service');
