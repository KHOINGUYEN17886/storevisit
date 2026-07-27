const fs = require('fs');
const content = fs.readFileSync('index.html', 'utf8');
const match = content.match(/<script>([\s\S]*?)<\/script>/);
if (!match) { console.log('NO SCRIPT FOUND'); process.exit(1); }
const script = match[1];
const lines = script.split('\n');

console.log('=== SCRIPT AUDIT ===');
console.log('Script lines:', lines.length);

// Find key lines
const checks = [
  'function debugLog',
  'const storeData',
  'let activeStoreData',
  'function initDropdowns',
  'function initApp',
  'function handleModeChange',
  'function updateStoreList',
  'function switchMainMenu',
  'function loadHtmlPreviewFromForm',
  'function renderHtmlReport',
  'function printHtmlReport',
  'function loadHistoricalSubmissions',
  'function filterHistory',
];
checks.forEach(item => {
  const lineNo = lines.findIndex(l => l.includes(item));
  console.log(lineNo >= 0 ? 'Line ' + (lineNo+1) + ': ' + item : 'NOT FOUND: ' + item);
});

// Try new Function to catch JS syntax
try {
  new Function(script);
  console.log('\nJS Syntax: OK');
} catch(e) {
  console.error('\nJS SYNTAX ERROR:', e.message);
}

// Check for key HTML elements
const htmlChecks = [
  'id="asmName"',
  'id="regionSelect"',
  'id="storeCode"',
  'id="modeSelect"',
  'id="main-view-history"',
  'id="main-view-html-preview"',
  'id="menu-history-tab"',
  'id="menu-htmlpreview-tab"',
  'id="historyTableBody"',
  'id="htmlPreviewContent"',
  'onclick="switchMainMenu',
];
console.log('\n=== HTML ELEMENT AUDIT ===');
htmlChecks.forEach(item => {
  console.log(content.includes(item) ? 'FOUND: ' + item : 'MISSING: ' + item);
});

// Check storeData structure
const storeDataMatch = script.match(/const storeData\s*=\s*\{([\s\S]*?)\};\s*\n\s*let activeStoreData/);
if (storeDataMatch) {
  const hasAsms = script.includes('"asms":');
  const hasRegions = script.includes('"regions":');
  const hasMappingByAsm = script.includes('"mapping_by_asm":');
  console.log('\n=== storeData STRUCTURE ===');
  console.log('Has "asms":', hasAsms);
  console.log('Has "regions":', hasRegions);
  console.log('Has "mapping_by_asm":', hasMappingByAsm);
} else {
  console.log('\nWARNING: Cannot find storeData definition');
}
