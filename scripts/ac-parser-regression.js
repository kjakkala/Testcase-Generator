const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const indexPath = path.join(__dirname, '..', 'public', 'index.html');
const source = fs.readFileSync(indexPath, 'utf8');

function extractFunctionCode(text, fnName) {
  const startToken = `function ${fnName}(`;
  const start = text.indexOf(startToken);
  if (start === -1) {
    throw new Error(`Function not found: ${fnName}`);
  }

  const openBrace = text.indexOf('{', start);
  if (openBrace === -1) {
    throw new Error(`Malformed function: ${fnName}`);
  }

  let i = openBrace;
  let depth = 0;
  let inSingle = false;
  let inDouble = false;
  let inTemplate = false;
  let inLineComment = false;
  let inBlockComment = false;
  let escaped = false;

  for (; i < text.length; i++) {
    const ch = text[i];
    const next = i + 1 < text.length ? text[i + 1] : '';

    if (inLineComment) {
      if (ch === '\n') inLineComment = false;
      continue;
    }
    if (inBlockComment) {
      if (ch === '*' && next === '/') {
        inBlockComment = false;
        i++;
      }
      continue;
    }

    if (!inSingle && !inDouble && !inTemplate) {
      if (ch === '/' && next === '/') {
        inLineComment = true;
        i++;
        continue;
      }
      if (ch === '/' && next === '*') {
        inBlockComment = true;
        i++;
        continue;
      }
    }

    if (inSingle || inDouble || inTemplate) {
      if (escaped) {
        escaped = false;
        continue;
      }
      if (ch === '\\') {
        escaped = true;
        continue;
      }
      if (inSingle && ch === "'") inSingle = false;
      else if (inDouble && ch === '"') inDouble = false;
      else if (inTemplate && ch === '`') inTemplate = false;
      continue;
    }

    if (ch === "'") {
      inSingle = true;
      continue;
    }
    if (ch === '"') {
      inDouble = true;
      continue;
    }
    if (ch === '`') {
      inTemplate = true;
      continue;
    }

    if (ch === '{') depth++;
    if (ch === '}') {
      depth--;
      if (depth === 0) {
        return text.slice(start, i + 1);
      }
    }
  }

  throw new Error(`Unterminated function: ${fnName}`);
}

const needed = [
  'extractACFromDescription',
  'extractDesc',
  'extractADF',
  'extractFieldText',
  'extractAC',
];

const context = vm.createContext({ console });
for (const fn of needed) {
  const code = extractFunctionCode(source, fn);
  vm.runInContext(code, context, { filename: 'public/index.html' });
}

function runTest(name, fn) {
  try {
    fn();
    console.log(`PASS: ${name}`);
  } catch (err) {
    console.error(`FAIL: ${name}`);
    console.error(err.message);
    process.exitCode = 1;
  }
}

runTest('Parses AC from same-line heading in description', () => {
  const fields = {
    description: 'Story details\nAcceptance Criteria: 1. User can save\n2. Success toast appears\nNotes\nOther text',
  };
  const out = context.extractAC(fields, {});
  assert.ok(out.includes('User can save'), 'Missing first criterion');
  assert.ok(out.includes('Success toast appears'), 'Missing second criterion');
});

runTest('Uses acceptance custom field by Jira names metadata', () => {
  const fields = {
    description: 'No AC in description',
    customfield_12345: {
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          content: [{ type: 'text', text: 'User sees status badge' }],
        },
      ],
    },
  };
  const names = { customfield_12345: 'Acceptance Criteria' };
  const out = context.extractAC(fields, names);
  assert.ok(out.includes('User sees status badge'), 'Did not extract ADF AC field');
});

runTest('Falls back to embedded acceptance text in any custom field', () => {
  const fields = {
    description: 'General description only',
    customfield_99999: 'Acceptance Criteria:\n- Preview image loads\n- Upload button enabled',
  };
  const names = { customfield_99999: 'Implementation Notes' };
  const out = context.extractAC(fields, names);
  assert.ok(out.includes('Preview image loads'), 'Missing embedded AC bullet');
  assert.ok(out.includes('Upload button enabled'), 'Missing second embedded AC bullet');
});

runTest('Returns empty when no AC is present', () => {
  const fields = { description: 'This story has no explicit checklist.' };
  const out = context.extractAC(fields, {});
  assert.strictEqual(out, '', 'Expected empty AC result');
});

if (process.exitCode && process.exitCode !== 0) {
  process.exit(process.exitCode);
}

console.log('All AC parser regression checks passed.');
