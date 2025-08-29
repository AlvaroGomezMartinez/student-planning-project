#!/usr/bin/env node
// split-students.js
// Split a batch PDF into one PDF per student using an ID regex found on the first page
// Usage: node split-students.js input.pdf out_dir [idRegex]

const fs = require('fs');
const path = require('path');
const { PDFDocument } = require('pdf-lib');
const mkdirp = require('mkdirp');
// pdfjs legacy build is ESM .mjs; we'll import it dynamically at runtime so this CJS file works
let pdfjsLib = null;

async function ensurePdfjs() {
  if (pdfjsLib) return pdfjsLib;
  // Resolve the installed package path to the legacy build .mjs
  const pkgPath = require.resolve('pdfjs-dist/package.json');
  const pkgDir = require('path').dirname(pkgPath);
  const mjsPath = `file://${pkgDir}/legacy/build/pdf.mjs`;
  pdfjsLib = await import(mjsPath);
  return pdfjsLib;
}

async function extractPageTextPdfjs(srcBytes, pageNum, loadingTask) {
  const pdf = await loadingTask.promise;
  const page = await pdf.getPage(pageNum);
  const content = await page.getTextContent();
  const pageText = content.items.map(i => i.str || '').join(' ');
  return { pageText, page };
}

// Default regex matches patterns like: "Student ID: 123456" and captures the digits
async function splitById(inputPath, outDir, idRegexStr = 'Student ID:\\s*(\\d{6})', prefix = 'EPT') {
  const idRe = new RegExp(idRegexStr, 'i');
  mkdirp.sync(outDir);

  const srcBytes = fs.readFileSync(inputPath);
  // pdf-lib accepts Buffer, but pdfjs expects a Uint8Array
  const srcUint8 = new Uint8Array(srcBytes);
  const srcPdf = await PDFDocument.load(srcBytes);
  const totalPages = srcPdf.getPageCount();

  const pdfjs = await ensurePdfjs();
  const loadingTask = pdfjs.getDocument({ data: srcUint8 });
  // loadingTask.promise resolves to a PDFDocumentProxy

  // Map to accumulate pages per cleaned student id
  const idToPdf = Object.create(null);
  let lastSeenId = null;
  const idCounts = Object.create(null);

  for (let i = 1; i <= totalPages; i++) {
    const { pageText } = await extractPageTextPdfjs(srcBytes, i, loadingTask);
    const match = pageText.match(idRe);

    let thisId = null;
    if (match) {
      const rawId = (match[1] || match[0]).toString().trim();
      // sanitize: prefer the first digit sequence if present, else strip unsafe chars
      const digitMatch = rawId.match(/\d+/);
      if (digitMatch) thisId = digitMatch[0];
      else thisId = rawId.replace(/[^A-Za-z0-9_-]/g, '').replace(/^_+/, '');
      lastSeenId = thisId;
    } else {
      thisId = lastSeenId;
    }

    if (!thisId) {
      thisId = 'unknown';
      lastSeenId = thisId;
    }

    // ensure a PDFDocument exists for this id
    if (!idToPdf[thisId]) {
      idToPdf[thisId] = await PDFDocument.create();
    }

    const targetPdf = idToPdf[thisId];
    const [copied] = await targetPdf.copyPages(srcPdf, [i - 1]);
    targetPdf.addPage(copied);
  }

  // Save all accumulated PDFs
  for (const idKey of Object.keys(idToPdf)) {
    const doc = idToPdf[idKey];
    const outBytes = await doc.save();
    const safeName = uniqueNameForId(idKey, idCounts, prefix);
    fs.writeFileSync(path.join(outDir, safeName), outBytes);
  }

  // destroy pdfjs loading task to free resources
  try { loadingTask.destroy(); } catch (e) { /* ignore */ }

  return { outDir };
}

function uniqueNameForId(id, counts, prefix = 'EPT') {
  // Clean the id: prefer the first digit sequence; otherwise strip leading underscores and unsafe chars
  function sanitizeKey(raw) {
    const s = String(raw || '').trim();
    const d = s.match(/\d+/);
    if (d) return d[0];
    const cleaned = s.replace(/^_+/, '').replace(/[^A-Za-z0-9_-]/g, '');
    return cleaned || 'unknown';
  }

  const key = sanitizeKey(id);
  if (!counts[key]) {
    counts[key] = 1;
    return `${prefix}_${key}.pdf`;
  }
  counts[key] += 1;
  return `${prefix}_${key}.${counts[key]}.pdf`;
}

// CLI
if (require.main === module) {
  const [,, inputPath, outDir, idRegex, prefixArg] = process.argv;
  if (!inputPath || !outDir) {
    console.error('Usage: node split-students.js input.pdf out_dir [idRegex] [prefix]');
    process.exit(2);
  }
  const prefix = prefixArg || 'EPT';
  splitById(inputPath, outDir, idRegex, prefix).then(() => {
    console.log('Split complete ->', outDir);
  }).catch(err => {
    console.error('Error while splitting:', err);
    process.exit(1);
  });
}

module.exports = { splitById };
