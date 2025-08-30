#!/usr/bin/env node
/*
 * clean-student-folders.js
 * Usage:
 *   node clean-student-folders.js "<rootPath>" [--yes] [--ext pdf] [--dry-run]
 *
 * By default the script runs in dry-run mode and will only print which files
 * it would delete. Use --yes to actually delete files. Use --ext to restrict
 * deletions to a specific extension (e.g. pdf).
 */

const fs = require('fs').promises;
const path = require('path');

async function cleanStudentFolders(rootPath, options = {}) {
  const { dryRun = true, ext = null } = options;
  const summary = { foldersChecked: 0, filesFound: 0, filesDeleted: 0 };

  async function safeReaddir(p) {
    try {
      return await fs.readdir(p, { withFileTypes: true });
    } catch (err) {
      throw new Error(`Failed to read directory ${p}: ${err.message}`);
    }
  }

  // Validate root
  const absRoot = path.resolve(rootPath);
  let stat;
  try {
    stat = await fs.stat(absRoot);
  } catch (err) {
    throw new Error(`Root path does not exist: ${absRoot}`);
  }
  if (!stat.isDirectory()) {
    throw new Error(`Root path is not a directory: ${absRoot}`);
  }

  const entries = await safeReaddir(absRoot);
  const folders = entries.filter(e => e.isDirectory()).map(d => path.join(absRoot, d.name));

  for (const folder of folders) {
    summary.foldersChecked++;
    let innerEntries = await safeReaddir(folder);
    // files only (ignore directories)
    let files = innerEntries.filter(e => e.isFile()).map(f => f.name);

    if (ext) {
      const wanted = ext.replace(/^\./, '').toLowerCase();
      files = files.filter(fn => path.extname(fn).toLowerCase().replace('.', '') === wanted);
    }

    if (files.length === 0) {
      // nothing to delete
      continue;
    }

    for (const fileName of files) {
      summary.filesFound++;
      const filePath = path.join(folder, fileName);
      if (dryRun) {
        console.log(`[DRY-RUN] Would delete: ${filePath}`);
      } else {
        try {
          await fs.unlink(filePath);
          summary.filesDeleted++;
          console.log(`Deleted: ${filePath}`);
        } catch (err) {
          console.error(`Failed to delete ${filePath}: ${err.message}`);
        }
      }
    }
  }

  return summary;
}

// CLI wrapper
if (require.main === module) {
  (async () => {
    const argv = process.argv.slice(2);
    if (argv.length === 0) {
      console.error('Usage: node clean-student-folders.js "<rootPath>" [--yes] [--ext pdf] [--dry-run]');
      process.exit(2);
    }
    const rootPath = argv[0];
    const dryRun = !argv.includes('--yes') && !argv.includes('-y');
    const extFlagIndex = argv.findIndex(a => a === '--ext');
    let ext = null;
    if (extFlagIndex !== -1 && argv[extFlagIndex + 1]) ext = argv[extFlagIndex + 1];
    const start = Date.now();
    try {
      console.log(`Root: ${path.resolve(rootPath)}`);
      console.log(`Mode: ${dryRun ? 'DRY-RUN (no files will be deleted)' : 'DELETE'}`);
      if (ext) console.log(`Extension filter: ${ext}`);
      const result = await cleanStudentFolders(rootPath, { dryRun, ext });
      const elapsed = ((Date.now() - start) / 1000).toFixed(2);
      console.log('---');
      console.log(`Folders checked: ${result.foldersChecked}`);
      console.log(`Files found matching criteria: ${result.filesFound}`);
      console.log(`Files deleted: ${result.filesDeleted}`);
      console.log(`Elapsed: ${elapsed}s`);
      if (dryRun) console.log('Dry-run mode: no files were deleted. Rerun with --yes to perform deletions.');
    } catch (err) {
      console.error('Error:', err.message);
      process.exit(1);
    }
  })();
}

module.exports = { cleanStudentFolders };
