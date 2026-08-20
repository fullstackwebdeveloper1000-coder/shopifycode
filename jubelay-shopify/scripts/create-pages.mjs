#!/usr/bin/env node
/**
 * Create JUBELAY store pages via Shopify CLI.
 * Usage: node create-pages.mjs your-store.myshopify.com
 */

import { readFileSync, writeFileSync } from 'fs';
import { execSync } from 'child_process';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const store = process.argv[2];
if (!store) {
  console.error('Usage: node create-pages.mjs <store.myshopify.com>');
  process.exit(1);
}

const normalizedStore = store.replace(/\.myshopify\.com$/, '') + '.myshopify.com';
const scriptDir = dirname(fileURLToPath(import.meta.url));
const pagesPath = join(scriptDir, '..', 'import', 'pages.json');
const pages = JSON.parse(readFileSync(pagesPath, 'utf8'));

const mutation = `mutation pageCreate($page: PageCreateInput!) {
  pageCreate(page: $page) {
    page { id title handle }
    userErrors { field message }
  }
}`;

for (const page of pages) {
  const input = {
    input: {
      title: page.title,
      handle: page.handle,
      body: page.body,
      isPublished: true,
    },
  };

  const tmpFile = `/tmp/jubelay-page-${page.handle}.json`;
  writeFileSync(tmpFile, JSON.stringify(input));

  try {
    const result = execSync(
      `npx shopify store execute --store ${normalizedStore} --allow-mutations ` +
        `--query '${mutation.replace(/'/g, "'\\''")}' ` +
        `--variables "$(cat ${tmpFile})"`,
      { encoding: 'utf8', cwd: join(scriptDir, '..', '..') }
    );
    console.log(`✓ Created page: ${page.title} (${page.handle})`);
  } catch (err) {
    console.error(`✗ Failed to create page: ${page.title}`);
    if (err.stdout) console.error(err.stdout);
    if (err.stderr) console.error(err.stderr);
  }
}

console.log('\nPages import complete.');
