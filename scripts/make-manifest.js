#!/usr/bin/env node
// Writes manifest.xml from manifest.template.xml with the hosting address filled in.
// Usage: node scripts/make-manifest.js https://NAME.github.io/word-parallel-reader
'use strict';
const fs = require('fs');
const path = require('path');

const base = (process.argv[2] || '').replace(/\/+$/, '');
if (!/^https:\/\/[^/]+/.test(base)) {
  console.error('Give the HTTPS address of the hosted files, e.g. https://NAME.github.io/word-parallel-reader');
  process.exit(1);
}
const root = path.join(__dirname, '..');
const tpl = fs.readFileSync(path.join(root, 'manifest.template.xml'), 'utf8');
// AppDomain must be the origin only (scheme + host)
const origin = new URL(base).origin;
const out = tpl
  .replace(/<AppDomain>\{\{BASE\}\}<\/AppDomain>/, '<AppDomain>' + origin + '</AppDomain>')
  .replace(/\{\{BASE\}\}/g, base);
fs.writeFileSync(path.join(root, 'manifest.xml'), out);
console.log('manifest.xml written for ' + base);
