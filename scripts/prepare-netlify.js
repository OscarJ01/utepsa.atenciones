const fs = require('node:fs');
const path = require('node:path');

const projectRoot = path.resolve(__dirname, '..');
const outputDirectory = path.join(projectRoot, 'dist');
const htmlFiles = fs.readdirSync(projectRoot).filter((file) => file.endsWith('.html'));
const assets = [
  'app.js',
  'styles.css',
  'logo-utepsa.png',
  'tailwind.config.js'
];

if (htmlFiles.length !== 1) {
  throw new Error(`Expected one main HTML file in the project root; found ${htmlFiles.length}.`);
}

for (const asset of assets) {
  if (!fs.existsSync(path.join(projectRoot, asset))) {
    throw new Error(`Required Netlify asset is missing: ${asset}`);
  }
}

fs.mkdirSync(outputDirectory, { recursive: true });
fs.copyFileSync(path.join(projectRoot, htmlFiles[0]), path.join(outputDirectory, 'index.html'));

for (const asset of assets) {
  fs.copyFileSync(path.join(projectRoot, asset), path.join(outputDirectory, asset));
}
