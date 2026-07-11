// Usage: node scripts/write-build-env.js <resourcePrefix>
const fs = require('fs');
const path = require('path');

const resourcePrefix = process.argv[2];
if (!resourcePrefix) {
  console.error('Usage: node scripts/write-build-env.js <resourcePrefix>');
  process.exit(1);
}

const outPath = path.join(__dirname, '../dist/.build-env.json');
fs.writeFileSync(outPath, JSON.stringify({ resourcePrefix }) + '\n');
