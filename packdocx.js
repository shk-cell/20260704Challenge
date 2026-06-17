const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const srcDir = path.join('C:', 'Users', 'Shin', 'AppData', 'Local', 'Temp', 'form_extracted');
const outFile = path.join('C:', 'Claude', '20260704Challenge', '신청서_CircuitLab_작성본.docx');

const zip = new JSZip();

function addDir(zip, dirPath, zipPath) {
  const entries = fs.readdirSync(dirPath);
  for (const entry of entries) {
    const fullPath = path.join(dirPath, entry);
    const zipEntry = zipPath ? zipPath + '/' + entry : entry;
    const stat = fs.statSync(fullPath);
    if (stat.isDirectory()) {
      addDir(zip, fullPath, zipEntry);
    } else {
      zip.file(zipEntry, fs.readFileSync(fullPath));
    }
  }
}

addDir(zip, srcDir, '');

zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  .then(buf => {
    fs.writeFileSync(outFile, buf);
    console.log('Done:', outFile, '|', buf.length, 'bytes');
  });
