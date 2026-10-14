// fix_frames2.js: 赤ピクセル検出で全枠を削除し、必要な枠だけ追加する
const sharp = require('sharp');
const fs = require('fs');
const base = 'C:/gas/smart-simple-system/manual-deploy/deploy-assets/screenshots/';

// 赤ピクセル(R>200,G<80,B<80)を検出して白(255,255,255)に置換
async function removeRedPixels(inFile) {
  const img = sharp(inFile);
  const meta = await img.metadata();
  const { width, height, channels } = meta;
  const raw = await img.raw().toBuffer();
  for (let i = 0; i < raw.length; i += channels) {
    if (raw[i] > 200 && raw[i+1] < 80 && raw[i+2] < 80) {
      raw[i] = 255; raw[i+1] = 255; raw[i+2] = 255;
      if (channels > 3) raw[i+3] = 255;
    }
  }
  return { raw, width, height, channels };
}

async function generate(inFile, outFile, frames) {
  const { raw, width, height, channels } = await removeRedPixels(inFile);
  const cleanPng = outFile + '.clean.png';
  await sharp(raw, { raw: { width, height, channels } }).toFile(cleanPng);

  const svgRects = frames.map(f =>
    `<rect x="${f.x}" y="${f.y}" width="${f.w}" height="${f.h}" fill="none" stroke="red" stroke-width="3" rx="2"/>`
  ).join('');
  const svg = `<svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">${svgRects}</svg>`;
  const tmp = outFile + '.tmp.png';
  await sharp(cleanPng).composite([{ input: Buffer.from(svg), top: 0, left: 0 }]).toFile(tmp);
  fs.unlinkSync(cleanPng);
  fs.renameSync(tmp, outFile);
  console.log('ok ->', outFile.replace(base, ''));
}

// s03 (946×60): 燃料代(x=276,y=25,w=83,h=28) / 支払い(x=359) / 利益(x=525)
// s04 (947×54): 支払い(x=358,y=22,w=83,h=28) / 利益(x=524)

(async () => {
  const s03 = base + 's03-06-shuukei-kyuryo.png';
  const s04 = base + 's04-05-shuukei-pct.png';

  // カード03: 燃料代のみ
  await generate(s03, base + 's03-crop.png', [{ x:276, y:25, w:83, h:28 }]);
  // カード04 給料: 支払いのみ
  await generate(s03, base + 's04-kyuryo.png', [{ x:359, y:25, w:83, h:28 }]);
  // カード04 %: 支払いのみ
  await generate(s04, base + 's04-pct.png', [{ x:358, y:22, w:83, h:28 }]);
  // カード05 給料: 利益のみ
  await generate(s03, base + 's05-kyuryo.png', [{ x:525, y:25, w:83, h:28 }]);
  // カード05 %: 利益のみ
  await generate(s04, base + 's05-pct.png', [{ x:524, y:22, w:83, h:28 }]);

  console.log('全完了');
})().catch(e => { console.error(e); process.exit(1); });
