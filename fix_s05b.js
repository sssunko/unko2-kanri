// s05-nomaster と s05-yukyu に枠を付ける
// s05-nomaster: 売上・支払いに青枠、利益に赤枠（マスタにない車 = 売上 - 支払い）
// s05-yukyu: 有休手当に黄枠、利益に赤枠

const sharp = require('sharp');
const base = 'C:/gas/smart-simple-system/manual-deploy/deploy-assets/screenshots/';

const BLUE = '#1e88e5';
const YELLOW = '#f5c400';
const RED = 'red';

async function addBorders(srcFile, outFile, rects) {
  const meta = await sharp(srcFile).metadata();
  const { width: w, height: h } = meta;

  const svgRects = rects.map(r =>
    `<rect x="${r.x}" y="${r.y}" width="${r.w}" height="${r.h}" ` +
    `fill="none" stroke="${r.color}" stroke-width="3" rx="2"/>`
  ).join('');
  const svg = `<svg width="${w}" height="${h}" xmlns="http://www.w3.org/2000/svg">${svgRects}</svg>`;

  const tmp = outFile + '.tmp.png';
  await sharp(srcFile)
    .composite([{ input: Buffer.from(svg), top: 0, left: 0 }])
    .toFile(tmp);
  const fs = require('fs');
  fs.renameSync(tmp, outFile);
  console.log('ok ->', outFile.split('/').pop());
}

(async () => {
  // s05-nomaster.png (592x40)
  // 売上: x=37-155 (text "40,000" at 70-114)
  // 支払い: x=430-503 (text "35,000" at 432-483)
  // 利益: x=547-591 (text "5,000" at 562-589)
  await addBorders(base + 's05-nomaster.png', base + 's05-nomaster.png', [
    { x: 38,  y: 2, w: 116, h: 36, color: BLUE },   // 売上
    { x: 431, y: 2, w: 71,  h: 36, color: BLUE },   // 支払い
    { x: 548, y: 2, w: 42,  h: 36, color: RED  },   // 利益
  ]);

  // s05-yukyu.png (633x47)
  // 利益: x=493-579 (text "-9,000" at 536-576)
  // 有休手当: x=579-632 (text "9,000" at 596-627)
  await addBorders(base + 's05-yukyu.png', base + 's05-yukyu.png', [
    { x: 494, y: 2, w: 84,  h: 43, color: RED    },  // 利益
    { x: 580, y: 2, w: 51,  h: 43, color: YELLOW },  // 有休手当
  ]);
})().catch(e => { console.error(e); process.exit(1); });
