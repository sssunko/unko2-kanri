const sharp = require('sharp');
const path = require('path');
const base = 'C:/gas/smart-simple-system/manual-deploy/deploy-assets/screenshots/';

async function addFrames(inFile, outFile, boxes) {
  const meta = await sharp(inFile).metadata();
  const { width, height } = meta;
  console.log(path.basename(inFile), width + 'x' + height);

  const svgRects = boxes.map(b =>
    `<rect x="${b.x}" y="${b.y}" width="${b.w}" height="${b.h}" fill="none" stroke="red" stroke-width="3" rx="2"/>`
  ).join('');
  const svg = `<svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">${svgRects}</svg>`;

  const tmpFile = outFile + '.tmp.png';
  await sharp(inFile)
    .composite([{ input: Buffer.from(svg), top: 0, left: 0 }])
    .toFile(tmpFile);
  const fs = require('fs');
  fs.renameSync(tmpFile, outFile);
  console.log('-> saved', path.basename(outFile));
}

// 列情報（黄色セパレーター確認済み）:
// 集計表: 黄色ライン x=927,1010,1093,1176,1259,1342,1508,1591
// 各列x（左端から）: 燃料代Z=1176, 支払いAA=1259, 経費AB=1342, 利益AC=1425, 備考AD=1508, 仮日数AE=1591, 給料AF=1674
// s03データ行y: y=25 h=28
// s04データ行y: y=22 h=28 (s04列は1px左にずれ: x=-1)
// マスタ列: 行番号列(0-40)+A(42-)+ ... N=1355, O=1455, P(s04)=1560
// マスタ データ行y: y=27 h=48

(async () => {
  // ① 集計表_給料方式: 燃料代Z・支払いAA・利益AC を赤枠
  await addFrames(
    base + 's03-06-shuukei-kyuryo.png',
    base + 's03-06-shuukei-kyuryo.png',
    [
      { x: 1176, y: 25, w: 83, h: 28 }, // 燃料代(Z)
      { x: 1259, y: 25, w: 83, h: 28 }, // 支払い(AA)
      { x: 1425, y: 25, w: 83, h: 28 }, // 利益(AC)
    ]
  );

  // ② 集計表_%方式: 支払いAA・利益AC を赤枠 (列x=s03-1)
  await addFrames(
    base + 's04-05-shuukei-pct.png',
    base + 's04-05-shuukei-pct.png',
    [
      { x: 1258, y: 22, w: 83, h: 28 }, // 支払い(AA)
      { x: 1424, y: 22, w: 83, h: 28 }, // 利益(AC)
    ]
  );

  // ③ マスタ_給料方式: 仮日数N・給料O を赤枠
  await addFrames(
    base + 's06-master-kyuryo.png',
    base + 's06-master-kyuryo.png',
    [
      { x: 1355, y: 27, w: 101, h: 48 }, // 仮日数(N)
      { x: 1455, y: 27, w: 102, h: 48 }, // 給料(O) ← blue
    ]
  );

  // ④ マスタ_%方式: %(P) を赤枠
  await addFrames(
    base + 's06-master-pct.png',
    base + 's06-master-pct.png',
    [
      { x: 1560, y: 24, w: 102, h: 48 }, // %(P) ← blue
    ]
  );
})();
