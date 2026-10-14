const sharp = require('sharp');
const fs = require('fs');
const base = 'C:/gas/smart-simple-system/manual-deploy/deploy-assets/screenshots/';
const tmp  = 'C:/gas/smart-simple-system/';

// 売上が見える元画像（ユーザー貼付のスプレッドシート行スクショ）
const src_kyuryo = tmp + 'tmp_chk2496_0.png'; // 1846x60 給料方式
const src_pct    = tmp + 'tmp_chk2496_3.png'; // 1847x54 %方式

const Y = '#f5c400';

// ============================================================
// 列セパレーター位置（原画像基準・83px均等）
//   降完時刻:513-596, 売上:596-679, 請求:679-762, 実費:762-845,
//   合計:845-928, 距離:928-1011, 燃費:1011-1094,
//   ガソリン代:1094-1177, 燃料代:1177-1260, 支払い:1260-1343,
//   経費合計:1343-1426, 利益:1426-1509
// クロップ: x=505 から幅 1010px
// ============================================================
const CROP_X = 505;
const CROP_W = 1010;
const col = (origX) => origX - CROP_X;

// ソース画像の既存黄枠(旧アノテーション)を白に消去してRAWバッファを返す
async function eraseOldYellow(srcFile) {
  const meta = await sharp(srcFile).metadata();
  const { width, height, channels } = meta;
  const raw = await sharp(srcFile).raw().toBuffer();
  for (let pi = 0; pi < raw.length; pi += channels) {
    const r = raw[pi], g = raw[pi+1], b = raw[pi+2];
    if (r > 240 && g > 200 && g < 225 && b < 20) {
      raw[pi] = 255; raw[pi+1] = 255; raw[pi+2] = 255;
      if (channels > 3) raw[pi+3] = 255;
    }
  }
  return { raw, width, height, channels };
}

async function gen(srcFile, outFile, rects) {
  const { raw, width, height, channels } = await eraseOldYellow(srcFile);
  const cw = CROP_W, ch = height;

  // 既存黄枠消去済みのクロップを一時ファイルに書き出す
  const cleanFile = outFile + '.clean.png';
  await sharp(raw, { raw: { width, height, channels } })
    .extract({ left: CROP_X, top: 0, width: CROP_W, height: ch })
    .toFile(cleanFile);

  const svgRects = rects.map(f =>
    `<rect x="${f.x}" y="${f.y}" width="${f.w}" height="${f.h}" ` +
    `fill="none" stroke="${f.color || 'red'}" stroke-width="3" rx="2"/>`
  ).join('');
  const svg = `<svg width="${cw}" height="${ch}" xmlns="http://www.w3.org/2000/svg">${svgRects}</svg>`;

  const tmp2 = outFile + '.tmp.png';
  await sharp(cleanFile)
    .composite([{ input: Buffer.from(svg), top: 0, left: 0 }])
    .toFile(tmp2);
  fs.unlinkSync(cleanFile);
  fs.renameSync(tmp2, outFile);
  console.log('ok ->', outFile.replace(base, ''));
}

(async () => {
  // フレーム（セパレーター左端+4px, w=75, 左右4px余白）
  // y=4, h=51(60px高)/h=45(54px高)：stroke外縁が画像上下端に達しない
  const makeFrames = (h) => [
    { x: col(596)  + 1, y: 4, w: 81, h, color: Y },     // 売上
    { x: col(1177) + 1, y: 4, w: 81, h, color: Y },     // 燃料代
    { x: col(1260) + 1, y: 4, w: 81, h, color: Y },     // 支払い
    { x: col(1343) + 1, y: 4, w: 81, h, color: Y },     // 経費合計
    { x: col(1426) + 1, y: 4, w: 81, h, color: 'red' }, // 利益
  ];

  await gen(src_kyuryo, base + 's05-kyuryo.png', makeFrames(51));
  await gen(src_pct,    base + 's05-pct.png',    makeFrames(45));
})().catch(e => { console.error(e); process.exit(1); });
