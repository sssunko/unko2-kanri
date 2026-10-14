const sharp = require('sharp');
const fs = require('fs');
const base = 'C:/gas/smart-simple-system/manual-deploy/deploy-assets/screenshots/';

// フレームを消去: raw pixelバッファで内側からサンプリングして枠を上書き
async function eraseFrames(inFile, outFile, frames) {
  const img = sharp(inFile);
  const meta = await img.metadata();
  const { width, height, channels } = meta;
  const raw = await img.raw().toBuffer();

  for (const f of frames) {
    const bw = 4; // 消去幅（stroke-width=3 + 1px余裕）
    // 内側サンプル色（左上角の内側）
    const sx = f.x + bw + 4;
    const sy = f.y + bw + 4;
    const si = (sy * width + sx) * channels;
    const r = raw[si], g = raw[si+1], b = raw[si+2], a = channels>3 ? raw[si+3] : 255;

    const x1 = f.x - bw, y1 = f.y - bw;
    const x2 = f.x + f.w + bw, y2 = f.y + f.h + bw;

    // 上辺・下辺
    for (let y = Math.max(0,y1); y <= Math.min(height-1, f.y+bw); y++) {
      for (let x = Math.max(0,x1); x <= Math.min(width-1,x2); x++) {
        const i = (y*width+x)*channels;
        raw[i]=r; raw[i+1]=g; raw[i+2]=b; if(channels>3) raw[i+3]=a;
      }
    }
    for (let y = Math.max(0, f.y+f.h-bw); y <= Math.min(height-1,y2); y++) {
      for (let x = Math.max(0,x1); x <= Math.min(width-1,x2); x++) {
        const i = (y*width+x)*channels;
        raw[i]=r; raw[i+1]=g; raw[i+2]=b; if(channels>3) raw[i+3]=a;
      }
    }
    // 左辺・右辺
    for (let y = Math.max(0,y1); y <= Math.min(height-1,y2); y++) {
      for (let x = Math.max(0,x1); x <= Math.min(width-1, f.x+bw); x++) {
        const i = (y*width+x)*channels;
        raw[i]=r; raw[i+1]=g; raw[i+2]=b; if(channels>3) raw[i+3]=a;
      }
    }
    for (let y = Math.max(0,y1); y <= Math.min(height-1,y2); y++) {
      for (let x = Math.max(0, f.x+f.w-bw); x <= Math.min(width-1,x2); x++) {
        const i = (y*width+x)*channels;
        raw[i]=r; raw[i+1]=g; raw[i+2]=b; if(channels>3) raw[i+3]=a;
      }
    }
  }

  const tmp = outFile + '.tmp.png';
  await sharp(raw, { raw: { width, height, channels } }).toFile(tmp);
  fs.renameSync(tmp, outFile);
  console.log('erased ->', outFile.replace(base,''));
}

// フレームを追加（SVGオーバーレイ）
async function addFrames(inFile, outFile, boxes) {
  const meta = await sharp(inFile).metadata();
  const { width, height } = meta;
  const svgRects = boxes.map(b =>
    `<rect x="${b.x}" y="${b.y}" width="${b.w}" height="${b.h}" fill="none" stroke="red" stroke-width="3" rx="2"/>`
  ).join('');
  const svg = `<svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">${svgRects}</svg>`;
  const tmp = outFile + '.tmp.png';
  await sharp(inFile).composite([{ input: Buffer.from(svg), top:0, left:0 }]).toFile(tmp);
  fs.renameSync(tmp, outFile);
  console.log('added frames ->', outFile.replace(base,''));
}

// s03 (946×60): 燃料代(x=276,y=25,w=83,h=28), 支払い(x=359), 利益(x=525)
// s04 (947×54): 支払い(x=358,y=22,w=83,h=28), 利益(x=524)
// s06p (653×73): %(x=360,y=24,w=102,h=48), 仮日数(x=155,y=24,w=101,h=48) ← 追加

(async () => {
  const s03 = base + 's03-06-shuukei-kyuryo.png';
  const s04 = base + 's04-05-shuukei-pct.png';

  // カード03: 燃料代のみ → 支払い・利益を消去
  await eraseFrames(s03, base + 's03-crop.png', [
    { x:359, y:25, w:83, h:28 }, // 支払い
    { x:525, y:25, w:83, h:28 }, // 利益
  ]);

  // カード04 給料方式: 支払いのみ → 燃料代・利益を消去
  await eraseFrames(s03, base + 's04-kyuryo.png', [
    { x:276, y:25, w:83, h:28 }, // 燃料代
    { x:525, y:25, w:83, h:28 }, // 利益
  ]);

  // カード04 %方式: 支払いのみ → 利益を消去
  await eraseFrames(s04, base + 's04-pct.png', [
    { x:524, y:22, w:83, h:28 }, // 利益
  ]);

  // カード05 給料方式: 利益のみ → 燃料代・支払いを消去
  await eraseFrames(s03, base + 's05-kyuryo.png', [
    { x:276, y:25, w:83, h:28 }, // 燃料代
    { x:359, y:25, w:83, h:28 }, // 支払い
  ]);

  // カード05 %方式: 利益のみ → 支払いを消去
  await eraseFrames(s04, base + 's05-pct.png', [
    { x:358, y:22, w:83, h:28 }, // 支払い
  ]);

  // カード06 %方式: 仮日数枠を追加（元ファイルを上書き）
  await addFrames(base + 's06-master-pct.png', base + 's06-master-pct.png', [
    { x:155, y:24, w:101, h:48 }, // 仮日数(N)
  ]);

  console.log('全完了');
})().catch(e => { console.error(e); process.exit(1); });
