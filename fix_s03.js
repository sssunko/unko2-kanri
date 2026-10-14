const sharp = require('sharp');
const fs = require('fs');
const base = 'C:/gas/smart-simple-system/manual-deploy/deploy-assets/screenshots/';

async function gen(inFile, outFile, toErase, toAdd) {
  const meta = await sharp(inFile).metadata();
  const { width, height, channels } = meta;
  const raw = await sharp(inFile).raw().toBuffer();

  // Pass1: 赤系ピクセルを白に置換し変更箇所を記録
  const changed = new Uint8Array(width * height);
  for (let idx=0, pi=0; pi<raw.length; pi+=channels, idx++) {
    const r=raw[pi], g=raw[pi+1], b=raw[pi+2];
    if (r>150 && r>g+30 && r>b+30) {
      changed[idx]=1;
      raw[pi]=255; raw[pi+1]=255; raw[pi+2]=255; if(channels>3) raw[pi+3]=255;
    }
  }

  // Pass2: 水平ボーダー修復
  // クリーン参照列: x=445〜520 (経費合計列, フレーム外)
  const cx1=445, cx2=520;
  for (let y=0; y<height; y++) {
    let nonWhite=0, sumR=0, sumG=0, sumB=0;
    for (let x=cx1; x<=cx2; x++) {
      const pi=(y*width+x)*channels;
      const r=raw[pi], g=raw[pi+1], b=raw[pi+2];
      if (r<230 || g<230 || b<230) { nonWhite++; sumR+=r; sumG+=g; sumB+=b; }
    }
    if (nonWhite > (cx2-cx1)*0.3) { // 30%以上がボーダーピクセル → 修復
      const ar=Math.round(sumR/nonWhite), ag=Math.round(sumG/nonWhite), ab=Math.round(sumB/nonWhite);
      for (const f of toErase) {
        for (let x=Math.max(0,f.x-3); x<=Math.min(width-1,f.x+f.w+3); x++) {
          const idx=y*width+x;
          if (changed[idx]) {
            const pi=idx*channels;
            raw[pi]=ar; raw[pi+1]=ag; raw[pi+2]=ab; if(channels>3) raw[pi+3]=255;
          }
        }
      }
    }
  }

  // クリーン画像 → 枠追加 → 保存
  const cleanPng=outFile+'.clean.png';
  await sharp(raw,{raw:{width,height,channels}}).toFile(cleanPng);
  const rects=toAdd.map(f=>
    `<rect x="${f.x}" y="${f.y}" width="${f.w}" height="${f.h}" fill="none" stroke="${f.color||'red'}" stroke-width="3" rx="2"/>`
  ).join('');
  const svg=`<svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">${rects}</svg>`;
  const tmp=outFile+'.tmp.png';
  await sharp(cleanPng).composite([{input:Buffer.from(svg),top:0,left:0}]).toFile(tmp);
  fs.unlinkSync(cleanPng); fs.renameSync(tmp,outFile);
  console.log('ok ->',outFile.replace(base,''));
}

// s03 フレーム座標 (クロップ後 946×60)
// 燃料代Z:x=276 / 支払いAA:x=359 / 利益AC:x=525, 全てy=25,w=83,h=28
const s03 = base+'s03-06-shuukei-kyuryo.png';

(async()=>{
  await gen(s03, base+'s03-crop.png',
    [{x:359,y:25,w:83,h:28},{x:525,y:25,w:83,h:28}], // 消す: 支払い・利益
    [
      {x:27,  y:25, w:83, h:28, color:'#f5c400'}, // 距離
      {x:110, y:25, w:83, h:28, color:'#f5c400'}, // 燃費
      {x:193, y:25, w:83, h:28, color:'#f5c400'}, // ガソリン代
      {x:276, y:25, w:83, h:28, color:'red'},       // 燃料代
    ]
  );
})().catch(e=>{console.error(e);process.exit(1);});
