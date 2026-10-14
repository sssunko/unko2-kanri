const sharp = require('sharp');
const fs = require('fs');
const base = 'C:/gas/smart-simple-system/manual-deploy/deploy-assets/screenshots/';

async function gen(inFile, outFile, toErase, toAdd, cleanX1, cleanX2) {
  const meta = await sharp(inFile).metadata();
  const { width, height, channels } = meta;
  const raw = await sharp(inFile).raw().toBuffer();

  // Pass1: 赤系ピクセルを白に置換・変更箇所を記録
  const changed = new Uint8Array(width * height);
  for (let idx=0, pi=0; pi<raw.length; pi+=channels, idx++) {
    const r=raw[pi], g=raw[pi+1], b=raw[pi+2];
    if (r>150 && r>g+30 && r>b+30) {
      changed[idx]=1;
      raw[pi]=255; raw[pi+1]=255; raw[pi+2]=255; if(channels>3) raw[pi+3]=255;
    }
  }

  // Pass2: 水平ボーダー修復（クリーン参照列でサンプリング）
  for (let y=0; y<height; y++) {
    let nonWhite=0, sumR=0, sumG=0, sumB=0;
    for (let x=cleanX1; x<=cleanX2; x++) {
      const pi=(y*width+x)*channels;
      const r=raw[pi], g=raw[pi+1], b=raw[pi+2];
      if (r<230 || g<230 || b<230) { nonWhite++; sumR+=r; sumG+=g; sumB+=b; }
    }
    if (nonWhite > (cleanX2-cleanX1)*0.3) {
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

const s03 = base+'s03-06-shuukei-kyuryo.png';
const s04 = base+'s04-05-shuukei-pct.png';
const Y = '#f5c400';

(async()=>{
  // s04-kyuryo.png: s03から 燃料代・利益を消去 / 仮日数・給料に黄枠・支払いに赤枠
  await gen(s03, base+'s04-kyuryo.png',
    [{x:276,y:25,w:83,h:28},{x:525,y:25,w:83,h:28}], // 消す
    [
      {x:691,y:25,w:83,h:28,color:Y}, // 仮日数AE
      {x:774,y:25,w:83,h:28,color:Y}, // 給料AF
      {x:359,y:25,w:83,h:28,color:'red'}, // 支払いAA
    ],
    445, 520 // クリーン参照列（経費合計）
  );

  // s04-pct.png: s04から 利益を消去 / %に黄枠・支払いに赤枠
  await gen(s04, base+'s04-pct.png',
    [{x:524,y:22,w:83,h:28}], // 消す
    [
      {x:690,y:22,w:83,h:28,color:Y}, // 仮日数AE
      {x:856,y:22,w:83,h:28,color:Y}, // %AG
      {x:358,y:22,w:83,h:28,color:'red'}, // 支払いAA
    ],
    444, 520 // クリーン参照列（経費合計）
  );
})().catch(e=>{console.error(e);process.exit(1);});
