// Read-only visual check of independently parsed native output. The renderer never rewrites the .xls.
import fs from 'node:fs/promises';import path from 'node:path';import {createRequire} from 'node:module';import {pathToFileURL} from 'node:url';
const run=path.resolve(process.argv[2]),require=createRequire(path.join(run,'runtime-entry.cjs'));
const {Workbook}=await import(pathToFileURL(require.resolve('@oai/artifact-tool')).href);
const data=JSON.parse(await fs.readFile(path.join(run,'native-view.json'),'utf8'));
const corpus=JSON.parse(await fs.readFile(path.join(run,'image_corpus.json'),'utf8'));
const byId=new Map(corpus.flatMap(x=>x.ids.map(id=>[id,x])));const wb=Workbook.create();
const col=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const parts=a=>[Number(a.match(/\d+/)[0]),[...a.match(/[A-Z]+/)[0]].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1];
for(const s of data.sheets){
  const sheet=wb.worksheets.add(s.name);if(!Object.keys(s.cells).length)continue;
  const last=Math.max(...Object.keys(s.cells).map(a=>parts(a)[0])),lastcol=Math.max(...Object.keys(s.cells).map(a=>parts(a)[1]));
  sheet.getRange(`A1:${col(lastcol)}${last}`).format.font={name:'宋体',size:11,color:'#000000'};
  sheet.getRange(`A1:${col(lastcol)}${last}`).format.verticalAlignment='center';
  for(const c of s.cols){const end=Math.min(c.last,lastcol);if(c.first<=end)sheet.getRange(`${col(c.first)}1:${col(end)}${last}`).format.columnWidth=c.width;}
  for(const [r,v]of Object.entries(s.rows))if(Number(r)<=last)sheet.getRange(`A${r}:${col(lastcol)}${r}`).format.rowHeight=v.height||15;
  for(const [a,idx]of Object.entries(s.styles)){
    if(parts(a)[0]>last||parts(a)[1]>lastcol)continue;const xf=data.xfs[idx],font=data.fonts[xf.font],rg=sheet.getRange(a);
    if(font)rg.format.font={name:font.name,size:font.size,bold:font.bold,italic:font.italic,color:'#000000'};
    if([1,2,3].includes(xf.align))rg.format.horizontalAlignment=({1:'left',2:'center',3:'right'})[xf.align];
    rg.format.wrapText=xf.wrap;let fmt=data.formats[xf.format];if([14,31,176].includes(xf.format))fmt='yyyy"年"m"月"d"日"';if(fmt)rg.setNumberFormat(fmt);
    if(Buffer.from(xf.hex,'hex').readUInt32LE(10)&0xffff)rg.format.borders={preset:'all',style:'thin',color:'#000000'};
  }
  for(const m of s.merges)sheet.mergeCells(m);
  for(const [a,v]of Object.entries(s.cells)){
    const id=typeof v==='string'?v.match(/^=DISPIMG\("(ID_[^"]+)"/)?.[1]:null;
    if(!id){sheet.getRange(a).values=[[v]];continue;}
    const im=byId.get(id);if(!im)throw Error('Missing image '+id);const [r,c]=parts(a);
    const height=(s.rows[String(r)]?.height||15)*4/3;const width=(s.cols.find(x=>x.first<=c&&x.last>=c)?.width||8.43)*7+5;
    const scale=Math.min((width-4)/im.width,(height-4)/im.height),w=im.width*scale,h=im.height*scale;
    const buf=await fs.readFile(path.join(run,im.path)),mime=im.path.endsWith('.png')?'image/png':'image/jpeg';
    sheet.images.add({dataUrl:`data:${mime};base64,${buf.toString('base64')}`,anchor:{from:{row:r-1,col:c,rowOffsetPx:2,colOffsetPx:(width-w)/2},extent:{widthPx:w,heightPx:h}}});
  }
}
const out=path.join(run,'native-previews');await fs.mkdir(out,{recursive:true});
for(const s of data.sheets){if(s.name==='Sheet3')continue;const image=s.name.endsWith('截图');const range=image?(s.name==='福利费支付截图'?'A1:F4':s.name==='办公费支付截图'?'A1:E4':'A1:B3'):'A1:E12';const blob=await wb.render({sheetName:s.name,range,scale:1,format:'png'});await fs.writeFile(path.join(out,s.name+'.png'),new Uint8Array(await blob.arrayBuffer()));}
console.log(JSON.stringify({rendered:10,path:out}));
