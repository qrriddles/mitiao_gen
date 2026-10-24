'use strict';
const APP={version:'1.1.0-preview5',date:'2026-10-24',rows:[],orphanRows:[],warnings:[],overlaps:[],fileName:''};
const NS={p:'http://schemas.openxmlformats.org/presentationml/2006/main',a:'http://schemas.openxmlformats.org/drawingml/2006/main',r:'http://schemas.openxmlformats.org/officeDocument/2006/relationships',rel:'http://schemas.openxmlformats.org/package/2006/relationships',ct:'http://schemas.openxmlformats.org/package/2006/content-types'};
const $=id=>document.getElementById(id);
// Namespace-agnostic OOXML selectors. Browsers do not consistently match
// prefixed PowerPoint tags with ordinary CSS tag-name selectors.
function qa(el,selector){
 if(!el)return [];
 const result=[];
 for(const expression of selector.split(',')){
  const bits=expression.trim().split(/\s*>\s*/);
  let found=[el];
  bits.forEach((part,i)=>{
   const next=[];
   for(const item of found){
    const nodes=i===0?item.getElementsByTagName('*'):[...(item.children||[])];
    for(const node of nodes)if(node.localName===part)next.push(node);
   }
   found=next;
  });
  result.push(...found);
 }
 return [...new Set(result)];
}
const q=(el,selector)=>qa(el,selector)[0]||null;
const clean=s=>String(s??'').replace(/\s/gu,'');
const xml=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(d.querySelector('parsererror'))throw new Error('XML 解析失败');return d;};
const serial=d=>new XMLSerializer().serializeToString(d);
const info=s=>{$('status').textContent=s};
const escapeHtml=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function highlightChars(text,chars){const matches=new Set(chars);return [...String(text)].map(ch=>matches.has(ch)?`<span class="overlap-char">${escapeHtml(ch)}</span>`:escapeHtml(ch)).join('')}
function pptFileName(){const base=APP.fileName.replace(/\.xlsx$/i,'').replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').trim()||'mitiao';const now=new Date();const two=n=>String(n).padStart(2,'0');const stamp=`${now.getFullYear()}${two(now.getMonth()+1)}${two(now.getDate())}_${two(now.getHours())}${two(now.getMinutes())}${two(now.getSeconds())}`;return `${base}_${stamp}.pptx`}
function direct(node,tag){return [...(node?.children||[])].find(x=>x.localName===tag)||null}
function nsEl(doc,ns,name,attrs={}){const e=doc.createElementNS(ns,name);for(const [k,v] of Object.entries(attrs))e.setAttribute(k,v);return e}
function download(blob,name){const a=document.createElement('a'),u=URL.createObjectURL(blob);a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),20000)}
function layoutFor(n){return n===0?2:n<=3?3:n<=6?n:n<=8?7:n<=10?8:n<=12?9:n<=15?10:2}
function fontMu(n){return n<=4?120:n<=10?96:n<=12?84:72}
function fontMian(n){return n<=20?108:n<=24?96:84}
const MAX_EXCEL_BYTES=10*1024*1024;
const LARGE_BATCH_SLIDES=300;
const MAX_ZIP_TOTAL=40*1024*1024;
const MAX_ZIP_ENTRY=16*1024*1024;
const MAX_ZIP_SHEET=12*1024*1024;
const MAX_ZIP_SHARED=8*1024*1024;
const MAX_ZIP_ENTRIES=2000;
const MAX_RIDDLE_ROWS=1500;
async function preflightExcelZip(file){
 const tailSize=Math.min(file.size,65557);
 const tail=new DataView(await file.slice(file.size-tailSize).arrayBuffer());
 let eocd=-1;
 for(let i=tail.byteLength-22;i>=0;i--){
  if(tail.getUint32(i,true)===0x06054b50&&i+22+tail.getUint16(i+20,true)===tail.byteLength){eocd=i;break}
 }
 if(eocd<0)throw new Error('无法识别 Excel 压缩包目录。');
 const entries=tail.getUint16(eocd+10,true),directorySize=tail.getUint32(eocd+12,true),directoryAt=tail.getUint32(eocd+16,true);
 if(entries===65535||directorySize===0xffffffff||directoryAt===0xffffffff)throw new Error('不支持 ZIP64 格式的 Excel 文件。');
 if(entries>MAX_ZIP_ENTRIES||directoryAt+directorySize>file.size||directorySize>4*1024*1024)throw new Error('Excel 文件结构过大，请精简工作簿。');
 const d=new DataView(await file.slice(directoryAt,directoryAt+directorySize).arrayBuffer());
 let pos=0,total=0;
 const names=new TextDecoder('utf-8');
 for(let i=0;i<entries;i++){
  if(pos+46>d.byteLength||d.getUint32(pos,true)!==0x02014b50)throw new Error('Excel 压缩包目录损坏。');
  const compressed=d.getUint32(pos+20,true),expanded=d.getUint32(pos+24,true),nameLen=d.getUint16(pos+28,true),extra=d.getUint16(pos+30,true),comment=d.getUint16(pos+32,true);
  const flags=d.getUint16(pos+8,true),method=d.getUint16(pos+10,true);
  if(pos+46+nameLen+extra+comment>d.byteLength)throw new Error('Excel 压缩包目录损坏。');
  if(flags&1||![0,8].includes(method))throw new Error('Excel 文件使用了不支持的压缩或加密方式。');
  if(compressed===0xffffffff||expanded===0xffffffff)throw new Error('不支持 ZIP64 格式的 Excel 文件。');
  const name=names.decode(new Uint8Array(d.buffer,d.byteOffset+pos+46,nameLen));
  if(expanded>MAX_ZIP_ENTRY||((/^xl\/worksheets\/.*\.xml$/i).test(name)&&expanded>MAX_ZIP_SHEET)||
     (name==='xl/sharedStrings.xml'&&expanded>MAX_ZIP_SHARED))throw new Error('Excel 内部文件展开后过大，请精简工作簿后重试。');
  if(compressed>0&&expanded>compressed*200&&expanded>1024*1024)throw new Error('Excel 文件压缩比例异常，可能导致浏览器内存不足。');
  total+=expanded;
  if(total>MAX_ZIP_TOTAL)throw new Error('Excel 文件展开后超过 40 MB 限制，请精简工作簿。');
  pos+=46+nameLen+extra+comment;
 }
 if(pos!==d.byteLength)throw new Error('Excel 压缩包目录格式异常。');
}

function checkWorkbookFile(file){
 if(!/\.xlsx$/i.test(file.name))throw new Error('请选择 .xlsx 格式的 Excel 文件（不支持 .xls、.csv、PDF 或其他文件）。');
 if(!file.size)throw new Error('文件为空，请选择有效的 Excel 谜稿。');
 if(file.size>MAX_EXCEL_BYTES)throw new Error('Excel 文件超过 10 MB 限制，请缩小文件后重试。');
}
async function checkXlsxSignature(file){
 const bytes=new Uint8Array(await file.slice(0,4).arrayBuffer());
 if(bytes.length<4||bytes[0]!==0x50||bytes[1]!==0x4b||![0x03,0x05,0x07].includes(bytes[2])||![0x04,0x06,0x08].includes(bytes[3]))
  throw new Error('这不是有效的 .xlsx 文件，或文件已损坏。请使用 Excel 另存为 .xlsx 后重试。');
}
async function parseWorkbook(file){
 checkWorkbookFile(file);
 await checkXlsxSignature(file);
 await preflightExcelZip(file);
 let z;
 try{z=await JSZip.loadAsync(file)}catch(e){throw new Error('无法打开 Excel 文件，文件可能损坏或格式不正确。')}
 if(!z.file('[Content_Types].xml')||!z.file('xl/workbook.xml')||!z.file('xl/_rels/workbook.xml.rels'))
  throw new Error('文件缺少 Excel 工作簿的必要内容，请使用 Excel 另存为 .xlsx 后重试。');
 let shared=[];
 if(z.file('xl/sharedStrings.xml')){const d=xml(await z.file('xl/sharedStrings.xml').async('string'));shared=qa(d,'si').map(si=>qa(si,'t').map(n=>n.textContent).join(''));}
 // Identify Riddles sheet using workbook relationships rather than assuming order.
 const book=xml(await z.file('xl/workbook.xml').async('string'));
 const sheets=qa(book,'sheet');const s=sheets.find(x=>x.getAttribute('name')==='Riddles')||sheets[0];
 if(!s)throw new Error('工作簿没有工作表');
 const relDoc=xml(await z.file('xl/_rels/workbook.xml.rels').async('string'));
 const rid=s.getAttributeNS(NS.r,'id');const rel=qa(relDoc,'Relationship').find(x=>x.getAttribute('Id')===rid);
 if(!rel)throw new Error('无法找到谜稿工作表');
 const target=rel.getAttribute('Target').replace(/^\//,'');
 const sheetPath=target.startsWith('xl/')?target:'xl/'+target.replace(/^\.\//,'');
 if(!z.file(sheetPath))throw new Error('工作表文件不存在：'+sheetPath);
 const doc=xml(await z.file(sheetPath).async('string'));let rows=[];let orphanRows=[];
 for(const rr of qa(doc,'sheetData > row')){
  const excelrow=Number(rr.getAttribute('r'));if(excelrow<3)continue;
  if(rows.length+orphanRows.length>=MAX_RIDDLE_ROWS)throw new Error('谜稿超过 1500 行处理上限，请分成多个 Excel 文件。');
  const r={excelrow,A:'',B:'',C:'',D:''};
  for(const cell of qa(rr,'c')){
    const col=(cell.getAttribute('r')||'').match(/^[A-Z]+/)?.[0];if(!['A','B','C','D'].includes(col))continue;
    let v=q(cell,'v')?.textContent||'';
    if(cell.getAttribute('t')==='s')v=shared[Number(v)]??'';
    if(cell.getAttribute('t')==='inlineStr')v=qa(cell,'t').map(n=>n.textContent).join('');
    r[col]=v;
  }
  for(const c of ['B','C','D'])r[c]=clean(r[c]);
  r.A=String(r.A).trim();
  if(r.B||r.C)rows.push(r);else if(r.D)orphanRows.push(r);
 }
 rows.forEach((r,i)=>{if(!r.A)r.A=String(i+1)});
 APP.orphanRows=orphanRows;return rows;
}
function analyze(){
 APP.overlaps=[];APP.warnings=[];const seqSeen=new Set();
 for(const r of APP.orphanRows)APP.warnings.push(`Excel 第 ${r.excelrow} 行：谜面（B）和谜目（C）均为空，但谜底（D）有内容；此行不会生成幻灯片。`);
 for(const r of APP.rows){
   const intersect=[...new Set([...r.B].filter(x=>/\p{Script=Han}/u.test(x)&&r.D.includes(x)))];
   if(intersect.length)APP.overlaps.push({r,chars:intersect});
   if(!r.B||!r.C)APP.warnings.push(`Excel 第 ${r.excelrow} 行：${!r.B?'谜面':'谜目'}为空，仍可生成。`);
   if(seqSeen.has(r.A))APP.warnings.push(`题号 ${r.A} 重复，请检查。`);seqSeen.add(r.A);
 }
 const panel=$('results');panel.hidden=false;
 $('summary').innerHTML=`读取 <b>${APP.rows.length}</b> 条可生成谜条；露春疑似 <b>${APP.overlaps.length}</b> 条；其他提示 <b>${APP.warnings.length}</b> 条。`;
 $('overlaps').innerHTML=APP.overlaps.length?APP.overlaps.map(x=>`<tr><td>${x.r.excelrow}</td><td>${escapeHtml(x.r.A)}</td><td>${highlightChars(x.r.B,x.chars)}</td><td>${highlightChars(x.r.D,x.chars)}</td><td><strong>${escapeHtml(x.chars.join('、'))}</strong></td></tr>`).join(''):'<tr><td colspan="5">未发现谜面与谜底共有的汉字。</td></tr>';
 $('warnings').innerHTML=APP.warnings.length?APP.warnings.map(w=>`<li>${escapeHtml(w)}</li>`).join(''):'<li>没有其他提示。</li>';
 $('generateBtn').disabled=!APP.rows.length;$('generateGoogleBtn').disabled=!APP.rows.length;
 $('reportBtn').disabled=!(APP.rows.length||APP.orphanRows.length);
 $('warningsDetails').open=true;
 info(`已读取：${APP.fileName}。检查完成，您可以下载报告、重新上传或继续生成 PPT。`);
}
function extractOriginalFont(shape){return Number(q(shape,'lstStyle > lvl1pPr > defRPr')?.getAttribute('sz')||12000)/100;}
function setFont(shape,pt){qa(shape,'p > r > rPr,p > endParaRPr').forEach(x=>x.setAttribute('sz',String(Math.round(pt*100))))}
function tweakTitle(shape,n,layout){
 const ppr=q(shape,'txBody > p > pPr');if(ppr)ppr.setAttribute('algn',(layout===4||layout===5)&&n<=layout?'ctr':'l');
 const spc={2:110,3:145,4:57.44,5:57.44,6:160,7:160,8:145,9:145,10:120}[layout];
 if(ppr){let ln=q(ppr,'lnSpc');if(!ln){ln=nsEl(shape.ownerDocument,NS.a,'a:lnSpc');ppr.appendChild(ln)}ln.replaceChildren(nsEl(shape.ownerDocument,NS.a,'a:spcPts',{val:String(Math.round(spc*100))}));}
 const off=q(shape,'spPr > xfrm > off'),ext=q(shape,'spPr > xfrm > ext');
 if(off&&ext){let x=+off.getAttribute('x'),y=+off.getAttribute('y'),w=+ext.getAttribute('cx'),h=+ext.getAttribute('cy');
  if((layout===4||layout===5)&&n<=layout){off.setAttribute('x',String(Math.max(0,x-180000)));ext.setAttribute('cx',String(w+180000))}
  if([7,8,9,10].includes(layout)){off.setAttribute('y',String(Math.max(0,y-80000)));ext.setAttribute('cy',String(h+120000))}
 }
 if(n>=16)setFont(shape,fontMian(n));
}
function tweakCategory(shape,text){
 const n=[...text].length, body=q(shape,'txBody > bodyPr');
 if(body){body.setAttribute('anchor','b');body.setAttribute('vert',n>=10?'eaVert':'horz');body.setAttribute('rtlCol',n>=10?'1':'0');
  qa(body,'noAutofit,spAutoFit,normAutofit').forEach(x=>x.remove());body.appendChild(nsEl(shape.ownerDocument,NS.a,'a:normAutofit'));
 }
 const ppr=q(shape,'txBody > p > pPr');if(ppr)ppr.setAttribute('algn',n<=6?'ctr':'r');
 setFont(shape,fontMu(n));
}
function slideXml(layoutXml,r,show){
 const ld=xml(layoutXml),d=document.implementation.createDocument(NS.p,'p:sld');
 const root=d.documentElement;root.setAttribute('xmlns:a',NS.a);root.setAttribute('xmlns:r',NS.r);
 const cs=nsEl(d,NS.p,'p:cSld');root.appendChild(cs);const tr=nsEl(d,NS.p,'p:spTree');cs.appendChild(tr);
 const grp=nsEl(d,NS.p,'p:nvGrpSpPr');tr.appendChild(grp);grp.append(nsEl(d,NS.p,'p:cNvPr',{id:'1',name:''}),nsEl(d,NS.p,'p:cNvGrpSpPr'),nsEl(d,NS.p,'p:nvPr'));
 const gp=nsEl(d,NS.p,'p:grpSpPr');tr.appendChild(gp);const xf=nsEl(d,NS.a,'a:xfrm');gp.appendChild(xf);
 for(const [tag,attrs] of [['off',{x:0,y:0}],['ext',{cx:0,cy:0}],['chOff',{x:0,y:0}],['chExt',{cx:0,cy:0}]])xf.appendChild(nsEl(d,NS.a,'a:'+tag,attrs));
 const layout=layoutFor([...r.B].length);
 for(const old of qa(ld,'spTree > sp')){
  const ph=q(old,'nvSpPr > nvPr > ph');if(!ph)continue;
  let key=ph.getAttribute('type')==='title'?'B':ph.getAttribute('idx')==='13'?'A':ph.getAttribute('idx')==='14'?'C':null;
  if(!key||(key==='A'&&!show))continue;
  const shape=d.importNode(old,true);
  const t=q(shape,'txBody > p > r > t');if(t)t.textContent=r[key];
  if(key==='B')tweakTitle(shape,[...r.B].length,layout);
  if(key==='C')tweakCategory(shape,r.C);
  qa(shape,'creationId').forEach(x=>x.remove());
  tr.appendChild(shape);
 }
 const cm=nsEl(d,NS.p,'p:clrMapOvr');cm.appendChild(nsEl(d,NS.a,'a:masterClrMapping'));root.appendChild(cm);
 return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+serial(d);
}
// Google Slides compatibility: ordinary horizontal text boxes with upright
// Chinese characters. Each vertical column remains a single editable text box.
// The standard PowerPoint slideXml() function is deliberately left unchanged.
function compatColumns(shape,text,baseFont,isRiddle=false){
 const chars=[...text];if(!chars.length)return [shape];
 const xfrm=q(shape,'spPr > xfrm'),off=q(xfrm,'off'),ext=q(xfrm,'ext');
 if(!off||!ext)return [shape];
 const x=Number(off.getAttribute('x')),y=Number(off.getAttribute('y'));
 const w=Number(ext.getAttribute('cx')),h=Number(ext.getAttribute('cy'));
 const EMU=12700;
 let chosen=null;
 // Google Slides uses a different visual scale. For four-character riddles,
 // favor one column. Use 140pt where the original layout has sufficient
 // vertical room, otherwise 115pt with compact line spacing.
 if(isRiddle&&chars.length===4){
  const fits140=h >= 4*140*EMU*1.05;
  const pt=fits140?140:115;
  chosen={pt,perCol:4,count:1,colW:pt*EMU*1.16};
 }
 // Five-character riddles: keep all five upright characters in one column.
 // Use the requested 115pt size and tighter paragraph spacing for Google Slides.
 if(isRiddle&&chars.length===5){
  const pt=115;
  chosen={pt,perCol:5,count:1,colW:pt*EMU*1.16};
 }
 // Nine/ten-character riddles: keep two right-to-left columns at 115pt.
 if(isRiddle&&(chars.length===9||chars.length===10)){
  const pt=115;
  chosen={pt,perCol:Math.ceil(chars.length/2),count:2,colW:pt*EMU*1.16};
 }
 if(!chosen)for(let pt=baseFont;pt>=36;pt-=2){
  const rowH=pt*EMU*1.13, colW=pt*EMU*1.16;
  const perCol=Math.max(1,Math.floor(h/rowH));
  const count=Math.ceil(chars.length/perCol);
  if(count*colW<=w*0.98){chosen={pt,perCol,count,colW};break}
 }
 if(!chosen){
  const pt=36,perCol=Math.max(1,Math.floor(h/(pt*EMU*1.13)));
  chosen={pt,perCol,count:Math.ceil(chars.length/perCol),colW:pt*EMU*1.16};
 }
 const {pt,perCol,count,colW}=chosen;
 const result=[];
 for(let i=0;i<count;i++){
  const clone=shape.cloneNode(true);
  const o=q(clone,'spPr > xfrm > off'),e=q(clone,'spPr > xfrm > ext');
  // For one-column 谜面 (including 1–5 characters), center the
  // editable column within the original riddle-text area instead of
  // right-aligning it. Multi-column placement remains unchanged.
  const columnX=isRiddle&&count===1
   ? x+(w-colW)/2
   : x+w-(i+1)*colW;
  o.setAttribute('x',String(Math.round(columnX)));
  o.setAttribute('y',String(y));
  e.setAttribute('cx',String(Math.round(colW)));
  e.setAttribute('cy',String(h));
  const ph=q(clone,'nvSpPr > nvPr > ph');if(ph)ph.remove();
  // A former PowerPoint layout placeholder must become a complete ordinary
  // text box. Google Slides may discard text in a placeholder-derived shape
  // that has no explicit rectangle geometry.
  const nv=q(clone,'nvSpPr > cNvPr');if(nv){nv.setAttribute('id',String(200+i));nv.setAttribute('name',`Editable vertical column ${i+1}`)}
  const cnv=q(clone,'nvSpPr > cNvSpPr');if(cnv)cnv.setAttribute('txBox','1');
  const sppr=q(clone,'spPr');
  if(sppr&&!q(sppr,'prstGeom')){
   const rect=nsEl(clone.ownerDocument,NS.a,'a:prstGeom',{prst:'rect'});
   rect.appendChild(nsEl(clone.ownerDocument,NS.a,'a:avLst'));
   sppr.appendChild(rect);
  }
  const body=q(clone,'txBody > bodyPr');
  if(body){body.setAttribute('vert','horz');body.setAttribute('anchor','t');body.setAttribute('rtlCol','0');
   for(const a of ['lIns','rIns','tIns','bIns'])body.setAttribute(a,'0');
   qa(body,'normAutofit,spAutoFit,noAutofit').forEach(a=>a.remove());
   body.appendChild(nsEl(clone.ownerDocument,NS.a,'a:noAutofit'));
  }
  const tx=q(clone,'txBody');const original=q(tx,'p');
  if(!tx||!original)continue;
  qa(tx,'p').forEach(p=>p.remove());
  const chunk=chars.slice(i*perCol,(i+1)*perCol);
  for(const ch of chunk){
   const para=original.cloneNode(true);
   const textNode=q(para,'r > t');if(textNode)textNode.textContent=ch;
   const ppr=q(para,'pPr');
   if(ppr){ppr.setAttribute('algn','ctr');
    for(const tag of ['spcBef','spcAft'])qa(ppr,tag).forEach(e=>e.remove());
    let ln=q(ppr,'lnSpc');if(!ln){ln=nsEl(clone.ownerDocument,NS.a,'a:lnSpc');ppr.appendChild(ln)}
    ln.replaceChildren(nsEl(clone.ownerDocument,NS.a,'a:spcPts',{val:String(Math.round(pt*(isRiddle&&(chars.length===4||chars.length===5)?100:113)))}));
   }
   tx.appendChild(para);
  }
  // Explicit font and spacing on every run, rather than relying on the
  // PowerPoint placeholder's layout/master style inheritance.
  setFont(clone,pt);
  qa(clone,'p > r > rPr').forEach(rpr=>{
   rpr.setAttribute('spc','0');
   let latin=q(rpr,'latin'),ea=q(rpr,'ea');
   if(!latin){latin=nsEl(clone.ownerDocument,NS.a,'a:latin',{typeface:'KaiTi'});rpr.appendChild(latin)}
   if(!ea){ea=nsEl(clone.ownerDocument,NS.a,'a:ea',{typeface:'KaiTi'});rpr.appendChild(ea)}
  });
  result.push(clone);
 }
 return result;
}
function slideXmlGoogle(layoutXml,r,show){
 // Reuse existing layout placement, then replace only vertical text shapes.
 const d=xml(slideXml(layoutXml,r,show));
 const tree=q(d,'spTree');
 for(const shape of [...qa(tree,'sp')]){
  const ph=q(shape,'nvSpPr > nvPr > ph');
  if(!ph)continue;
  const isTitle=ph.getAttribute('type')==='title';
  const isCategory=ph.getAttribute('idx')==='14';
  const text=isTitle?r.B:isCategory?r.C:'';
  if(!isTitle&&!(isCategory&&[...text].length>=10))continue;
  const font=isTitle?([...text].length>=16?fontMian([...text].length):extractOriginalFont(shape)):fontMu([...text].length);
  const cols=compatColumns(shape,text,font,isTitle);
  for(const c of cols)tree.insertBefore(c,shape);
  shape.remove();
 }
 return serial(d);
}
async function generate(show,google=false){
 // Let JSZip decode the packaged base64 template directly (more compatible with mobile browsers).
 const z=await JSZip.loadAsync(window.TEMPLATE_BASE64,{base64:true});
 const p=xml(await z.file('ppt/presentation.xml').async('string'));
 const rel=xml(await z.file('ppt/_rels/presentation.xml.rels').async('string'));
 const ct=xml(await z.file('[Content_Types].xml').async('string'));
 const ids=q(p,'sldIdLst');if(!ids)throw new Error('模板不包含幻灯片列表');ids.replaceChildren();
 qa(rel,'Relationship').forEach(r=>{if((r.getAttribute('Type')||'').endsWith('/slide')||(r.getAttribute('Type')||'').includes('vbaProject'))r.remove()});
 qa(ct,'Override').forEach(el=>{if(el.getAttribute('PartName')==='/ppt/presentation.xml')el.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml');if(el.getAttribute('PartName')==='/ppt/slides/slide1.xml')el.remove()});
 qa(ct,'Default').forEach(el=>{if(el.getAttribute('Extension')==='bin')el.remove()});
 for(const name of ['ppt/vbaProject.bin','ppt/slides/slide1.xml','ppt/slides/_rels/slide1.xml.rels','ppt/media/image1.png','ppt/media/image2.png'])z.remove(name);
 // Remove the original VBA template's stale Explorer preview.
 // Also remove its package-level relationship, so the PPTX remains consistent.
 z.remove('docProps/thumbnail.jpeg');
 if(z.file('_rels/.rels')){
  const packageRels=xml(await z.file('_rels/.rels').async('string'));
  qa(packageRels,'Relationship').forEach(el=>{
   if((el.getAttribute('Type')||'').endsWith('/thumbnail'))el.remove();
  });
  z.file('_rels/.rels',serial(packageRels));
 }
 qa(ct,'Override').forEach(el=>{
  if(el.getAttribute('PartName')==='/docProps/thumbnail.jpeg')el.remove();
 });
 const layouts={};for(const row of APP.rows){const i=layoutFor([...row.B].length);if(!layouts[i])layouts[i]=await z.file(`ppt/slideLayouts/slideLayout${i}.xml`).async('string')}
 APP.rows.forEach((row,idx)=>{
  const i=idx+1,layout=layoutFor([...row.B].length),rid=`rId${100+i}`;
  const id=nsEl(p,NS.p,'p:sldId',{id:String(255+i)});id.setAttributeNS(NS.r,'r:id',rid);ids.appendChild(id);
  rel.documentElement.appendChild(nsEl(rel,NS.rel,'Relationship',{Id:rid,Type:NS.r+'/slide',Target:`slides/slide${i}.xml`}));
  z.file(`ppt/slides/slide${i}.xml`,(google?slideXmlGoogle(layouts[layout],row,show):slideXml(layouts[layout],row,show)));
  const lr=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${NS.r}/slideLayout" Target="../slideLayouts/slideLayout${layout}.xml"/></Relationships>`;
  z.file(`ppt/slides/_rels/slide${i}.xml.rels`,lr);
  ct.documentElement.appendChild(nsEl(ct,NS.ct,'Override',{PartName:`/ppt/slides/slide${i}.xml`,ContentType:'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'}));
 });
 z.file('ppt/presentation.xml',serial(p));z.file('ppt/_rels/presentation.xml.rels',serial(rel));z.file('[Content_Types].xml',serial(ct));
 if(z.file('docProps/app.xml')){const ap=xml(await z.file('docProps/app.xml').async('string'));const slides=q(ap,'Slides');if(slides)slides.textContent=String(APP.rows.length);z.file('docProps/app.xml',serial(ap))}
 const blob=await z.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6},mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation'});
 download(blob,google?pptFileName().replace(/\.pptx$/i,"_GoogleSlides_兼容版.pptx"):pptFileName());
}
function reportText(){
 const lines=[`灯谜露春检查报告 v${APP.version}`,`检查日期：${new Date().toLocaleString('zh-CN')}`,`源文件：${APP.fileName}`,`可生成条数：${APP.rows.length}`,`疑似露春条数：${APP.overlaps.length}`,'', '【疑似露春】'];
 for(const {r,chars} of APP.overlaps)lines.push(`Excel第${r.excelrow}行｜题号${r.A}｜重复汉字：${chars.join('、')}｜谜面：${r.B}｜谜底：${r.D}`);
 if(!APP.overlaps.length)lines.push('未发现重复汉字。');
 lines.push('','【其他提示】',...(APP.warnings.length?APP.warnings:['没有其他提示。']),'','说明：只比较谜面与谜底共有的汉字，不判断灯谜语义。');
 return lines.join('\r\n');
}
window.addEventListener('DOMContentLoaded',()=>{
 $('version').textContent=`随心 制作 · v${APP.version} · ${APP.date}`;
 $('excel').addEventListener('change',async e=>{
  const f=e.target.files?.[0];if(!f)return;info('正在读取 Excel…');APP.rows=[];APP.orphanRows=[];APP.warnings=[];APP.overlaps=[];$('results').hidden=true;$('generateBtn').disabled=true;$('generateGoogleBtn').disabled=true;$('reportBtn').disabled=true;
  try{
   APP.fileName=f.name;APP.rows=await parseWorkbook(f);analyze();
   if(APP.rows.length>LARGE_BATCH_SLIDES){
    info(`已读取 ${APP.rows.length} 条谜条。超过 ${LARGE_BATCH_SLIDES} 条，生成 PPT 可能较慢；点击「生成 PPT」前请确认。`);
   }
  }catch(err){
   info('读取失败：'+(err?.message||'文件格式不正确或读取时发生错误。')+' 您可以重新选择文件。');
   $('results').hidden=true;$('generateBtn').disabled=true;console.error(err);
  }finally{e.target.value=''}
 });
 $('reportBtn').onclick=()=>download(new Blob(['\ufeff'+reportText()],{type:'text/plain;charset=utf-8'}),`露春检查报告_v${APP.version}.txt`);
 $('generateBtn').onclick=async()=>{
  if(APP.rows.length>LARGE_BATCH_SLIDES&&!window.confirm(`将生成 ${APP.rows.length} 张幻灯片，可能需要较长时间或较多内存。确定继续吗？`))return;
  info('正在生成 PowerPoint…');$('generateBtn').disabled=true;try{await generate($('numbers').checked);info('PPT 已生成并开始下载。请在 Microsoft PowerPoint 中检查长谜面的排版。')}catch(e){info('生成失败：'+e.message);console.error(e)}finally{$('generateBtn').disabled=false}};
 $('generateGoogleBtn').onclick=async()=>{
  if(APP.rows.length>LARGE_BATCH_SLIDES&&!window.confirm(`将生成 ${APP.rows.length} 张幻灯片，可能需要较长时间。确定继续吗？`))return;
  info('正在生成 Google Slides 兼容版…');$('generateGoogleBtn').disabled=true;
  try{await generate($('numbers').checked,true);info('兼容版 PPTX 已生成。请上传至 Google Slides 并检查竖排及编辑效果。')}
  catch(e){info('兼容版生成失败：'+e.message);console.error(e)}
  finally{$('generateGoogleBtn').disabled=false}
 };
 $('recheck').onclick=()=>$('excel').click();
});
