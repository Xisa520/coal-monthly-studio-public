/* 月报图表工作台 V3.0：所有原始数值只在访问者的浏览器内解析。 */
(() => {
'use strict';
const META=window.MONTHLY_CHART_CATALOG||{sheets:[],charts:[],totals:{}};
const $=id=>document.getElementById(id);
const S={wb:null,filename:'',current:null,category:'全部',query:'',focus:null,chartData:null,sheetName:null,sheetData:null};
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt=v=>Number.isFinite(+v)?Number(v).toLocaleString('zh-CN',{maximumFractionDigits:2}):String(v??'—');
const num=v=>typeof v==='number'&&Number.isFinite(v)?v:(typeof v==='string'&&v.trim()&&Number.isFinite(+v)?+v:null);
const unique=a=>[...new Set(a)];
const colors=['#1478a6','#e18b4c','#26a6a0','#7d70ba','#97a865','#cf6481','#4587c1','#8a8f9b'];
function toDate(v){
 if(v instanceof Date&&!isNaN(+v))return v.toISOString().slice(0,10);
 if(typeof v==='number'&&v>30000&&v<80000&&window.XLSX?.SSF?.parse_date_code){const p=window.XLSX.SSF.parse_date_code(v);if(p?.y&&p?.m&&p?.d)return `${p.y}-${String(p.m).padStart(2,'0')}-${String(p.d).padStart(2,'0')}`;}
 const t=String(v??'').trim();if(/^\d{8}$/.test(t)&&/^20\d{2}/.test(t))return `${t.slice(0,4)}-${t.slice(4,6)}-${t.slice(6)}`;
 return t;
}
function idOf(addr){let m=String(addr||'').match(/^\$?([A-Za-z]+)\$?(\d+)$/);return m?{col:m[1].toUpperCase(),row:+m[2]}:null;}
function colNum(s){return [...s].reduce((v,x)=>v*26+x.charCodeAt(0)-64,0)-1;}
function colLabel(n){let s='';for(n++;n>0;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;}
function parseRef(formula){
 if(!formula||/#REF!|\[\d+\]/.test(formula))return null;
 const i=formula.lastIndexOf('!');if(i<1)return null;
 let sheet=formula.slice(0,i).trim();if(sheet[0]==="'"&&sheet.at(-1)==="'")sheet=sheet.slice(1,-1).replace(/''/g,"'");
 const addr=formula.slice(i+1).replace(/\$/g,'').split(':');
 const a=idOf(addr[0]),b=idOf(addr[1]||addr[0]);
 if(!a||!b)return null;
 return{sheet,a,b};
}
function cellVal(sh,c,r){let x=sh?.[`${c}${r}`];if(!x)return null;return x.v===undefined?null:x.v;}
function getRefValue(f){let ref=parseRef(f);if(!ref||!S.wb?.Sheets[ref.sheet])return null;return cellVal(S.wb.Sheets[ref.sheet],ref.a.col,ref.a.row);}
function readRange(f){
 const r=parseRef(f), sh=r&&S.wb?.Sheets[r.sheet];if(!r||!sh)return null;
 let c1=colNum(r.a.col),c2=colNum(r.b.col),rr1=r.a.row,rr2=r.b.row;
 if(c2<c1||rr2<rr1||((c2-c1)>5&&(rr2-rr1)>0))return null;
 let out=[],limit=5000;
 for(let row=rr1;row<=rr2&&out.length<limit;row++)for(let c=c1;c<=c2&&out.length<limit;c++)out.push(cellVal(sh,colLabel(c),row));
 return out;
}
function resolveChart(ch){
 let all=[],bad=0;
 for(let plot of ch.plots)for(let ss of plot.series){
  if(ss.broken){bad++;continue;}
  const yy=readRange(ss.y);if(!yy){bad++;continue;}
  let xx=readRange(ss.x)||[];
  const label=ss.label||getRefValue(ss.labelRef)||getRefValue(ss.y.replace(/!.*$/,'')+'!'+(parseRef(ss.y)?.a.col||'')+'2')||'指标 '+(all.length+1);
  const points=[];
  for(let i=0;i<yy.length;i++){const v=num(yy[i]);if(v===null)continue;points.push({x:toDate(xx[i]??(i+1)),v});}
  if(points.length)all.push({label:String(label).slice(0,50),type:plot.type,points});else bad++;
 }
 return{title:ch.title,series:all,missing:bad,source:ch.sourceSheet,chart:ch};
}
function svgFor(d,mini=false){
 const W=mini?540:980,H=mini?165:430,p={l:mini?28:68,r:mini?12:64,t:mini?15:32,b:mini?26:60};
 const series=(d?.series||[]).slice(0,mini?5:12);
 if(!series.length)return `<div class="lab-empty">${S.wb?'该图引用缺失、源表不存在或当前值为空。可尝试在下方“工作表自助出图”选择有效列。':'尚未加载Excel数据；上传文件后将绘制原图，并可以查看全部数值。'}</div>`;
 const pie=series.every(s=>s.type==='pie');
 const all=series.flatMap(s=>s.points.map(x=>x.v)).filter(Number.isFinite);
 if(!all.length)return '<div class="lab-empty">没有有效的数值数据。</div>';
 if(pie){const ss=series[0];let arr=ss.points.filter(p=>p.v>0).slice(0,mini?10:18);const total=arr.reduce((a,b)=>a+b.v,0)||1;const cx=mini?W*.45:W*.43,cy=H*.53,rad=mini?58:148;
  let angle=-Math.PI/2,svg=`<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="${esc(d.title)}">`;
  arr.forEach((pt,i)=>{const next=angle+pt.v/total*Math.PI*2,delta=Math.min(2*Math.PI-.00001,next-angle),end=angle+delta;const x1=cx+rad*Math.cos(angle),y1=cy+rad*Math.sin(angle),x2=cx+rad*Math.cos(end),y2=cy+rad*Math.sin(end);svg+=`<path d="M${cx} ${cy} L${x1} ${y1} A${rad} ${rad} 0 ${delta>Math.PI?1:0} 1 ${x2} ${y2} Z" fill="${colors[i%colors.length]}"><title>${esc(pt.x)} ${fmt(pt.v)}</title></path>`;angle=next;});
  if(!mini)arr.slice(0,8).forEach((p,i)=>{svg+=`<rect x="${W-278}" y="${35+i*37}" width="13" height="13" rx="3" fill="${colors[i%colors.length]}"/><text x="${W-258}" y="${47+i*37}" font-size="13" fill="#43556a">${esc(String(p.x).slice(0,18))} · ${(100*p.v/total).toFixed(1)}%</text>`;});
  return svg+'</svg>';
 }
 const keys=unique(series.flatMap(s=>s.points.map(p=>p.x)));let xs=keys.sort((a,b)=>String(a).localeCompare(String(b),'zh-CN',{numeric:true}));
 if(xs.length>280){const stride=Math.ceil(xs.length/280);xs=xs.filter((_,i)=>i%stride===0||i===xs.length-1);}
 const vals=series.flatMap(s=>s.points.map(p=>p.v));let magnitudes=series.map(s=>Math.max(...s.points.map(p=>Math.abs(p.v))));
 const veryBig=Math.max(...magnitudes)>50;
 const secund=new Set(series.map((s,i)=>veryBig&&magnitudes[i]<2&&s.points.some(p=>p.v!==0)?i:-1).filter(i=>i>=0));
 const leftVals=series.flatMap((s,i)=>secund.has(i)?[]:s.points.map(p=>p.v));const rightVals=series.flatMap((s,i)=>secund.has(i)?s.points.map(p=>p.v):[]);
 function bounds(arr){if(!arr.length)return[0,1];let lo=Math.min(...arr),hi=Math.max(...arr);if(lo===hi){hi+=Math.max(1,Math.abs(hi)*.1);lo-=Math.max(1,Math.abs(lo)*.1)}let span=hi-lo;return[Math.min(0,lo-span*.09),hi+span*.10]}
 const [min,max]=bounds(leftVals),[min2,max2]=bounds(rightVals);const pw=W-p.l-p.r,ph=H-p.t-p.b;
 const xfn=i=>p.l+(xs.length===1?pw/2:i*pw/(xs.length-1));const yfn=(v,secondary=false)=>{const a=secondary?min2:min,b=secondary?max2:max;return p.t+ph*(b-v)/(b-a)};
 let svg=`<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="${esc(d.title)}">`;
 for(let j=0;j<=4;j++){const y=p.t+ph*j/4;svg+=`<line x1="${p.l}" x2="${W-p.r}" y1="${y}" y2="${y}" stroke="#e8edf2"/><text x="${p.l-9}" y="${y+4}" text-anchor="end" fill="#7890a5" font-size="${mini?12:13}">${esc(fmt(max-(max-min)*j/4))}</text>`;
 if(rightVals.length&&!mini)svg+=`<text x="${W-p.r+9}" y="${y+4}" fill="#7890a5" font-size="12">${fmt((max2-(max2-min2)*j/4)*100)}%</text>`;}
 const xGap=Math.max(1,Math.ceil(xs.length/(mini?5:12))); const isOneYearDates=xs.length>0&&xs.every(x=>/^\d{4}-\d{2}-\d{2}$/.test(x)&&x.slice(0,4)===xs[0].slice(0,4));
 xs.forEach((x,i)=>{if(i%xGap&&i!==xs.length-1)return;svg+=`<text x="${xfn(i)}" y="${H-p.b+23}" text-anchor="middle" fill="#7c8da0" font-size="${mini?11:12}">${esc(isOneYearDates?String(x).slice(5,10):String(x).slice(0,mini?8:10))}</text>`;});
 const active=new Set(xs);
 series.forEach((s,i)=>{const mp=new Map(s.points.filter(pt=>active.has(pt.x)).map(pt=>[pt.x,pt.v]));const secondary=secund.has(i),color=colors[i%colors.length],bar=s.type==='bar',area=s.type==='area';let path='',areaPath='',segments=[];let pactive=false;
 xs.forEach((x,k)=>{const v=mp.get(x);if(!Number.isFinite(v)){pactive=false;return;}const xx=xfn(k),yy=yfn(v,secondary);if(bar){let width=Math.max(1,Math.min(13,pw/xs.length*.65/series.length));let delta=(i-(series.length-1)/2)*width*1.15;const zero=yfn(0,secondary);svg+=`<rect x="${xx+delta-width/2}" y="${Math.min(yy,zero)}" width="${width}" height="${Math.max(1,Math.abs(zero-yy))}" fill="${color}" opacity=".9"><title>${esc(s.label)} / ${esc(x)}：${fmt(v)}</title></rect>`;}else{path+=`${pactive?'L':'M'}${xx.toFixed(1)} ${yy.toFixed(1)} `;pactive=true;segments.push([xx,yy,x,v]);}});
 if(!bar&&path){if(area&&segments.length>1)svg+=`<path d="${path}L${segments.at(-1)[0]} ${yfn(0,secondary)} L${segments[0][0]} ${yfn(0,secondary)} Z" fill="${color}" opacity=".12"/>`;
 svg+=`<path d="${path}" stroke="${color}" stroke-width="${mini?2.0:2.5}" fill="none" stroke-linecap="round" stroke-linejoin="round"/>`;
 if(!mini&&xs.length<45)segments.forEach(([xx,yy,x,v])=>{svg+=`<circle cx="${xx}" cy="${yy}" r="2.6" fill="${color}"><title>${esc(s.label)} / ${esc(x)}：${fmt(v)}</title></circle>`;});}
 if(!mini&&i<7){let ox=p.l+(i%3)*255,oy=H-13-Math.floor(i/3)*18;svg+=`<rect x="${ox}" y="${oy-10}" width="10" height="10" rx="2" fill="${color}"/><text x="${ox+15}" y="${oy}" font-size="12" fill="#526579">${esc(s.label.slice(0,21))}${secondary?'(右轴)':''}</text>`;}
 });return svg+'</svg>';
}
function blobDownload(content,name,mime){const url=URL.createObjectURL(new Blob([content],{type:mime}));let a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),5000);}
const cleanName=s=>String(s||'chart').replace(/[\\/:*?"<>|]+/g,'_').slice(0,75);
function csvFor(d){const table=[['图表名称','系列','横坐标','数值'],...d.series.flatMap(s=>s.points.map(p=>[d.title,s.label,p.x,p.v]))];return '\ufeff'+table.map(row=>row.map(v=>{const text=String(v??'');const safe=(typeof v==='string'&&/^[=+\-@]/.test(text))?"'"+text:text;return '"'+safe.replaceAll('"','""')+'"'}).join(',')).join('\r\n')}
function renderTable(d){let s=d.series||[];let lines=s.flatMap(k=>k.points.map(p=>[k.label,p.x,p.v]));return `<table><thead><tr><th>指标序列</th><th>时间 / 类别</th><th>原表数值</th></tr></thead><tbody>${lines.slice(0,120).map(v=>`<tr><td>${esc(v[0])}</td><td>${esc(v[1])}</td><td>${fmt(v[2])}</td></tr>`).join('')}</tbody></table><p>预览${Math.min(120,lines.length)} / ${lines.length}条记录；可通过“导出CSV”下载该图全部记录。</p>`;}
function drawCards(){let list=META.charts.filter(ch=>(S.category==='全部'||ch.category===S.category)&&`${ch.title} ${ch.sourceSheet}`.toLowerCase().includes(S.query));
 $('labChartGrid').innerHTML=list.map(ch=>{const d=S.wb?resolveChart(ch):null;const n=d?.series?.length||0;return `<button type="button" class="lab-chart-card ${S.focus===ch.id?'chosen':''}" data-id="${esc(ch.id)}"><div class="lab-chart-card-top"><span class="lab-num">${String(ch.index).padStart(2,'0')}</span><span class="lab-type">${esc(ch.plotType)}</span><span class="lab-card-badge ${ch.broken?'bad':''}">${ch.broken?'引用待修复':S.wb?(n?'已绘制':'待匹配'):'待加载'}</span></div><h3>${esc(ch.title)}</h3><small title="${esc(ch.sourceSheet)}">${esc(ch.sourceSheet)}</small><div class="lab-chart-thumb">${d&&n?svgFor(d,true):'<div class="lab-mini-empty">上传底稿后显示完整曲线</div>'}</div><div class="lab-card-bottom"><span>${ch.seriesCount}条序列</span><span>${S.wb?(n?`已读取${n}条`:'查看引用'): '查看图表 →'}</span></div></button>`;}).join('')||'<p class="lab-empty">没有匹配的图表</p>';
 $('labChartGrid').querySelectorAll('[data-id]').forEach(b=>b.addEventListener('click',()=>showFocus(b.dataset.id)));
}
function showFocus(id){S.focus=id;let ch=META.charts.find(x=>x.id===id);if(!ch)return;
 const d=S.wb?resolveChart(ch):{title:ch.title,series:[],source:ch.sourceSheet,missing:ch.seriesCount,chart:ch};S.chartData=d;
 $('labFocusTitle').textContent=ch.title;$('labFocusSubtitle').textContent=`原工作表：${ch.sourceSheet} ｜ ${ch.plotType} ｜ 原Excel ${ch.seriesCount}个序列`;
 $('labChartMain').innerHTML=svgFor(d);$('labChartNotice').textContent=ch.broken?'原Excel中该图存在失效或外部工作簿引用。系统仅绘制有效的源序列；缺失部分不会编造数据。':(S.wb?`成功绘制${d.series.length}个序列；若缺失${d.missing}个系列，请检查单元格引用或公式缓存。`:'上传本期原Excel后，当前42个图表将自动读取对应的原始数据。');
 $('labChartTable').innerHTML=d.series.length?renderTable(d):'';
 $('labFocus').scrollIntoView({behavior:'smooth',block:'nearest'});drawCards();
}
function outputMain(kind){const d=S.chartData;if(!d?.series?.length)return alert('请先上传包含图表原始数据的工作簿。');const name=cleanName(d.title);
 if(kind==='csv'){blobDownload(csvFor(d),name+'.csv','text/csv;charset=utf-8');return;}
 let svg=$('labChartMain').querySelector('svg')?.outerHTML;if(!svg)return;
 if(kind==='svg'){blobDownload(svg,name+'.svg','image/svg+xml;charset=utf-8');return;}
 const url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml;charset=utf-8'})),img=new Image();img.onload=()=>{const c=document.createElement('canvas');c.width=1960;c.height=860;const ctx=c.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,c.width,c.height);ctx.drawImage(img,0,0,c.width,c.height);c.toBlob(blob=>{if(blob){const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download=name+'.png';a.click();setTimeout(()=>URL.revokeObjectURL(u),5000)}URL.revokeObjectURL(url)},'image/png')};img.onerror=()=>{URL.revokeObjectURL(url);alert('当前浏览器未能转换PNG，请下载SVG文件。')};img.src=url;
}
function actualCols(name){const sh=S.wb?.Sheets[name];if(!sh)return[];let s=new Set();Object.keys(sh).forEach(k=>{const m=k.match(/^([A-Z]+)\d+$/);if(m)s.add(m[1])});return [...s].sort((a,b)=>colNum(a)-colNum(b));}
function header(name,col){const sh=S.wb?.Sheets[name];let labels=[];for(let r=1;r<=5;r++){let v=cellVal(sh,col,r);if(typeof v==='string'&&v.trim()&&!/^20\d\d年?$/.test(v.trim()))labels.push(v.trim().slice(0,28));}return labels.at(-1)||col;}
function numericCols(name){return actualCols(name).filter(col=>{const sh=S.wb?.Sheets[name];let numCount=0;for(let row=3;row<58;row++)if(num(cellVal(sh,col,row))!==null)numCount++;return numCount>=2});}
function renderSheets(){ $('labSheetGrid').innerHTML=META.sheets.map((s,i)=>`<button class="lab-sheet-card" data-sheet="${i}"><span>${esc(s.category)}</span><b>${esc(s.name)}</b><small>${s.rowCount}行 · ${s.chartCount}图 · ${esc(s.dimension)}</small></button>`).join('');
 $('labSheetGrid').querySelectorAll('[data-sheet]').forEach(el=>el.onclick=()=>{let name=META.sheets[+el.dataset.sheet].name;$('labSheetSelect').value=name;selectSheet(name);$('labSheetChart').scrollIntoView({behavior:'smooth',block:'center'});});
 $('labSheetSelect').innerHTML=META.sheets.map(s=>`<option value="${esc(s.name)}">${esc(s.name)}</option>`).join('');}
function selectSheet(name){S.sheetName=name;let sh=S.wb?.Sheets[name];if(!sh){$('labSheetHint').textContent='请先上传包含此工作表的原始Excel文件。';$('labXcol').innerHTML='';$('labYcols').innerHTML='';$('labSheetChart').innerHTML='<div class="lab-empty">尚未加载工作表</div>'; $('labSheetTable').innerHTML='';return;}
 const cols=actualCols(name),numeric=numericCols(name);
 $('labXcol').innerHTML=cols.map(c=>`<option value="${c}">${c} · ${esc(header(name,c))}</option>`).join('');
 $('labYcols').innerHTML=numeric.map(c=>`<option value="${c}">${c} · ${esc(header(name,c))}</option>`).join('');
 const first=cols[0];$('labXcol').value=first;
 const ys=numeric.filter(c=>c!==first).slice(0,3);[...$('labYcols').options].forEach(o=>o.selected=ys.includes(o.value));
 $('labSheetHint').textContent=`已识别${cols.length}个实际使用列，其中${numeric.length}个数值列。Ctrl/Command + 鼠标点击可多选指标；图表不会修改原Excel。`;
 drawSheet();
}
function sheetRows(name,cols,limit=3000){let sh=S.wb?.Sheets[name];if(!sh)return[];
 let rMax=0;Object.keys(sh).forEach(k=>{let m=k.match(/^[A-Z]+(\d+)$/);if(m)rMax=Math.max(rMax,+m[1])});rMax=Math.min(rMax,limit);
 const out=[];for(let r=1;r<=rMax;r++){let vals=cols.map(c=>cellVal(sh,c,r));if(vals.some(v=>v!==null&&v!==''))out.push({row:r,vals})}return out;
}
function drawSheet(){const name=S.sheetName;if(!S.wb?.Sheets[name])return;
 const xcol=$('labXcol').value,ys=[...$('labYcols').selectedOptions].map(x=>x.value).filter(x=>x!==xcol);
 if(!xcol||!ys.length){$('labSheetChart').innerHTML='<div class="lab-empty">请至少选择1个数值列</div>';return}
 const cols=[xcol,...ys],rr=sheetRows(name,cols,4500);
 const series=ys.map((yc,i)=>({label:header(name,yc),type:'line',points:rr.map(r=>({x:toDate(r.vals[0]??r.row),v:num(r.vals[i+1])})).filter(p=>p.v!==null)})).filter(s=>s.points.length>0);
 const data={title:name+' · 自选指标',series,source:name};S.sheetData=data;
 $('labSheetChart').innerHTML=svgFor(data);
 $('labSheetTable').innerHTML=renderTable(data);
}
function downloadSheet(){const d=S.sheetData;if(!d?.series.length)return alert('请选择列并生成图表。');blobDownload(csvFor(d),cleanName(S.sheetName)+'_选定指标.csv','text/csv;charset=utf-8');}
function init(){if(!$('labChartGrid'))return;
 $('labSheetCount').textContent=META.totals.sheets;$('labChartCount').textContent=META.totals.charts;$('labBrokenCount').textContent=META.totals.chartsWithBrokenRefs;
 const categories=unique(META.charts.map(x=>x.category));$('labCategory').innerHTML='<option value="全部">全部主题</option>'+categories.map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join('');
 $('labCategory').onchange=e=>{S.category=e.target.value;drawCards()};$('labQuery').oninput=e=>{S.query=e.target.value.trim().toLowerCase();drawCards()};
 $('labFile').onchange=async e=>{const f=e.target.files?.[0];if(!f)return;if(!window.XLSX){$('labFileStatus').textContent='Excel解析组件未加载。请检查网络连接或浏览器安全设置。';return}
  $('labFileStatus').textContent='正在浏览器内解析Excel并关联原图…';try{const bytes=await f.arrayBuffer();S.wb=window.XLSX.read(bytes,{type:'array',cellDates:false,cellFormula:true});S.filename=f.name;$('labLoadedCount').textContent=1;const matching=META.sheets.filter(s=>S.wb.Sheets[s.name]).length;
  $('labFileStatus').textContent=`已加载 ${f.name}，匹配到 ${matching}/21 张月报工作表；无需逐图手动上传。`;
  drawCards();selectSheet(matching?META.sheets.find(x=>S.wb.Sheets[x.name]).name:S.wb.SheetNames[0]);$('labSheetSelect').value=S.sheetName;if(S.focus)showFocus(S.focus);
  }catch(err){$('labFileStatus').textContent='文件解析失败：'+err.message;}
 };
 $('labSvg').onclick=()=>outputMain('svg');$('labPng').onclick=()=>outputMain('png');$('labCsv').onclick=()=>outputMain('csv');
 $('labSheetSelect').onchange=e=>selectSheet(e.target.value);$('labDrawSheet').onclick=drawSheet;$('labDownloadSheet').onclick=downloadSheet;
 renderSheets();drawCards();showFocus('chart-01');
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
