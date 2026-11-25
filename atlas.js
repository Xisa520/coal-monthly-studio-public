'use strict';
/* Standalone V2.0 public statistics atlas. No file upload to a server. */
(()=>{
const E=id=>document.getElementById(id);
const A={public:[],local:[],selected:'prod_annual',category:'全部',groups:[],display:'auto',search:'',currentRows:[],svg:'',loading:false};
const PALETTE=['#237f94','#d19b57','#517bad','#79a76b','#bd7b8d','#9467ae','#9e9579','#38a0aa'];
const SVG_NS='http://www.w3.org/2000/svg';
const categories=['全部','产量','进口','区域','需求','投资','绿色','港口库存','动力煤价','电厂与运输'];
const safe=s=>String(s??'').replace(/[&<>"']/g,x=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[x]));
const compact=v=>Number.isFinite(v)?new Intl.NumberFormat('zh-CN',{maximumFractionDigits:2}).format(v):'—';
const onlyNum=v=>typeof v==='number'&&Number.isFinite(v);
const all=()=>[...A.public,...A.local];
const selected=()=>all().find(d=>d.id===A.selected);
const getGroups=d=>[...new Set(d.rows.map(r=>String(r.group)))].sort((a,b)=>a.localeCompare(b,'zh-CN',{numeric:true}));
const orderValues=(d,rows)=>[...new Set(rows.map(r=>r.x))].sort((a,b)=>{
 const ar=rows.find(r=>r.x===a),br=rows.find(r=>r.x===b);
 const oa=ar&&ar.order!=null?ar.order:NaN,ob=br&&br.order!=null?br.order:NaN;
 return Number.isFinite(oa)&&Number.isFinite(ob)?oa-ob:String(a).localeCompare(String(b),'zh-CN',{numeric:true});
 });
const escapeCSV=x=>'"'+String(x??'').replace(/"/g,'""')+'"';
function setLoading(text){E('atlasMainChart').innerHTML=`<div class="atlas-chart-empty">${safe(text)}</div>`;}
function dateX(v){
 if(v instanceof Date&&!isNaN(v))return v.toISOString().slice(0,10);
 if(onlyNum(v)&&v>=30000&&v<70000){const p=window.XLSX?.SSF.parse_date_code(v);if(p)return `${p.y}-${String(p.m).padStart(2,'0')}-${String(p.d).padStart(2,'0')}`;}
 const s=String(v??'').trim();
 let m=s.match(/^(20\d{2})[-\/.年](\d{1,2})(?:[-\/.月](\d{1,2}))?/);
 if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${(m[3]||'01').padStart(2,'0')}`;
 return '';
}
function dimsData(id,name,category,unit,kind,source_sheet,rows,about){return{id,name,category,unit,kind,source_sheet,source:'仅用户浏览器本次导入；没有上传到GitHub或第三方服务器',about,rows,local:true};}
function toRows(ds,opts={}){const gs=(opts.groups||A.groups);return ds.rows.filter(r=>gs.includes(String(r.group)))}
function miniChart(d){
 let g=getGroups(d).at(-1);if(!g)return '<div style="height:170px"></div>';
 let rs=d.rows.filter(r=>r.group===g),xs=orderValues(d,rs),index=new Map(rs.map(r=>[r.x,r.value]));
 if(d.kind==='province') {xs=xs.sort((a,b)=>(index.get(b)||0)-(index.get(a)||0)).slice(0,12);return svgChart(d,rs,[g],{width:480,height:180,bar:true,minimal:true,focusXs:xs})}
 return svgChart(d,rs,[g],{width:480,height:180,minimal:true});
}
function svgChart(d,raw,gs,opts={}){
 let W=opts.width||1000,H=opts.height||410;
 let pad=opts.minimal?{l:46,r:12,t:17,b:32}:{l:76,r:34,t:35,b:67};
 let isBar=(opts.bar||((A.display==='auto'&&(d.kind==='province'||d.kind==='category'))||A.display==='bar'));
 const plotW=W-pad.l-pad.r,plotH=H-pad.t-pad.b;
 const byG=new Map(gs.map(g=>[g,new Map(raw.filter(r=>r.group===g).map(r=>[String(r.x),r.value]))]));
 let xs=opts.focusXs||orderValues(d,raw);
 if(!xs.length) return `<svg xmlns="${SVG_NS}" viewBox="0 0 ${W} ${H}" role="img"><text x="50%" y="50%" text-anchor="middle" class="atlas-axis">没有匹配的数据</text></svg>`;
 // Provincial and coal type categories are most readable as sorted horizontal bars
 const horizontal = isBar&&(d.kind==='province'||d.kind==='category');
 if(horizontal){
    xs=[...xs].sort((a,b)=>Math.max(...gs.map(g=>byG.get(g)?.get(b)||0))-Math.max(...gs.map(g=>byG.get(g)?.get(a)||0))).slice(0,opts.minimal?10:18);
    const mx=Math.max(1,...raw.map(r=>r.value));const left=opts.minimal?72:115;const top=opts.minimal?9:25;const right=opts.minimal?28:70;const bottom=opts.minimal?17:42;
    const rowH=(H-top-bottom)/xs.length, bw=Math.max(3,Math.min(opts.minimal?12:17,rowH/(gs.length+1.35))),scale=(W-left-right)/(mx*1.09);
    let svg=`<svg xmlns="${SVG_NS}" viewBox="0 0 ${W} ${H}" role="img" aria-label="${safe(d.name)}">`;
    for(let t=0;t<=4;t++){const xx=left+(W-left-right)*t/4;svg+=`<line x1="${xx}" y1="${top-4}" x2="${xx}" y2="${H-bottom+5}" stroke="#ecf1f3"/><text class="atlas-axis" x="${xx}" y="${H-bottom+19}" text-anchor="middle">${safe(compact(mx*1.09*t/4))}</text>`;}
    xs.forEach((x,i)=>{let yy=top+(i+.5)*rowH;svg+=`<text class="atlas-axis" x="${left-11}" y="${yy+3}" text-anchor="end">${safe(x)}</text>`;
      gs.forEach((g,j)=>{const v=byG.get(g)?.get(x);if(!onlyNum(v))return;const y=yy+(j-(gs.length-1)/2)*(bw+2)-bw/2;svg+=`<rect x="${left}" y="${y.toFixed(1)}" width="${Math.max(0,v*scale).toFixed(1)}" height="${bw.toFixed(1)}" rx="2" fill="${PALETTE[j%PALETTE.length]}" opacity=".9"><title>${safe(g)} · ${safe(x)}：${compact(v)} ${safe(d.unit)}</title></rect>`;});
    });
    if(!opts.minimal)svg+=legendSvg(gs,W-230,17);
    return svg+'</svg>';
 }
 // Prevent 2500 daily data points from making an enormous DOM on weak machines.
 const skip=xs.length>450?Math.ceil(xs.length/450):1;
 if(skip>1)xs=xs.filter((_,i)=>i%skip===0||i===xs.length-1);
 let vals=raw.map(r=>r.value).filter(onlyNum);let ymin=Math.min(0,...vals),ymax=Math.max(...vals,1);
 if(!isBar){let mn=Math.min(...vals),mx=Math.max(...vals);const span=Math.max(1,(mx-mn));ymin=Math.max(0,mn-span*.18);ymax=mx+span*.18;}
 if(ymax<=ymin)ymax=ymin+1;
 const xx=i=>pad.l+(xs.length===1?plotW/2:i*plotW/(xs.length-1));
 const yy=v=>pad.t+(ymax-v)/(ymax-ymin)*plotH;
 let svg=`<svg xmlns="${SVG_NS}" viewBox="0 0 ${W} ${H}" role="img" aria-label="${safe(d.name)}">`;
 for(let i=0;i<=4;i++){const v=ymax-(ymax-ymin)*i/4,y=pad.t+plotH*i/4;svg+=`<line x1="${pad.l}" y1="${y.toFixed(2)}" x2="${W-pad.r}" y2="${y.toFixed(2)}" stroke="#e9eff3"/><text class="atlas-axis" x="${pad.l-10}" y="${(y+4).toFixed(1)}" text-anchor="end">${safe(compact(v))}</text>`;}
 const labelEvery=Math.max(1,Math.ceil(xs.length/(opts.minimal?5:8)));
 xs.forEach((x,i)=>{if(i%labelEvery!==0&&i!==xs.length-1)return;const t=String(x);svg+=`<text class="atlas-axis" x="${xx(i).toFixed(1)}" y="${H-pad.b+22}" text-anchor="middle">${safe(t.slice(0,11))}</text>`;});
 gs.forEach((g,k)=>{const mp=byG.get(g);let points=[]; xs.forEach((x,i)=>{const v=mp?.get(x);if(onlyNum(v))points.push([xx(i),yy(v),v,x])});if(!points.length)return;
 const color=PALETTE[k%PALETTE.length];
 if(isBar){let barW=Math.max(2,Math.min(22,(plotW/xs.length)*.70/gs.length));points.forEach(([x,y,v,label])=>{const gx=x+(k-(gs.length-1)/2)*barW;svg+=`<rect x="${(gx-barW*.45).toFixed(2)}" y="${y.toFixed(2)}" width="${(barW*.88).toFixed(2)}" height="${Math.max(0,pad.t+plotH-y).toFixed(2)}" rx="2" fill="${color}"><title>${safe(g)} · ${safe(label)}：${compact(v)} ${safe(d.unit)}</title></rect>`;});}
 else {let path='';let last=-2; xs.forEach((x,i)=>{const v=mp?.get(x);if(onlyNum(v)){path+=(last===i-1?'L':'M')+`${xx(i).toFixed(1)} ${yy(v).toFixed(1)} `;last=i}else last=-2;});svg+=`<path d="${path}" fill="none" stroke="${color}" stroke-width="${opts.minimal?2.2:2.6}" stroke-linecap="round" stroke-linejoin="round" opacity=".97"/>`;
   if(xs.length<=85||opts.minimal)points.forEach(([x,y,v,label])=>{svg+=`<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="${opts.minimal?2.4:3.0}" fill="${color}"><title>${safe(g)} · ${safe(label)}：${compact(v)} ${safe(d.unit)}</title></circle>`;});
 }
 });
 if(!opts.minimal)svg+=legendSvg(gs,W-248,16);
 return svg+'</svg>';
}
function legendSvg(groups,x,y){let out='<g>';
 groups.slice(0,6).forEach((g,i)=>{const xx=x+(i%2)*120, yy=y+Math.floor(i/2)*17;out+=`<rect x="${xx}" y="${yy}" width="9" height="9" rx="2" fill="${PALETTE[i%PALETTE.length]}"/><text class="atlas-axis" x="${xx+14}" y="${yy+9}">${safe(String(g).slice(0,11))}</text>`});return out+'</g>'}
function summarize(d,raw){const latest=d.rows.slice(-1)[0];const years=getGroups(d),lastG=years.at(-1);let rr=d.rows.filter(r=>r.group===lastG);const vv=rr.map(r=>r.value).filter(onlyNum);const max=vv.length?Math.max(...vv):NaN;const min=vv.length?Math.min(...vv):NaN;const sum=vv.reduce((a,b)=>a+b,0);return [
 {label:'展示记录',value:compact(raw.length)+' 条'},
 {label:'当前筛选峰值',value:compact(Math.max(...raw.map(r=>r.value)))+' '+d.unit},
 {label:'当前筛选最小值',value:compact(Math.min(...raw.map(r=>r.value)))+' '+d.unit},
 {label:'最后已收录年份',value:d.rows.filter(r=>/^20\d{2}$/.test(String(r.group))).map(r=>String(r.group)).sort().at(-1)||String(latest?.x||'—')}
 ]}
function renderKPIs(){const get=(id,x,group)=>A.public.find(d=>d.id===id)?.rows.find(r=>r.x===x&&(!group||r.group===group));const kpis=[
 ['2025年全国原煤产量',get('prod_annual','2025')?.value,'亿吨','年度数据 / 原表统计局口径'],
 ['2025年煤炭进口量',get('import_annual','2025')?.value,'万吨','年度海关进口口径'],
 ['2025年煤炭采选投资',get('mining_invest','2025')?.value,'亿元','原表投资额 / 待核查'],
 ['2019—2025年分省产量',new Set(A.public.find(d=>d.id==='prov_prod')?.rows.map(x=>x.x)).size,'个省区','可按年度比较与排序']];
 E('atlasKpis').innerHTML=kpis.map(([t,v,u,s])=>`<div class="atlas-kpi"><small>${safe(t)}</small><strong>${compact(v)} <em>${safe(u)}</em></strong><span class="kpi-sub">${safe(s)}</span></div>`).join('');
 E('atlasPublicNum').textContent=A.public.length;E('atlasDatasetCount').textContent=`${A.public.length}个统计专题 · ${A.public.reduce((n,d)=>n+d.rows.length,0)}条历史记录`;
}
function renderMini(){const ids=['prod_annual','prod_month','import_annual','prov_prod','mining_invest','import_month'];const items=ids.map(id=>A.public.find(d=>d.id===id)).filter(Boolean);
 E('atlasMiniGrid').innerHTML=items.map((d,i)=>`<article class="atlas-mini-card" tabindex="0" role="button" data-mini="${safe(d.id)}" aria-label="查看${safe(d.name)}详情"><div class="atlas-mini-header"><div><h3>${safe(d.name)}</h3><p>${safe(d.unit)} · ${d.kind==='province'?'最新年度省区排序':d.kind==='month'?'分年度月度对比':'年度历史走势'}</p></div><span class="atlas-mini-badge">专题 ${String(i+1).padStart(2,'0')}</span></div><div class="atlas-mini-svg">${miniChart(d)}</div><div class="atlas-mini-footer"><span>${safe(d.source_sheet)}</span><span>查看交互图表 →</span></div></article>`).join('');
 E('atlasMiniGrid').querySelectorAll('[data-mini]').forEach(el=>{const trigger=()=>{E('atlasCategory').value='全部';A.category='全部';rebuildSelect();E('atlasDataset').value=el.dataset.mini;choose(el.dataset.mini);E('atlasExplorer').scrollIntoView({behavior:'smooth',block:'start'})};el.addEventListener('click',trigger);el.addEventListener('keydown',e=>{if(e.key==='Enter')trigger()})});
}
function rebuildSelect(){const valid=all().filter(d=>A.category==='全部'||d.category===A.category);E('atlasDataset').innerHTML=valid.map(d=>`<option value="${safe(d.id)}">${d.local?'⌁ ':''}${safe(d.name)}</option>`).join('');
 if(!valid.some(d=>d.id===A.selected))A.selected=valid[0]?.id||'';E('atlasDataset').value=A.selected;}
function choose(id){A.selected=id;const d=selected();if(!d)return setLoading('没有可展示的数据');const groups=getGroups(d);A.groups=(d.kind==='province'||d.kind==='category')?[groups.at(-1)]:groups.slice(-Math.min(d.kind==='year'?1:3,groups.length));if(d.kind==='timeseries')A.groups=groups.slice(0,4);E('atlasSearch').value='';A.search='';renderGroups();renderDetail();}
function renderGroups(){const d=selected();if(!d)return;E('atlasGroupOptions').innerHTML=getGroups(d).map(g=>`<button type="button" class="atlas-chip ${A.groups.includes(g)?'active':''}" data-g="${safe(g)}" aria-pressed="${A.groups.includes(g)}">${safe(g)}</button>`).join('');E('atlasGroupOptions').querySelectorAll('[data-g]').forEach(b=>b.onclick=()=>{const g=b.dataset.g;A.groups=A.groups.includes(g)?A.groups.filter(x=>x!==g):[...A.groups,g];if(!A.groups.length)A.groups=[g];renderGroups();renderDetail()});}
function renderDetail(){const d=selected();if(!d)return;const rs=toRows(d);A.currentRows=rs;E('atlasDetailTitle').textContent=d.name;E('atlasDetailDescription').textContent=d.about;E('atlasMainChart').innerHTML=svgChart(d,rs,A.groups,{});A.svg=E('atlasMainChart').querySelector('svg')?.outerHTML||'';E('atlasDetailStats').innerHTML=summarize(d,rs).map(v=>`<div class="atlas-stat"><small>${safe(v.label)}</small><strong>${safe(v.value)}</strong></div>`).join('');E('atlasSource').innerHTML=`<strong>数据源：</strong>${safe(d.source)}　｜　<strong>工作表：</strong>${safe(d.source_sheet)}　｜　<strong>单位：</strong>${safe(d.unit)}　｜　${d.local?'仅本地导入、未上传':'未进行逐项官方核验'}。`;
 renderTable();}
function renderTable(){const d=selected();if(!d)return;let term=A.search.trim().toLowerCase();const rs=A.currentRows.filter(r=>!term||`${r.x} ${r.group} ${r.value}`.toLowerCase().includes(term));E('atlasRowCount').textContent=`共${rs.length}条`;
 E('atlasTable').innerHTML=rs.length?`<table><thead><tr><th>时期 / 区域</th><th>年份 / 序列</th><th style="text-align:right">数值 (${safe(d.unit)})</th></tr></thead><tbody>${rs.slice(0,300).map(r=>`<tr><td>${safe(r.x)}</td><td>${safe(r.group)}</td><td class="value">${compact(r.value)}</td></tr>`).join('')}</tbody></table>${rs.length>300?'<div style="padding:12px;color:#8194a1;font-size:11px">预览显示前300行；CSV可导出全部筛选记录。</div>':''}`:'<p style="padding:25px;color:#8194a1">没有匹配的数据</p>';
}
function fileDown(blob,name){const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),6000);}
function saveCSV(){const d=selected();if(!d)return;let txt='\ufeff时期,年份或序列,值,单位,工作表\r\n'+A.currentRows.map(r=>[r.x,r.group,r.value,d.unit,d.source_sheet].map(escapeCSV).join(',')).join('\r\n');fileDown(new Blob([txt],{type:'text/csv;charset=utf-8'}),`${d.id}-数据.csv`)}
function savePNG(){const d=selected();if(!d||!A.svg)return;const svg=A.svg,img=new Image(),canvas=document.createElement('canvas');canvas.width=1600;canvas.height=656;const ctx=canvas.getContext('2d'),blob=new Blob([svg],{type:'image/svg+xml;charset=utf-8'}),u=URL.createObjectURL(blob);img.onload=()=>{ctx.fillStyle='#fff';ctx.fillRect(0,0,1600,656);ctx.drawImage(img,0,0,1600,656);URL.revokeObjectURL(u);canvas.toBlob(b=>b&&fileDown(b,`${d.id}-趋势图.png`),'image/png')};img.onerror=()=>{URL.revokeObjectURL(u);alert('图片导出失败，建议使用浏览器截图。')};img.src=u;}
function parseWorkbook(wb,fileName){const out=[];let counter=0;
 function fromDateValueSheet(sh,label,category){const rows=window.XLSX.utils.sheet_to_json(wb.Sheets[sh],{header:1,raw:true,defval:null,blankrows:false});const info=String(rows[0]?.[1]||label);const unitMatch=info.match(/（([^）]+)）/);const unit=unitMatch?unitMatch[1]:(/价格|平仓价|坑口价/.test(info)?'元/吨':'万吨');const data=[];
  for(const row of rows.slice(1)){const dt=dateX(row[1]);const vv=row[2];if(dt&&onlyNum(vv))data.push({x:dt,group:info.replace(/（[^）]*）/g,''),value:+vv})};
  if(data.length>10)out.push(dimsData('local_'+(++counter)+'_'+Date.now(),info,category,unit,'timeseries',`${fileName} / ${sh}`,data,'由原始工作簿B列日期、C列数值自动生成；不会执行任何宏。'));
 }
 const nameCategories={
 '环渤海库存':'港口库存','秦港库存':'港口库存','秦港平仓价':'动力煤价','鄂尔多斯坑口价':'动力煤价',
 '大秦线发运量':'电厂与运输','环渤海调入量':'电厂与运输','环渤海调出量':'电厂与运输',
 '秦港调入量':'电厂与运输','秦港吞吐量':'电厂与运输','内河港口库存':'港口库存',
 '南方港口库存':'港口库存','华东港口库存':'港口库存','华南港口库存':'港口库存'
 };
 for(const s of wb.SheetNames){if(nameCategories[s])fromDateValueSheet(s,s,nameCategories[s]);}
 for(const sheet of ['25省日耗','25省库存','25省可用天数']){
   if(!wb.Sheets[sheet])continue;let raw=window.XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1,raw:true,defval:null,blankrows:false});
   for(const [col,area] of [[26,'25省合计'],[27,'内陆17省'],[28,'沿海八省']]){let rows=[];
     for(const r of raw.slice(1)){let dt=dateX(r[0]),vv=r[col];if(!dt||!onlyNum(vv))continue;const y=dt.slice(0,4),day=dt.slice(5);rows.push({x:day,group:y,value:+vv,order:Number(day.replace('-',''))})}
     if(rows.length>10)out.push(dimsData('local_'+(++counter)+'_'+Date.now(),`${area} · ${sheet.replace('25省','')}`,'电厂与运输',sheet.includes('可用天数')?'天':'万吨','month',`${fileName} / ${sheet} / ${{26:"AA",27:"AB",28:"AC"}[col]}列`,rows,`按原始工作表第${col+1}列（${area}）识别，已将不同年份按月日叠加比较；只在浏览器读取。`));
   }
 }
 if(wb.Sheets['Sheet1']){
   const rows=window.XLSX.utils.sheet_to_json(wb.Sheets['Sheet1'],{header:1,raw:true,defval:null,blankrows:false});
   if(String(rows[0]?.[1]||'').includes('秦皇岛')&&String(rows[0]?.[2]||'').includes('长协')){
     const rr=[];for(const row of rows.slice(1)){const dt=dateX(row[0]);if(!dt)continue;for(const [col,group] of [[1,'市场平仓价'],[2,'年度长协价']])if(onlyNum(row[col]))rr.push({x:dt,group,value:+row[col]});}
     if(rr.length>20)out.push(dimsData('local_'+(++counter)+'_'+Date.now(),'秦皇岛5500大卡市场价与长协价','动力煤价','元/吨','timeseries',`${fileName} / Sheet1`,rr,'读取原表A列日期、B列市场价、C列长协价；可并排比较两条曲线。'));
   }
 }
 // Annual yearbook has price + stock source tables that can be charted locally, without publicizing data.
 for(const sheet of wb.SheetNames.filter(n=>n.startsWith('33-秦皇岛动力煤价格'))){
   const rows=window.XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1,raw:true,defval:null,blankrows:false}),rr=[];
   for(const row of rows.slice(2)){const dt=dateX(row[0]);if(!dt)continue;for(const[col,group]of [[2,'秦港5500动力煤市场价'],[4,'环渤海动力煤指数']])if(onlyNum(row[col]))rr.push({x:dt,group,value:+row[col]});}
   if(rr.length>20)out.push(dimsData('local_'+(++counter)+'_'+Date.now(),'秦皇岛动力煤市场价与环渤海指数','动力煤价','元/吨','timeseries',`${fileName} / ${sheet}`,rr,'读取原始年报图表价格底稿（A/C/E列），只在本地浏览器使用。'));
 }
 return out;
}
async function importFiles(files){if(!files?.length)return;if(!window.XLSX){E('atlasImportStatus').textContent='Excel组件未加载，请检查浏览器网络或稍后重试。';return}E('atlasImportStatus').textContent='正在浏览器本地识别工作表及历史序列…';let count=0,notes=[];
 for(const file of files){if(file.size>27*1024*1024){notes.push(file.name+'大于27MB，请拆分后再试');continue}
 try{const bytes=await file.arrayBuffer(),wb=window.XLSX.read(bytes,{type:'array',raw:true});let sets=parseWorkbook(wb,file.name);if(!sets.length){notes.push(file.name+'暂无预设识别规则，可改用左侧“Excel数据导入”手动映射');continue}A.local.push(...sets);count+=sets.length;notes.push(`${file.name}：识别${sets.length}组指标`);}
 catch(e){notes.push(`${file.name}：读取失败 ${e.message}`)}
 }
 E('atlasLocalNum').textContent=A.local.length;E('atlasImportStatus').textContent=notes.join(' ｜ ')||'尚未识别到有效序列';
 if(count){const found=A.local[A.local.length-count];A.category=found.category;E('atlasCategory').value=A.category;rebuildSelect();choose(found.id);E('atlasDataset').value=found.id;E('atlasExplorer').scrollIntoView({behavior:'smooth',block:'start'});} }
async function initAtlas(){const select=E('atlasDataset');if(!select)return;setLoading('正在读取预置的煤炭市场统计数据…');
 E('atlasCategory').innerHTML=categories.map(c=>`<option>${c}</option>`).join('');
 E('atlasCategory').onchange=e=>{A.category=e.target.value;rebuildSelect();choose(A.selected)};
 E('atlasDataset').onchange=e=>choose(e.target.value);
 E('atlasChartType').onchange=e=>{A.display=e.target.value;renderDetail()};
 E('atlasSearch').oninput=e=>{A.search=e.target.value;renderTable()};
 E('atlasSaveCsv').onclick=saveCSV;E('atlasSavePng').onclick=savePNG;
 E('atlasFile').onchange=e=>importFiles(e.target.files);
 E('atlasScrollLocal').onclick=()=>E('atlasLocalSection').scrollIntoView({behavior:'smooth',block:'start'});
 const drop=E('atlasLocalSection');drop.addEventListener('dragover',e=>e.preventDefault());drop.addEventListener('drop',e=>{e.preventDefault();importFiles(e.dataTransfer.files)});
 try{let data=window.COAL_PUBLIC_DATA;if(!data){const response=await fetch('./data/coal-public.json',{cache:'no-store'});if(!response.ok)throw Error('HTTP '+response.status);data=await response.json()}A.public=data.datasets||[];
 if(!A.public.length)throw Error('数据文件没有指标'); renderKPIs();renderMini();rebuildSelect();choose('prod_annual');}
 catch(e){setLoading('无法加载预置数据，请确认已上传data/coal-public.json：'+e.message);E('atlasKpis').textContent='数据加载失败，请检查文件路径。';E('atlasMiniGrid').textContent='数据加载失败。';}
}
document.addEventListener('DOMContentLoaded',initAtlas);
})();
