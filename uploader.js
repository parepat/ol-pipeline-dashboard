var _D=null;

// ── SheetJS Loader ──────────────────────────────────────────────────────────
(function(){
  var cdns=[
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
  ];
  var i=0;
  function next(){
    if(i>=cdns.length){xst('❌ โหลดไม่ได้','color:var(--red-d)');return;}
    var s=document.createElement('script');s.src=cdns[i++];
    s.onload=function(){xst('✅ พร้อมอัปโหลด','color:var(--green-d)');};
    s.onerror=next;document.head.appendChild(s);
  }
  next();
})();

function xst(m,s){var e=document.getElementById('xst');if(e){e.textContent=m;if(s)e.style.cssText=s;}}
function st(t,m){var e=document.getElementById('st');e.className='status '+t;e.textContent=m;}

// ── File Handlers ────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded',function(){
  var fi=document.getElementById('fi');
  fi.addEventListener('change',function(e){if(e.target.files[0])onFile(e.target.files[0]);});
  var uz=document.getElementById('uz');
  uz.addEventListener('dragover',function(e){e.preventDefault();uz.classList.add('drag');});
  uz.addEventListener('dragleave',function(){uz.classList.remove('drag');});
  uz.addEventListener('drop',function(e){
    e.preventDefault();uz.classList.remove('drag');
    if(e.dataTransfer.files[0])onFile(e.dataTransfer.files[0]);
  });
});

function onFile(f){
  if(typeof XLSX==='undefined'){st('info','⏳ SheetJS กำลังโหลด รอสักครู่...');return;}
  st('info','⏳ กำลังอ่าน '+f.name+'...');
  var r=new FileReader();
  r.onload=function(e){parse(e.target.result,f.name);};
  r.readAsArrayBuffer(f);
}

// ── Parse Excel ──────────────────────────────────────────────────────────────
var TEAMS=['Home Loan Acquiring 1','Home Loan Acquiring 2','Home Loan Acquiring 3',
  'Small Business Acquiring 1 ทีม 1','Small Business Acquiring 1 ทีม 2',
  'Small Business Acquiring 1 ทีม 3','Small Business Acquiring 1 ทีม 4',
  'Small Business Acquiring 2 ทีม 1','Small Business Acquiring 2 ทีม 2'];
var SHORT={
  'Home Loan Acquiring 1':'Home Loan Acq. 1','Home Loan Acquiring 2':'Home Loan Acq. 2',
  'Home Loan Acquiring 3':'Home Loan Acq. 3',
  'Small Business Acquiring 1 ทีม 1':'SBA 1 ทีม 1','Small Business Acquiring 1 ทีม 2':'SBA 1 ทีม 2',
  'Small Business Acquiring 1 ทีม 3':'SBA 1 ทีม 3','Small Business Acquiring 1 ทีม 4':'SBA 1 ทีม 4',
  'Small Business Acquiring 2 ทีม 1':'SBA 2 ทีม 1','Small Business Acquiring 2 ทีม 2':'SBA 2 ทีม 2'
};
function sv(v){return String(v||'').trim();}
function nv(v){var x=parseFloat(v);return isNaN(x)?0:x;}
function mTeam(r){r=sv(r);return TEAMS.find(function(t){return r===t||t===r||r.includes(t)||t.includes(r);});}

function parse(buf,fname){
  try{
    var wb=XLSX.read(buf,{type:'array'});
    var pn=wb.SheetNames.find(function(x){return x.includes('Pipeline');});
    if(!pn){st('err','❌ ไม่พบ Sheet Pipeline Overview');return;}
    var rows=XLSX.utils.sheet_to_json(wb.Sheets[pn],{header:1,defval:''});
    var hdr=rows.findIndex(function(r){return r.some(function(c){return sv(c).includes('Pipeline Status');});});
    if(hdr<0)hdr=4;
    var data=rows.slice(hdr+1).filter(function(r){
      var a=sv(r[0]),b=sv(r[1]);return a&&b&&a!=='สรุปรวม'&&!a.startsWith('#');
    });
    if(!data.length){st('err','❌ ไม่พบข้อมูลใน Pipeline Overview');return;}

    var td={};
    TEAMS.forEach(function(t){
      td[t]={n:0,ready:0,wip:0,wait:0,none:0,pkS:0,gwS:0,cfS:0,cnt:0,premSum:0,premCnt:0,members:[],prods:{}};
    });
    var prodC={},objC={};

    data.forEach(function(r){
      var tk=mTeam(r[1]);if(!tk)return;
      var d=td[tk];d.n++;
      var pk=nv(r[12]),gw=nv(r[13]),cf=nv(r[14]);
      if(pk>0){d.pkS+=pk;d.gwS+=gw;d.cfS+=cf;d.cnt++;}
      var pipe=sv(r[15]);
      if(pipe==='พร้อม Submit'||pipe==='พร้อม Submit เร็ว')d.ready++;
      else if(pipe==='กำลังดำเนินการ')d.wip++;
      else if(pipe==='รอ/ติดเงื่อนไข')d.wait++;
      else if(pipe==='ยังไม่มีลูกค้า')d.none++;
      var prem=nv(r[7]);if(prem>0){d.premSum+=prem;d.premCnt++;}
      var prod=sv(r[6]);
      if(prod){var p2=prod.replace('KKPGEN ','').replace('(Par)','').trim();prodC[p2]=(prodC[p2]||0)+1;d.prods[p2]=(d.prods[p2]||0)+1;}
      var obj=sv(r[5]);if(obj)objC[obj]=(objC[obj]||0)+1;
      d.members.push({name:sv(r[0]),pos:sv(r[2]),prod:sv(r[6]),prem:prem,pk:pk,gw:gw,cf:cf,pipe:pipe,reason:sv(r[16])});
    });

    var teams=TEAMS.map(function(t){
      var d=td[t],c=d.cnt||1;
      var pk=d.cnt>0?Math.round(d.pkS/c*100)/100:0;
      var gw=d.cnt>0?Math.round(d.gwS/c*100)/100:0;
      var cf=d.cnt>0?Math.round(d.cfS/c*100)/100:0;
      var avg=(pk+gw+cf)/3;
      var avgPrem=d.premCnt>0?Math.round(d.premSum/d.premCnt):0;
      var st2=avg<3.5||d.ready===0?'เร่งด่วน':d.n>0&&d.ready/d.n>=0.3?'ปกติ':'ติดตาม';
      return{name:SHORT[t]||t,n:d.n,ready:d.ready,wip:d.wip,wait:d.wait,none:d.none,
        pk:pk,gw:gw,cf:cf,avg:avg,st:st2,prods:d.prods,avgPrem:avgPrem,members:d.members};
    });

    var totN=teams.reduce(function(a,t){return a+t.n;},0)||1;
    var totR=teams.reduce(function(a,t){return a+t.ready;},0);
    var totW=teams.reduce(function(a,t){return a+t.wip;},0);
    var totWt=teams.reduce(function(a,t){return a+t.wait;},0);
    var totNn=teams.reduce(function(a,t){return a+t.none;},0);
    var actv=teams.filter(function(t){return t.n>0;});
    var wPK=actv.length?Math.round(actv.reduce(function(a,t){return a+t.pk*t.n;},0)/totN*100)/100:0;
    var wGW=actv.length?Math.round(actv.reduce(function(a,t){return a+t.gw*t.n;},0)/totN*100)/100:0;
    var wCF=actv.length?Math.round(actv.reduce(function(a,t){return a+t.cf*t.n;},0)/totN*100)/100:0;

    var actions=[],an=wb.SheetNames.find(function(x){return x.includes('Action');});
    if(an){
      var ar=XLSX.utils.sheet_to_json(wb.Sheets[an],{header:1,defval:''});
      var ah=ar.findIndex(function(r){var j=r.map(function(c){return sv(c);}).join('|');return j.includes('#')&&j.includes('Action');});
      if(ah<0)ah=4;
      actions=ar.slice(ah+1).filter(function(r){return sv(r[1])&&sv(r[4])&&!sv(r[1]).startsWith('#');}).slice(0,8).map(function(r){
        return{no:r[0],name:sv(r[1]),team:sv(r[2]),pipe:sv(r[3]),action:sv(r[4]),type:sv(r[5]),owner:sv(r[6]),due:sv(r[7]),status:sv(r[8])};
      });
    }

    var risk=data.filter(function(r){var pk=nv(r[12]),gw=nv(r[13]),cf=nv(r[14]);return(pk>0&&pk<=2)||(gw>0&&gw<=2)||(cf>0&&cf<=2);})
      .map(function(r){var tk=mTeam(r[1]);return{name:sv(r[0]),team:SHORT[tk]||sv(r[1]),pk:nv(r[12]),gw:nv(r[13]),cf:nv(r[14]),pipe:sv(r[15]),reason:sv(r[16])};});

    _D={teams:teams,actions:actions,risk:risk,fname:fname,
      totN:totN,totR:totR,totW:totW,totWt:totWt,totNn:totNn,
      wPK:wPK,wGW:wGW,wCF:wCF,rate:Math.round(totR/totN*100),prodC:prodC,objC:objC};

    st('ok','✅ อ่านสำเร็จ — '+data.length+' คน · '+actv.length+' ทีม');
    showPreview(actv);
    document.getElementById('expBtn').style.display='block';
    document.getElementById('hint').textContent='พร้อม Export แล้วค่ะ';
  }catch(e){st('err','❌ '+e.message);console.error(e);}
}

function showPreview(actv){
  var pv=document.getElementById('pv');
  pv.innerHTML=actv.map(function(t){
    var col=t.st==='เร่งด่วน'?'background:var(--red-l);color:var(--red-d)':
             t.st==='ปกติ'?'background:var(--green-l);color:var(--green-d)':'background:var(--amber-l);color:var(--amber-d)';
    return '<div class="pr"><span class="pr-team">'+t.name+'</span>'+
      '<span class="pr-n">'+t.n+' คน</span>'+
      '<span class="pr-badge" style="'+col+'">'+t.st+'</span></div>';
  }).join('');
  pv.classList.add('show');
}

// ── Export Dashboard ─────────────────────────────────────────────────────────
function exportDB(){
  if(!_D){alert('อัปโหลด Excel ก่อนค่ะ');return;}
  var btn=document.getElementById('expBtn');
  btn.disabled=true;btn.textContent='⏳ กำลังโหลด...';
  fetch('ol_pipeline_dashboard.html')
    .then(function(r){if(!r.ok)throw new Error('โหลด ol_pipeline_dashboard.html ไม่ได้ — ต้องวาง 2 ไฟล์ใน folder เดียวกัน');return r.text();})
    .then(function(html){
      try{build(html);}catch(e){
        btn.disabled=false;btn.textContent='💾 Export Dashboard HTML';
        st('err','❌ '+e.message);console.error(e);
      }
    })
    .catch(function(e){
      btn.disabled=false;btn.textContent='💾 Export Dashboard HTML';
      st('err','❌ '+e.message);
    });
}

// ── Build Export HTML ────────────────────────────────────────────────────────
function build(html){
  var d=_D;

  // 1. ลบ upload zone
  var us=html.indexOf('<div class="uf" id="uf">');
  if(us>=0){
    var dep=0,i=us,ue=-1;
    while(i<html.length){
      if(html.slice(i,i+4)==='<div')dep++;
      else if(html.slice(i,i+6)==='</div>'){dep--;if(dep===0){ue=i+6;break;}}
      i++;
    }
    if(ue>0)html=html.slice(0,us)+html.slice(ue);
  }
  html=html.replace(/<button[^>]*id="exportBtn"[^>]*>[^<]*<\/button>/g,'');

  // 2. Escape data JSON
  var tj=JSON.stringify(d.teams).replace(/</g,'\\u003c');
  var aj=JSON.stringify(d.actions).replace(/</g,'\\u003c');
  var rj=JSON.stringify(d.risk).replace(/</g,'\\u003c');
  var mj=JSON.stringify({totN:d.totN,totR:d.totR,totW:d.totW,totWt:d.totWt,totNn:d.totNn,
    wPK:d.wPK,wGW:d.wGW,wCF:d.wCF,rate:d.rate,prodC:d.prodC,objC:d.objC,fname:d.fname}).replace(/</g,'\\u003c');

  // 3. Modal HTML
  var modal=[
    '<div class="modal-overlay" id="detailModal">',
    '<div class="modal">',
    '<div class="modal-hd">',
    '<div><div class="modal-hd-title" id="modalTitle"></div>',
    '<div class="modal-hd-sub" id="modalSub"></div></div>',
    '<button class="modal-close" id="modalCloseBtn">&#215;</button>',
    '</div>',
    '<div class="modal-body">',
    '<div class="modal-stat-row" id="modalStats"></div>',
    '<div style="overflow-x:auto">',
    '<table class="mtbl"><thead><tr>',
    '<th>ชื่อ-สกุล</th><th>ตำแหน่ง</th><th>ผลิตภัณฑ์</th>',
    '<th>เบี้ย (฿)</th><th>PK</th><th>GW</th><th>CF</th><th>Pipeline</th>',
    '</tr></thead><tbody id="modalBody"></tbody>',
    '</table></div></div></div></div>'
  ].join('');

  // 4. Export render script — เขียนเป็น base64 แล้ว decode ใน browser
  //    หลีกเลี่ยงปัญหา HTML tokenizer ทั้งหมด
  var renderCode=[
    '(function(){',
    'var T='+tj+';',
    'var A='+aj+';',
    'var R='+rj+';',
    'var M='+mj+';',
    'function sc(v){return v>=4?"var(--green-d)":v>=3?"var(--amber-d)":"var(--red-d)";}',
    'function ps(p){',
    '  if(!p)return"";',
    '  if(p.indexOf("\u0e1e\u0e23\u0e49\u0e2d\u0e21")>=0)return"background:var(--green-l);color:var(--green-d)";',
    '  if(p.indexOf("\u0e01\u0e33\u0e25\u0e31\u0e07")>=0)return"background:var(--blue-l);color:var(--blue-d)";',
    '  if(p.indexOf("\u0e23\u0e2d")>=0)return"background:var(--amber-l);color:var(--amber-d)";',
    '  return"background:var(--red-l);color:var(--red-d)";',
    '}',
    'window.openModalByIdx=function(idx){',
    '  var t=T[idx];if(!t)return;',
    '  var mo=document.getElementById("detailModal");if(!mo)return;',
    '  document.getElementById("modalTitle").textContent=t.name;',
    '  document.getElementById("modalSub").textContent=t.n+"\u0020\u0e04\u0e19\u0020\u00b7\u0020\u0e40\u0e1a\u0e35\u0e49\u0e22\u0e40\u0e09\u0e25\u0e35\u0e48\u0e22\u0020"+(t.avgPrem>0?t.avgPrem.toLocaleString():"-")+"\u0020\u0e3f/\u0e1b\u0e35";',
    '  var sd=[{n:t.ready,l:"\u0e1e\u0e23\u0e49\u0e2d\u0e21 Submit",bg:"var(--green-l)",fc:"var(--green-d)"},',
    '          {n:t.wip,l:"\u0e01\u0e33\u0e25\u0e31\u0e07\u0e14\u0e33\u0e40\u0e19\u0e34\u0e19\u0e01\u0e32\u0e23",bg:"var(--blue-l)",fc:"var(--blue-d)"},',
    '          {n:t.wait,l:"\u0e23\u0e2d/\u0e15\u0e34\u0e14\u0e40\u0e07\u0e37\u0e48\u0e2d\u0e19\u0e44\u0e02",bg:"var(--amber-l)",fc:"var(--amber-d)"},',
    '          {n:t.none,l:"\u0e22\u0e31\u0e07\u0e44\u0e21\u0e48\u0e21\u0e35\u0e25\u0e39\u0e01\u0e04\u0e49\u0e32",bg:"var(--red-l)",fc:"var(--red-d)"}];',
    '  document.getElementById("modalStats").innerHTML=sd.map(function(x){',
    '    return"<div class=\\"modal-stat\\" style=\\"background:"+x.bg+"\\">'+
        '<div class=\\"modal-stat-n\\" style=\\"color:"+x.fc+"\\">"+x.n+'+
        '"</div><div class=\\"modal-stat-l\\" style=\\"color:"+x.fc+"\\">"+x.l+"</div></div>";',
    '  }).join("");',
    '  var ms=t.members||[];',
    '  document.getElementById("modalBody").innerHTML=ms.length?ms.map(function(m){',
    '    return"<tr><td style=\\"font-weight:500\\">"+m.name+"</td>"+',
    '           "<td>"+m.pos+"</td>"+',
    '           "<td>"+(m.prod||"").replace("KKPGEN ","").replace("(Par)","").trim()+"</td>"+',
    '           "<td style=\\"text-align:right;color:var(--amber-d)\\">"+(m.prem>0?m.prem.toLocaleString():"-")+"</td>"+',
    '           "<td style=\\"color:"+sc(m.pk)+";font-weight:500\\">"+m.pk+"</td>"+',
    '           "<td style=\\"color:"+sc(m.gw)+";font-weight:500\\">"+m.gw+"</td>"+',
    '           "<td style=\\"color:"+sc(m.cf)+";font-weight:500\\">"+m.cf+"</td>"+',
    '           "<td><span style=\\"font-size:9px;padding:2px 6px;border-radius:10px;"+ps(m.pipe)+"\\">"+m.pipe+"</span></td>"+',
    '    "</tr>";',
    '  }).join(""):',
    '  "<tr><td colspan=\\"8\\" style=\\"text-align:center;padding:16px\\">\u0e44\u0e21\u0e48\u0e21\u0e35\u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25\u0e23\u0e32\u0e22\u0e04\u0e19</td></tr>";',
    '  mo.classList.add("open");',
    '  document.body.style.overflow="hidden";',
    '};',
    'document.addEventListener("click",function(e){',
    '  var mo=document.getElementById("detailModal");if(!mo)return;',
    '  if(e.target===mo||e.target.id==="modalCloseBtn")mo.classList.remove("open");',
    '});',
    'function run(){',
    '  if(typeof clearAll!="function"){setTimeout(run,50);return;}',
    '  window._dashTeams=T;',
    '  window._exportMeta=M;',
    '  window._exportActions=A;',
    '  window._exportRisk=R;',
    '  clearAll();',
    '  renderAll(T,M.totN,M.totR,M.totW,M.totWt,M.totNn,M.wPK,M.wGW,M.wCF,M.rate,M.prodC,M.objC,M.fname);',
    '  if(typeof renderActions=="function")renderActions(A);',
    '  if(typeof renderRisk=="function")renderRisk(R);',
    '}',
    'window.addEventListener("load",function(){run();});',
    '})();'
  ]

// build complete

  // Encode data as base64 via data attribute
  var payload = JSON.stringify({
    teams: d.teams, actions: d.actions, risk: d.risk,
    totN: d.totN, totR: d.totR, totW: d.totW, totWt: d.totWt, totNn: d.totNn,
    wPK: d.wPK, wGW: d.wGW, wCF: d.wCF, rate: d.rate, prodC: d.prodC, objC: d.objC, fname: d.fname
  });
  var b64 = btoa(unescape(encodeURIComponent(payload)));

  var modal = '<div class="modal-overlay" id="detailModal">' +
    '<div class="modal"><div class="modal-hd"><div>' +
    '<div class="modal-hd-title" id="modalTitle"></div>' +
    '<div class="modal-hd-sub" id="modalSub"></div>' +
    '</div><button class="modal-close" id="modalCloseBtn">&#215;</button></div>' +
    '<div class="modal-body"><div class="modal-stat-row" id="modalStats"></div>' +
    '<div style="overflow-x:auto"><table class="mtbl"><thead><tr>' +
    '<th>ชื่อ-สกุล</th><th>ตำแหน่ง</th><th>ผลิตภัณฑ์</th>' +
    '<th>เบี้ย (฿)</th><th>PK</th><th>GW</th><th>CF</th><th>Pipeline</th>' +
    '</tr></thead><tbody id="modalBody"></tbody></table></div>' +
    '</div></div></div>';

  var runner = '<div id="_xd" style="display:none" data-v="' + b64 + '"></div>\n' +
    '<scr' + 'ipt>\n' +
    '(function(){\n' +
    'function dec(){try{return JSON.parse(decodeURIComponent(escape(atob(document.getElementById("_xd").getAttribute("data-v")))));}catch(e){return null;}}\n' +
    'function _sc(v){return v>=4?"var(--green-d)":v>=3?"var(--amber-d)":"var(--red-d)";}\n' +
    'function _ps(p){\n' +
    '  if(!p)return"";\n' +
    '  if(p.indexOf("\\u0e1e\\u0e23\\u0e49\\u0e2d\\u0e21")>=0)return"background:var(--green-l);color:var(--green-d)";\n' +
    '  if(p.indexOf("\\u0e01\\u0e33\\u0e25\\u0e31\\u0e07")>=0)return"background:var(--blue-l);color:var(--blue-d)";\n' +
    '  if(p.indexOf("\\u0e23\\u0e2d")>=0)return"background:var(--amber-l);color:var(--amber-d)";\n' +
    '  return"background:var(--red-l);color:var(--red-d)";\n' +
    '}\n' +
    'window.openModalByIdx=function(idx){\n' +
    '  var _j=dec();if(!_j)return;\n' +
    '  var t=_j.teams[idx];if(!t)return;\n' +
    '  var mo=document.getElementById("detailModal");if(!mo)return;\n' +
    '  document.getElementById("modalTitle").textContent=t.name;\n' +
    '  document.getElementById("modalSub").textContent=t.n+" \\u0e04\\u0e19 \\u00b7 \\u0e40\\u0e1a\\u0e35\\u0e49\\u0e22\\u0e40\\u0e09\\u0e25\\u0e35\\u0e48\\u0e22 "+(t.avgPrem>0?t.avgPrem.toLocaleString():"-")+" \\u0e3f/\\u0e1b\\u0e35";\n' +
    '  var sd=[{n:t.ready,l:"\\u0e1e\\u0e23\\u0e49\\u0e2d\\u0e21 Submit",bg:"var(--green-l)",fc:"var(--green-d)"},{n:t.wip,l:"\\u0e01\\u0e33\\u0e25\\u0e31\\u0e07\\u0e14\\u0e33\\u0e40\\u0e19\\u0e34\\u0e19\\u0e01\\u0e32\\u0e23",bg:"var(--blue-l)",fc:"var(--blue-d)"},{n:t.wait,l:"\\u0e23\\u0e2d/\\u0e15\\u0e34\\u0e14\\u0e40\\u0e07\\u0e37\\u0e48\\u0e2d\\u0e19\\u0e44\\u0e02",bg:"var(--amber-l)",fc:"var(--amber-d)"},{n:t.none,l:"\\u0e22\\u0e31\\u0e07\\u0e44\\u0e21\\u0e48\\u0e21\\u0e35\\u0e25\\u0e39\\u0e01\\u0e04\\u0e49\\u0e32",bg:"var(--red-l)",fc:"var(--red-d)"}];\n' +
    '  document.getElementById("modalStats").innerHTML=sd.map(function(x){return"<div class=\\"modal-stat\\" style=\\"background:"+x.bg+"\\"><div class=\\"modal-stat-n\\" style=\\"color:"+x.fc+"\\">"+x.n+"</div><div class=\\"modal-stat-l\\" style=\\"color:"+x.fc+"\\">"+x.l+"</div></div>";}).join("");\n' +
    '  var ms=t.members||[];\n' +
    '  document.getElementById("modalBody").innerHTML=ms.length?ms.map(function(m){return"<tr><td style=\\"font-weight:500\\">"+m.name+"</td><td>"+m.pos+"</td><td>"+(m.prod||"").replace("KKPGEN ","").replace("(Par)","").trim()+"</td><td style=\\"text-align:right;color:var(--amber-d)\\">"+(m.prem>0?m.prem.toLocaleString():"-")+"</td><td style=\\"color:"+_sc(m.pk)+";font-weight:500\\">"+m.pk+"</td><td style=\\"color:"+_sc(m.gw)+";font-weight:500\\">"+m.gw+"</td><td style=\\"color:"+_sc(m.cf)+";font-weight:500\\">"+m.cf+"</td><td><span style=\\"font-size:9px;padding:2px 6px;border-radius:10px;"+_ps(m.pipe)+"\\">"+m.pipe+"</span></td></tr>";}).join(""):"<tr><td colspan=\\"8\\" style=\\"text-align:center;padding:16px\\">\\u0e44\\u0e21\\u0e48\\u0e21\\u0e35\\u0e02\\u0e49\\u0e2d\\u0e21\\u0e39\\u0e25\\u0e23\\u0e32\\u0e22\\u0e04\\u0e19</td></tr>";\n' +
    '  mo.classList.add("open");document.body.style.overflow="hidden";\n' +
    '};\n' +
    'document.addEventListener("click",function(e){var mo=document.getElementById("detailModal");if(!mo)return;if(e.target===mo||e.target.id==="modalCloseBtn")mo.classList.remove("open");});\n' +
    'function run(){\n' +
    '  if(typeof clearAll!="function"){setTimeout(run,50);return;}\n' +
    '  var _j=dec();if(!_j)return;\n' +
    '  window._dashTeams=_j.teams;window._exportMeta=_j;window._exportActions=_j.actions;window._exportRisk=_j.risk;\n' +
    '  clearAll();\n' +
    '  renderAll(_j.teams,_j.totN,_j.totR,_j.totW,_j.totWt,_j.totNn,_j.wPK,_j.wGW,_j.wCF,_j.rate,_j.prodC,_j.objC,_j.fname);\n' +
    '  if(typeof renderActions=="function")renderActions(_j.actions);\n' +
    '  if(typeof renderRisk=="function")renderRisk(_j.risk);\n' +
    '}\n' +
    'window.addEventListener("load",function(){run();});\n' +
    '})();\n' +
    '<' + '/scri' + 'pt>';

  var inject = modal + '\n' + runner + '\n';
  var bodyEnd = html.lastIndexOf('<' + '/body>');
  if (bodyEnd >= 0) html = html.slice(0, bodyEnd) + inject + html.slice(bodyEnd);
  else html += inject;

  var today = new Date();
  var fname = 'OL_Dashboard_' + String(today.getDate()).padStart(2,'0') +
    String(today.getMonth()+1).padStart(2,'0') + today.getFullYear() + '.html';
  var blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = fname;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  st('ok', 'Export สำเร็จ — ' + fname);
}
