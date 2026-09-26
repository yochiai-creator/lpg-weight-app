const vm = require('vm'), fs = require('fs');
const path=require('path'); const GAS=path.join(__dirname,'..','gas');
function colNum(l){let n=0;for(const c of l)n=n*26+c.charCodeAt(0)-64;return n;}
function a1(s){const m=s.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);const r=+m[2],c=colNum(m[1]);const r2=m[4]?+m[4]:r,c2=m[3]?colNum(m[3]):c;return [r,c,r2-r+1,c2-c+1];}
let sid=1;
class Sheet{constructor(ss,name){this.ss=ss;this.name=name;this.cells={};this.id=sid++;}
 getName(){return this.name} setName(n){this.name=n;return this} getSheetId(){return this.id}
 get(r,c){const v=this.cells[r+','+c];return v===undefined?'':v} set(r,c,v){this.cells[r+','+c]=v}
 getLastColumn(){let m=0;for(const k in this.cells){const c=+k.split(',')[1];if(this.cells[k]!==''&&c>m)m=c}return m}
 getLastRow(){let m=0;for(const k in this.cells){const r=+k.split(',')[0];if(this.cells[k]!==''&&r>m)m=r}return m}
 getRange(a,b,c,d){if(typeof a==='string')return new Range(this,...a1(a));return new Range(this,a,b,c||1,d||1)}
 appendRow(vals){const r=this.getLastRow()+1;vals.forEach((v,i)=>{if(typeof v==='string'&&v[0]==="'")v=v.slice(1);this.set(r,i+1,v)})}
 copyTo(ss){const s=new Sheet(ss,'コピー '+this.name);s.cells=JSON.parse(JSON.stringify(this.cells));ss.sheets.push(s);return s}
 setFrozenRows(){}}
class Range{constructor(sh,r,c,nr,nc){Object.assign(this,{sh,r,c,nr,nc})}
 getValues(){const o=[];for(let i=0;i<this.nr;i++){const row=[];for(let j=0;j<this.nc;j++)row.push(this.sh.get(this.r+i,this.c+j));o.push(row)}return o}
 getDisplayValues(){return this.getValues().map(r=>r.map(v=>v instanceof Date?v.toISOString().slice(0,10).replace(/-/g,'/'):String(v)))}
 setValues(v){v.forEach((row,i)=>row.forEach((x,j)=>{if(typeof x==='string'&&x[0]==="'")x=x.slice(1);this.sh.set(this.r+i,this.c+j,x)}));return this}
 setValue(x){return this.setValues([[x]])} clearContent(){for(let i=0;i<this.nr;i++)for(let j=0;j<this.nc;j++)this.sh.set(this.r+i,this.c+j,'');return this}
 setNumberFormat(){return this} insertCheckboxes(){return this} clearDataValidations(){return this}
 setFontWeight(){return this} setBackground(){return this} setFontColor(){return this}}
class SS{constructor(){this.sheets=[]} getId(){return 'SSID'} getSheetByName(n){return this.sheets.find(s=>s.name===n)||null}
 insertSheet(n){const s=new Sheet(this,n);this.sheets.push(s);return s} setActiveSheet(){} moveActiveSheet(){} getNumSheets(){return this.sheets.length}}
function load(){
 const ss=new SS(); const props={};
 const tpl=ss.insertSheet('書式_成績表'); tpl.set(5,2,'001'); tpl.set(5,4,3); tpl.set(5,6,',');
 const ctx={console,Date,Math,JSON,
  SpreadsheetApp:{getActiveSpreadsheet:()=>ss,openById:()=>ss,flush(){}},
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]||null,setProperty:(k,v)=>props[k]=v})},
  LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
  Session:{getActiveUser:()=>({getEmail:()=>'tester@example.com'})},
  Utilities:{formatDate:(d)=>d.toISOString().replace(/\D/g,'').slice(2,14),newBlob:(c,t,n)=>({name:n,type:t,content:c,getName(){return n}})},
  MailApp:{sendEmail:(to,sub,body,opt)=>{global.MAILS.push({to,sub,body,cc:opt.cc,name:opt.name,att:opt.attachments.map(a=>a.name)})}},
  UrlFetchApp:{fetch:()=>({getBlob:()=>({setName(n){this.name=n;return this}})})},
  ScriptApp:{getOAuthToken:()=>'t'},
  DriveApp:{createFolder:()=>({getId:()=>'F',createFile:(b)=>({getUrl:()=>'https://drive/'+b.name})}),getFolderById:()=>({createFile:(b)=>({getUrl:()=>'https://drive/'+b.name})})}};
 global.MAILS=[];
 vm.createContext(ctx); vm.runInContext(fs.readFileSync(GAS+'/Code.js','utf8'),ctx);
 vm.runInContext(fs.readFileSync(GAS+'/Setup.js','utf8'),ctx);
 vm.runInContext(fs.readFileSync(GAS+'/Delivery.js','utf8'),ctx);
 ensure=vm.runInContext('ensureSheet_',ctx); ensure(ss,'ロット',vm.runInContext('LOT_HEADERS',ctx)); ensure(ss,'入力記録',vm.runInContext('LOG_HEADERS',ctx));
 const st=ensure(ss,'設定',['項目','値','説明']); vm.runInContext('SETTING_ROWS',ctx).forEach((r,i)=>{st.set(i+2,1,r[0]);st.set(i+2,2,r[1]);});
 st.set(2,2,'nouhin@example.com');
 const w=ensure(ss,'担当者',['担当者名']); w.set(2,1,'山田'); w.set(3,1,'佐藤');
 return {ctx,ss};
}
module.exports={load};
