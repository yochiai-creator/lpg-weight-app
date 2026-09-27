const vm = require('vm'), fs = require('fs');
const path=require('path'); const GAS=path.join(__dirname,'..','gas');
function colNum(l){let n=0;for(const c of l)n=n*26+c.charCodeAt(0)-64;return n;}
function a1(s){const m=s.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);const r=+m[2],c=colNum(m[1]);const r2=m[4]?+m[4]:r,c2=m[3]?colNum(m[3]):c;return [r,c,r2-r+1,c2-c+1];}
let sid=1;
class Sheet{constructor(ss,name){this.ss=ss;this.name=name;this.cells={};this.fonts={};this.rich={};this.id=sid++;}
 getName(){return this.name} setName(n){this.name=n;return this} getSheetId(){return this.id}
 get(r,c){const v=this.cells[r+','+c];return v===undefined?'':v} set(r,c,v){this.cells[r+','+c]=v}
 getLastColumn(){let m=0;for(const k in this.cells){const c=+k.split(',')[1];if(this.cells[k]!==''&&c>m)m=c}return m}
 getMaxRows(){return Math.max(1000,this.getLastRow())}
 getLastRow(){let m=0;for(const k in this.cells){const r=+k.split(',')[0];if(this.cells[k]!==''&&r>m)m=r}return m}
 getRange(a,b,c,d){if(typeof a==='string')return new Range(this,...a1(a));return new Range(this,a,b,c||1,d||1)}
 appendRow(vals){const r=this.getLastRow()+1;vals.forEach((v,i)=>{if(typeof v==='string'&&v[0]==="'")v=v.slice(1);this.set(r,i+1,v)})}
 copyTo(ss){const s=new Sheet(ss,'コピー '+this.name);s.cells=JSON.parse(JSON.stringify(this.cells));s.fonts=JSON.parse(JSON.stringify(this.fonts));s.rich=JSON.parse(JSON.stringify(this.rich));ss.sheets.push(s);return s}
 setFrozenRows(){}
 deleteRow(r){const n={};for(const k in this.cells){const [rr,cc]=k.split(',').map(Number);if(rr<r)n[k]=this.cells[k];else if(rr>r)n[(rr-1)+','+cc]=this.cells[k];}this.cells=n;return this}}
class Range{constructor(sh,r,c,nr,nc){Object.assign(this,{sh,r,c,nr,nc})}
 getValues(){const o=[];for(let i=0;i<this.nr;i++){const row=[];for(let j=0;j<this.nc;j++)row.push(this.sh.get(this.r+i,this.c+j));o.push(row)}return o}
 getDisplayValues(){return this.getValues().map(r=>r.map(v=>v instanceof Date?v.toISOString().slice(0,10).replace(/-/g,'/'):String(v)))}
 setValues(v){v.forEach((row,i)=>row.forEach((x,j)=>{if(typeof x==='string'&&x[0]==="'")x=x.slice(1);this.sh.set(this.r+i,this.c+j,x)}));return this}
 setValue(x){return this.setValues([[x]])} getValue(){return this.sh.get(this.r,this.c)} clearContent(){for(let i=0;i<this.nr;i++)for(let j=0;j<this.nc;j++)this.sh.set(this.r+i,this.c+j,'');return this}
 setNumberFormat(){return this} insertCheckboxes(){return this} clearDataValidations(){return this}
 setFontWeight(){return this} setFontSize(n){for(let i=0;i<this.nr;i++)for(let j=0;j<this.nc;j++)this.sh.fonts[(this.r+i)+','+(this.c+j)]=n;return this} setRichTextValue(rv){this.sh.rich[this.r+','+this.c]=rv;return this} getRichText(){return this.sh.rich[this.r+','+this.c]} getFontSize(){return this.sh.fonts[this.r+','+this.c]||10} setBackground(){return this} setFontColor(){return this}}
class SS{constructor(){this.sheets=[];this.tz='America/Los_Angeles'} getSpreadsheetTimeZone(){return this.tz} setSpreadsheetTimeZone(z){this.tz=z} getId(){return 'SSID'} getSheetByName(n){return this.sheets.find(s=>s.name===n)||null}
 insertSheet(n){const s=new Sheet(this,n);this.sheets.push(s);return s} setActiveSheet(){} moveActiveSheet(){} getNumSheets(){return this.sheets.length} deleteSheet(sh){this.sheets=this.sheets.filter(x=>x!==sh)} getSheets(){return this.sheets.slice()}}
function load(){
 const ss=new SS(); const props={};
 const tpl=ss.insertSheet('書式_成績表'); tpl.set(1,1,'高圧ガス容器検査成績表\n\n※容器の製造年月を目視確認の上、容器番号前にチェックを入れる'); tpl.set(5,2,'001'); tpl.set(5,4,3); tpl.set(5,6,',');
 const ctx={console,Date,Math,JSON,
  SpreadsheetApp:{getActiveSpreadsheet:()=>ss,openById:()=>ss,flush(){},
   newTextStyle:()=>{const st={};const b={setFontSize(n){st.size=n;return b},setBold(x){st.bold=x;return b},build(){return st}};return b},
   newRichTextValue:()=>{const rv={runs:[]};const b={setText(t){rv.text=t;return b},setTextStyle(a,z,st){rv.runs.push({a,z,size:st.size,bold:st.bold});return b},build(){return rv}};return b}},
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
 const w=ensure(ss,'担当者',['担当者名']); ctx.__oldWorkers=w; w.set(2,1,'山田'); w.set(3,1,'佐藤');
 return {ctx,ss};
}
module.exports={load};
