const vm = require('vm'), fs = require('fs');
const path=require('path'); const GAS=path.join(__dirname,'..','gas');
function colNum(l){let n=0;for(const c of l)n=n*26+c.charCodeAt(0)-64;return n;}
function a1(s){const m=s.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);const r=+m[2],c=colNum(m[1]);const r2=m[4]?+m[4]:r,c2=m[3]?colNum(m[3]):c;return [r,c,r2-r+1,c2-c+1];}
let sid=1;
class Sheet{constructor(ss,name){this.ss=ss;this.name=name;this.cells={};this.fonts={};this.rich={};this.id=sid++;}
 getName(){return this.name} setName(n){this.name=n;return this} getSheetId(){return this.id}
 get(r,c){const v=this.cells[r+','+c];return v===undefined?'':v} set(r,c,v){this.cells[r+','+c]=v}
 getLastColumn(){let m=0;for(const k in this.cells){const c=+k.split(',')[1];if(this.cells[k]!==''&&c>m)m=c}return m}
 getMaxRows(){return Math.max(this.maxR||1000,this.getLastRow())} getMaxColumns(){return Math.max(this.maxC||26,this.getLastColumn())}
 deleteRows(st,n){this.maxR=this.getMaxRows()-n;const o={};for(const k in this.cells){const [r,c]=k.split(',').map(Number);if(r<st)o[k]=this.cells[k];else if(r>=st+n)o[(r-n)+','+c]=this.cells[k]}this.cells=o;return this} deleteColumns(st,n){this.maxC=this.getMaxColumns()-n;for(const k in this.cells){const c=+k.split(',')[1];if(c>=st&&c<st+n)delete this.cells[k]}return this}
 getLastRow(){let m=0;for(const k in this.cells){const r=+k.split(',')[0];if(this.cells[k]!==''&&r>m)m=r}return m}
 getRange(a,b,c,d){if(typeof a==='string')return new Range(this,...a1(a));return new Range(this,a,b,c||1,d||1)}
 appendRow(vals){const r=this.getLastRow()+1;vals.forEach((v,i)=>{if(typeof v==='string'&&v[0]==="'")v=v.slice(1);this.set(r,i+1,v)})}
 copyTo(ss){const s=new Sheet(ss,'コピー '+this.name);s.cells=JSON.parse(JSON.stringify(this.cells));s.fonts=JSON.parse(JSON.stringify(this.fonts));s.rich=JSON.parse(JSON.stringify(this.rich));s.maxR=this.maxR;s.maxC=this.maxC;ss.sheets.push(s);return s}
 setFrozenRows(){} hideSheet(){return this} getParent(){return this.ss}
 moveColumns(range,dest){const src=[];for(let j=0;j<range.nc;j++)src.push(range.c+j);let maxC=Math.max(this.getLastColumn(),dest,...src);const rest=[];for(let c=1;c<=maxC;c++)if(src.indexOf(c)<0)rest.push(c);const at=rest.filter(c=>c<dest).length;const order=rest.slice(0,at).concat(src,rest.slice(at));const n={};for(const k in this.cells){const [rr,cc]=k.split(',').map(Number);const ni=order.indexOf(cc);n[rr+','+(ni<0?cc:ni+1)]=this.cells[k];}this.cells=n;return this}
 deleteRow(r){const n={};for(const k in this.cells){const [rr,cc]=k.split(',').map(Number);if(rr<r)n[k]=this.cells[k];else if(rr>r)n[(rr-1)+','+cc]=this.cells[k];}this.cells=n;return this}}
class Range{constructor(sh,r,c,nr,nc){Object.assign(this,{sh,r,c,nr,nc})}
 getValues(){const o=[];for(let i=0;i<this.nr;i++){const row=[];for(let j=0;j<this.nc;j++)row.push(this.sh.get(this.r+i,this.c+j));o.push(row)}return o}
 getDisplayValues(){return this.getValues().map(r=>r.map(v=>v instanceof Date?v.toISOString().slice(0,10).replace(/-/g,'/'):String(v)))}
 setValues(v){v.forEach((row,i)=>row.forEach((x,j)=>{if(typeof x==='string'&&x[0]==="'")x=x.slice(1);this.sh.set(this.r+i,this.c+j,x)}));return this}
 setValue(x){return this.setValues([[x]])} getValue(){return this.sh.get(this.r,this.c)} clearContent(){for(let i=0;i<this.nr;i++)for(let j=0;j<this.nc;j++)this.sh.set(this.r+i,this.c+j,'');return this}
 setNumberFormat(){return this} sort(spec){const rows=this.getValues();rows.sort((a,b)=>{for(const k of spec){const i=k.column-this.c;if(a[i]<b[i])return -1;if(a[i]>b[i])return 1}return 0});this.setValues(rows.map(r=>r.map(v=>typeof v==='string'&&/^0/.test(v)?"'"+v:v)));return this} insertCheckboxes(){return this} clearDataValidations(){return this}
 setFontWeight(){return this} setFontSize(n){for(let i=0;i<this.nr;i++)for(let j=0;j<this.nc;j++)this.sh.fonts[(this.r+i)+','+(this.c+j)]=n;return this} setRichTextValue(rv){this.sh.rich[this.r+','+this.c]=rv;return this} getRichText(){return this.sh.rich[this.r+','+this.c]} getFontSize(){return this.sh.fonts[this.r+','+this.c]||10} setBackground(){return this} setFontColor(){return this}}
let ssid=1;const SSS={};class SS{constructor(name){this.sheets=[];this.tz='America/Los_Angeles';this.id='SS'+(ssid++);this.name=name;SSS[this.id]=this} getSpreadsheetTimeZone(){return this.tz} setSpreadsheetTimeZone(z){this.tz=z} getId(){return this.id} getUrl(){return 'https://docs.google.com/spreadsheets/d/'+this.id} getSheetByName(n){return this.sheets.find(s=>s.name===n)||null}
 insertSheet(n){const s=new Sheet(this,n);this.sheets.push(s);return s} setActiveSheet(){} moveActiveSheet(){} getNumSheets(){return this.sheets.length} deleteSheet(sh){this.sheets=this.sheets.filter(x=>x!==sh)} getSheets(){return this.sheets.slice()}}
function load(){
 const ss=new SS(); const props={};
 const tpl=ss.insertSheet('書式_成績表'); tpl.set(1,1,'高圧ガス容器検査成績表\n\n※容器の製造年月を目視確認の上、容器番号前にチェックを入れる'); tpl.set(5,2,'001'); tpl.set(5,4,3); tpl.set(5,6,',');
 const ctx={console,Date,Math,JSON,
  SpreadsheetApp:{getActiveSpreadsheet:()=>ss,openById:(id)=>SSS[id]||ss,create:(n)=>{const x=new SS(n);x.insertSheet('シート1');global.ARCHIVES.push(x);return x},flush(){},
   newTextStyle:()=>{const st={};const b={setFontSize(n){st.size=n;return b},setBold(x){st.bold=x;return b},build(){return st}};return b},
   newRichTextValue:()=>{const rv={runs:[]};const b={setText(t){rv.text=t;return b},setTextStyle(a,z,st){rv.runs.push({a,z,size:st.size,bold:st.bold});return b},build(){return rv}};return b}},
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]||null,setProperty:(k,v)=>props[k]=v,deleteProperty:k=>{delete props[k]}})},
  LockService:{getScriptLock:()=>({waitLock(){},tryLock(){return true},releaseLock(){}})},
  Session:{getActiveUser:()=>({getEmail:()=>'tester@example.com'})},
  Utilities:{formatDate:(d,tz,f)=>{const p=n=>('0'+n).slice(-2);return String(f).replace(/yyyy|yy|MM|dd|HH|mm|ss/g,t=>({yyyy:d.getFullYear(),yy:p(d.getFullYear()%100),MM:p(d.getMonth()+1),dd:p(d.getDate()),HH:p(d.getHours()),mm:p(d.getMinutes()),ss:p(d.getSeconds())})[t])},newBlob:(c,t,n)=>({name:n,type:t,content:c,getName(){return n}})},
  MailApp:{sendEmail:(to,sub,body,opt)=>{global.MAILS.push({to,sub,body,cc:opt.cc,name:opt.name,att:opt.attachments.map(a=>a.name)})}},
  UrlFetchApp:{fetch:()=>({getBlob:()=>({setName(n){this.name=n;return this}})})},
  ScriptApp:(()=>{const tr=[];global.TRIGGERS=tr;return {getScriptId:()=>global.SCRIPT_ID||'17mOYxTzPwzIsrILrYL7pH0Rn-jYOjQ8f52n5KZm0-0p-ojhxBG246vu4',getOAuthToken:()=>'t',getProjectTriggers:()=>tr.slice(),deleteTrigger:(t)=>{tr.splice(tr.indexOf(t),1)},
   newTrigger:(fn)=>{const o={fn};const b={timeBased:()=>b,everyDays:(n)=>{o.days=n;return b},atHour:(h)=>{o.hour=h;return b},inTimezone:(z)=>{o.tz=z;return b},create:()=>{const t={getHandlerFunction:()=>fn,o};tr.push(t);return t}};return b}}})(),
  HtmlService:{createHtmlOutput:(h)=>({getAs:(t)=>({html:h,type:t,name:'',setName(n){this.name=n;return this}})})},
  DriveApp:(()=>{const files={};let fid=1000000000,did=1;
   const pathOf=f=>f.parent?pathOf(f.parent)+'/'+f.name:'';
   const iter=l=>{let k=0;return {hasNext:()=>k<l.length,next:()=>l[k++]}};
   let clock=1;const mk=(name,parent)=>{const f={name,parent,kids:[],files:[],id:'D'+(did++),created:clock++,trashed:false,getId(){return this.id},getName(){return this.name},setName(n){this.name=n;return this},getDateCreated(){return new Date(this.created)},setTrashed(x){this.trashed=x;if(x&&this.parent)this.parent.kids=this.parent.kids.filter(z=>z!==this);return this},
     createFile(b){const id='file'+(fid++);const fo={blob:b,parent:f,trashed:false,updated:clock++,getId:()=>id,getName:()=>b.name,getLastUpdated(){return new Date(this.updated)},get path(){return pathOf(this.parent)+'/'+b.name},
       setTrashed(x){this.trashed=x},moveTo(d){this.parent.files=this.parent.files.filter(z=>z!==this);this.parent=d;d.files.push(this);return this}};
       files[id]=fo;f.files.push(fo);global.PDFBLOBS[pathOf(f)+'/'+b.name]=b;return {getUrl:()=>'https://drive.google.com/file/d/'+id+'/view',getId:()=>id}},
     getFiles(){return iter(f.files.filter(x=>!x.trashed))},getFolders(){return iter(f.kids.slice())},
     getFilesByName(n){return iter(f.files.filter(x=>!x.trashed&&x.blob.name===n))},
     getFoldersByName(n){return iter(f.kids.filter(k=>k.name===n))},
     createFolder(n){const k=mk(n,f);f.kids.push(k);return k},
     moveTo(d){f.parent.kids=f.parent.kids.filter(z=>z!==f);f.parent=d;d.kids.push(f);return f}};return f};
   const root=mk('ROOT',null);let base=null;global.DRIVE_ROOT=root;const allFolders=()=>{const o=[];const w=f=>{f.kids.forEach(k=>{o.push(k);w(k)})};w(root);return o};
   Object.defineProperty(global,'PDFS',{configurable:true,get:()=>Object.values(files).filter(x=>!x.trashed&&x.parent).map(x=>x.path),set:()=>{}});
   return {createFolder:(n)=>base=root.createFolder(n),getFolderById:(id)=>{const f=allFolders().find(x=>x.id===id);if(!f)throw new Error('none');return f},getFoldersByName:(n)=>iter(allFolders().filter(x=>x.name===n&&!x.trashed)),
     getFileById:(id)=>{if(SSS[id]){const x=SSS[id];return {moveTo(d){if(!x.fo){const fid2='ss'+id;x.fo={blob:{name:x.name},parent:d,trashed:false,updated:clock++,getId:()=>x.id,getName:()=>x.name,getLastUpdated(){return new Date(this.updated)},setTrashed(v){this.trashed=v},moveTo(dd){this.parent.files=this.parent.files.filter(z=>z!==this);this.parent=dd;dd.files.push(this);return this},get path(){return pathOf(this.parent)+'/'+x.name}};d.files.push(x.fo);files[fid2]=x.fo}else x.fo.moveTo(d);return this},getName:()=>x.name}};if(!files[id])throw new Error('no file');const fo=files[id];return {getBlob:()=>({name:fo.blob.name,setName(n){this.name=n;return this}}),moveTo:(d)=>fo.moveTo(d),getParents:()=>iter(fo.parent?[fo.parent]:[])}}}})()};
 global.ARCHIVES=[];global.PDFBLOBS={};
 global.MAILS=[];
 vm.createContext(ctx); vm.runInContext(fs.readFileSync(GAS+'/Code.js','utf8'),ctx);
 vm.runInContext(fs.readFileSync(GAS+'/Setup.js','utf8'),ctx);
 vm.runInContext(fs.readFileSync(GAS+'/Delivery.js','utf8'),ctx);
 vm.runInContext(fs.readFileSync(GAS+'/Archive.js','utf8'),ctx);
 vm.runInContext(fs.readFileSync(GAS+'/Saiban.js','utf8'),ctx);
 ensure=vm.runInContext('ensureSheet_',ctx); ensure(ss,'ロット',vm.runInContext('LOT_HEADERS',ctx)); ensure(ss,'入力記録',vm.runInContext('LOG_HEADERS',ctx));
 const st=ensure(ss,'設定',['項目','値','説明']); vm.runInContext('SETTING_ROWS',ctx).forEach((r,i)=>{st.set(i+2,1,r[0]);st.set(i+2,2,r[1]);});
 st.set(2,2,'nouhin@example.com');
 const w=ensure(ss,'担当者',['担当者名']); ctx.__oldWorkers=w; w.set(2,1,'山田'); w.set(3,1,'佐藤');
 return {ctx,ss};
}
module.exports={load};
