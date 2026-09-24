/*******************************************************
 MADRASA CMS — Google Apps Script Backend
 Sheet ID is configurable via Script Properties: SHEET_ID
 Required Web App: Execute as owner, access according to deployment.
*******************************************************/
const DEFAULT_SHEET_ID='1aQXnM2PeR82t1jm7LUyw3wEu_PeiCL2irbttl0n-4Ck';
const VERSION_KEY='DATA_VERSION';
const SESSION_TTL=21600;
const HEADERS={
Settings:['id','name','arabicName','logo','favicon','motto','phone','mobile','email','address','map','facebook','youtube','founded','principal','description','vision','heroTitle','seoTitle','seoDescription','theme','footer'],
About:['id','title','content','image','status'],Notices:['id','title','category','details','pdfUrl','publishDate','important','status'],
News:['id','title','author','thumbnail','short','details','category','tags','date','status'],Events:['id','title','description','image','startDate','endDate','location','status'],
Teachers:['id','name','photo','designation','department','subject','education','experience','mobile','email','joiningDate','status','bio'],
Staff:['id','name','photo','designation','department','mobile','email','joiningDate','status','bio'],
Students:['id','studentId','name','photo','roll','registration','class','department','session','gender','guardian','mobile','status'],
Courses:['id','name','department','class','description','duration','fee','status'],Departments:['id','name','description','head','status'],Classes:['id','name','section','shift','department','status'],
Results:['id','studentId','roll','name','class','exam','subject','marks','gpa','grade','result','status'],Routine:['id','type','class','day','period','subject','teacher','time','room','status'],
Admission:['id','name','father','mother','dob','gender','previousInstitution','previousResult','class','department','mobile','guardianMobile','address','photo','documents','status','createdAt'],
Gallery:['id','title','imageUrl','category','date','description','status'],Videos:['id','title','youtubeUrl','thumbnail','category','description','status'],
Articles:['id','title','author','thumbnail','short','details','category','tags','date','status'],FAQ:['id','question','answer','category','status'],
Contact:['id','name','email','mobile','subject','message','status','createdAt'],Donations:['id','method','accountName','accountNumber','instructions','qrUrl','status'],
Messages:['id','name','mobile','email','subject','message','status','createdAt'],Users:['id','username','passwordHash','role','status','createdAt'],Admin:['id','username','passwordHash','role','status','createdAt'],Logs:['id','username','action','sheet','recordId','createdAt','ip']
};
function prop(){return PropertiesService.getScriptProperties()}
function sheetId(){return prop().getProperty('SHEET_ID')||DEFAULT_SHEET_ID}
function db(){return SpreadsheetApp.openById(sheetId())}
function now(){return new Date().toISOString()}
function json(data){return ContentService.createTextOutput(JSON.stringify({success:true,timestamp:now(),version:getVersion(),data:data||{}})).setMimeType(ContentService.MimeType.JSON)}
function fail(message,code){return ContentService.createTextOutput(JSON.stringify({success:false,timestamp:now(),version:getVersion(),message,code:code||'ERROR'})).setMimeType(ContentService.MimeType.JSON)}
function getVersion(){return prop().getProperty(VERSION_KEY)||'1'}
function bump(){prop().setProperty(VERSION_KEY,String(Date.now()))}
function headersFor(n){return HEADERS[n]||['id','title','description','status']}
function getSheets(){return db().getSheets().map(s=>s.getName())}
function setupSheet(name,headers){let ss=db(),s=ss.getSheetByName(name)||ss.insertSheet(name);if(s.getLastRow()===0)s.appendRow(headers||headersFor(name));else setupSheetHeader(name,headers||headersFor(name));return s}
function setupSheetHeader(name,headers){let s=db().getSheetByName(name)||db().insertSheet(name);s.getRange(1,1,1,headers.length).setValues([headers]);return true}
function setupSheets(){Object.keys(HEADERS).forEach(n=>setupSheet(n,HEADERS[n]));if(!getVersion())bump();if(!getDefaultAdmin())createDefaultAdmin();return {sheets:getSheets(),version:getVersion()}}
function createHeaders(){Object.keys(HEADERS).forEach(n=>setupSheetHeader(n,HEADERS[n]));return true}
function getDefaultAdmin(){const s=db().getSheetByName('Admin');if(!s||s.getLastRow()<2)return null;return s.getRange(2,1,1,s.getLastColumn()).getValues()[0]}
function hashPassword(p){const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(p),Utilities.Charset.UTF_8);return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('')}
function createDefaultAdmin(){const s=setupSheet('Admin',HEADERS.Admin);s.appendRow(['ADM-1','admin',hashPassword('ChangeMeImmediately!'),'admin','active',now()]);setupSheet('Users',HEADERS.Users);bump()}
function getData(name){const s=db().getSheetByName(name);if(!s||s.getLastRow()<2)return [];const h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0];return s.getRange(2,1,s.getLastRow()-1,h.length).getValues().map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i]])))}
function getAllData(){setupSheets();const out={};Object.keys(HEADERS).forEach(n=>out[n.toLowerCase()]=getData(n));out.settings=getData('Settings')[0]||{};return out}
function saveData(name,obj){const s=setupSheet(name,headersFor(name));const h=headersFor(name);const id=obj.id||Utilities.getUuid();const row=h.map(k=>obj[k]!==undefined?obj[k]:(k==='id'?id:''));s.appendRow(row);bump();return {id}}
function updateData(name,id,obj){const s=db().getSheetByName(name);if(!s)throw Error('Sheet not found');const h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0];const values=s.getDataRange().getValues();for(let i=1;i<values.length;i++){if(String(values[i][0])===String(id)){s.getRange(i+1,1,1,h.length).setValues([h.map((k,j)=>obj[k]!==undefined?obj[k]:values[i][j])]);bump();return true}}throw Error('Record not found')}
function deleteData(name,id){const s=db().getSheetByName(name);if(!s)throw Error('Sheet not found');const v=s.getDataRange().getValues();for(let i=1;i<v.length;i++)if(String(v[i][0])===String(id)){s.deleteRow(i+1);bump();return true}throw Error('Record not found')}
function getSettings(){return getData('Settings')[0]||{}}
function getNotices(){return getData('Notices')} function getTeachers(){return getData('Teachers')} function getStudents(){return getData('Students')} function getCourses(){return getData('Courses')} function getResults(){return getData('Results')} function getGallery(){return getData('Gallery')} function getEvents(){return getData('Events')} function getNews(){return getData('News')} function getArticles(){return getData('Articles')} function getFAQ(){return getData('FAQ')}
function token(){return Utilities.getUuid()+'-'+Utilities.getUuid()}
function session(t){const c=CacheService.getScriptCache();return c.get('sess_'+t)}
function auth(t){return t&&session(t)}
function log(u,a,s,id){try{setupSheet('Logs',HEADERS.Logs).appendRow([Utilities.getUuid(),u,a,s,id||'',now(),''])}catch(e){}}
function doGate(){setupSheets();return true}
function syncData(){bump();return {version:getVersion()}}
function pushData(name,rows){rows.forEach(r=>saveData(name,r));return true}
function backupData(){const snap=JSON.stringify(getAllData());const s=setupSheet('Backup',['id','createdAt','payload']);s.appendRow([Utilities.getUuid(),now(),snap]);return true}
function checkAuth(payload){const u=String(payload.username||''),p=String(payload.password||'');const admins=getData('Admin');const a=admins.find(x=>x.username===u&&x.status==='active');if(!a||hashPassword(p)!==a.passwordHash)throw Error('Invalid credentials');const t=token();CacheService.getScriptCache().put('sess_'+t,JSON.stringify({username:u,role:a.role}),SESSION_TTL);return t}
function doGet(e){try{const action=(e.parameter&&e.parameter.action)||'getAllData';if(action==='login')return fail('Use POST for login','METHOD');if(action==='getAllData')return json(getAllData());if(action==='getSettings')return json(getSettings());const data=getAllData();return json(data)}catch(err){return fail(err.message)}}
function doPost(e){try{const p=JSON.parse((e.postData&&e.postData.contents)||'{}');const action=p.action,payload=p.payload||{};if(action==='login')return json({token:checkAuth(payload),user:payload.username});const u=auth(p.token);if(!u)return fail('Unauthorized','AUTH');let result=null;switch(action){case'setupSheets':result=setupSheets();break;case'saveData':result=saveData(payload.sheet,payload.data);log(JSON.parse(u).username,'create',payload.sheet,result.id);break;case'updateData':result=updateData(payload.sheet,payload.id,payload.data);log(JSON.parse(u).username,'update',payload.sheet,payload.id);break;case'deleteData':result=deleteData(payload.sheet,payload.id);log(JSON.parse(u).username,'delete',payload.sheet,payload.id);break;case'backupData':result=backupData();break;case'syncData':result=syncData();break;case'pushData':result=pushData(payload.sheet,payload.rows||[]);break;default:throw Error('Unknown action')}return json(result)}catch(err){return fail(err.message)}}
