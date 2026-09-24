const API_URL=localStorage.getItem('MADRASA_API_URL')||'https://script.google.com/macros/s/AKfycbyFr62Z-zPn0exJEYKkROFAjBpRNuATHO62scf-H8grDaOi8om4TXt2e0dK-rC7cAbR/exec';
let STORE={version:null,data:{}};
let DEFAULTS_READY=null;

async function ensureDefaults(){
  if(window.DEFAULT_DATA)return;
  DEFAULTS_READY=DEFAULTS_READY||new Promise(resolve=>{
    const s=document.createElement('script');
    s.src=location.pathname.startsWith('/admin/')?'../js/data.js':'js/data.js';
    s.onload=resolve;s.onerror=resolve;document.head.appendChild(s);
  });
  await DEFAULTS_READY;
}
function mergeDefaults(remote){
  const base=window.DEFAULT_DATA||{};
  const out={...base,...(remote||{})};
  Object.keys(base).forEach(k=>{
    const rv=out[k];
    if(Array.isArray(base[k])&&(!Array.isArray(rv)||rv.length===0))out[k]=base[k];
    else if(k==='settings'&&(!rv||!Object.keys(rv).length))out[k]=base[k];
  });
  return out;
}
function savedToken(){return sessionStorage.getItem('madrasa_token')||localStorage.getItem('madrasa_token')||''}
function isDemoSession(){return localStorage.getItem('madrasa_demo_mode')==='1'}
async function apiRequest(action,data={},token){
  await ensureDefaults();
  const realToken=token||savedToken();
  const options={cache:'no-store',signal:AbortSignal.timeout(7000)};
  const protectedActions=['login','saveData','updateData','deleteData','setupSheets','backupData','syncData','pushData'];

  // Demo sessions must never send the fake demo token to the protected backend.
  // Public getAllData remains a normal GET during demo mode.
  const usePost=protectedActions.includes(action)||(action==='getAllData'&&realToken&&!isDemoSession());
  if(usePost){
    options.method='POST';
    options.headers={'Content-Type':'text/plain;charset=utf-8'};
    options.body=JSON.stringify({
      action,
      payload:data,
      token:action==='login'?'':realToken
    });
  }else{
    const q=new URLSearchParams({action});
    try{
      const rr=await fetch(API_URL+'?'+q.toString(),options);
      if(!rr.ok)throw Error('HTTP '+rr.status);
      return await rr.json();
    }catch(e){
      if(action==='login')throw e;
      return {success:false,message:e.message,version:'demo',data:mergeDefaults({})};
    }
  }
  try{
    const r=await fetch(API_URL,options);
    if(!r.ok)throw Error('HTTP '+r.status);
    return await r.json();
  }catch(e){
    if(action==='login')throw e;
    return {success:false,message:e.message,version:'demo',data:mergeDefaults({})};
  }
}
async function fetchAllData(){
  return apiRequest('getAllData');
}
async function syncWebsiteData(){
  await ensureDefaults();
  try{
    const r=await fetchAllData();
    const merged=mergeDefaults(r.data||{});
    if(r.success&&r.version!==STORE.version){
      STORE={version:r.version,data:merged};
      window.dispatchEvent(new CustomEvent('datachange'));
    }else if(!STORE.version){
      STORE={version:r.version||'demo',data:merged};
      window.dispatchEvent(new CustomEvent('datachange'));
    }
    if(r.success)window.dispatchEvent(new CustomEvent('syncok'));
    else window.dispatchEvent(new CustomEvent('syncfail'));
    return {...r,data:merged};
  }catch(e){
    STORE.data=mergeDefaults(STORE.data);
    window.dispatchEvent(new CustomEvent('syncfail'));
    throw e;
  }
}
async function testApi(){
  try{
    const r=await apiRequest('ping');
    return r;
  }catch(e){
    return {success:false,message:e.message};
  }
}