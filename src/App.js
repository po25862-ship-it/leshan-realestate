import { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import { db } from "./firebase";
import { ref, onValue, set, off } from "firebase/database";

// shared=true: visible to all users of this artifact
const save = async (key, val, shared=false) => {
  if(shared) { fbSave(key, val); return; }
  try { localStorage.setItem(key, JSON.stringify(val)); } catch(_){}
};
const load = async (key, fb, shared=false) => {
  if(shared) return fb;
  try { const v = localStorage.getItem(key); if(v) return JSON.parse(v); } catch(_){}
  return fb;
};
const uid = () => Math.random().toString(36).slice(2,8); // v2.1

// ── Seed Data ─────────────────────────────────────────────────────────────────
const SEED_PROPS = [
  { id:"p1", name:"信義區精品兩房", price:1580, area:28, rooms:2, district:"信義區", floor:"中高樓層", type:"電梯大樓", layout:"2/2/1/1", features:"近捷運,有車位", notes:"", _fromSheet:false },
  { id:"p2", name:"大安區透天四房", price:3200, area:65, rooms:4, district:"大安區", floor:"透天", type:"透天厝", layout:"4/2/2/2", features:"近學區,有後院", notes:"", _fromSheet:false },
  { id:"p3", name:"內湖科技園套房", price:680, area:12, rooms:1, district:"內湖區", floor:"中樓層", type:"電梯大樓", layout:"1/1/1/0", features:"近捷運", notes:"", _fromSheet:false },
];

// ── Sheet Parser ──────────────────────────────────────────────────────────────
function parseSheetText(text) {
  if (!text||!text.trim()) return [];
  const lines = text.trim().split("\n").map(l=>l.replace("\r","")).filter(l=>l.trim());
  if (lines.length<2) return [];
  const headers = lines[0].split("\t").map(h=>h.trim());
  return lines.slice(1).map(line=>{
    const cols = line.split("\t");
    const obj = {};
    headers.forEach((h,i)=>{ obj[h]=(cols[i]||"").trim(); });
    return obj;
  }).filter(row=>headers.some(h=>row[h]&&row[h]!==""));
}

function rowToProperty(row, forceCategory="") {
  const str = v => { try { return (v===null||v===undefined||v instanceof Date)?"":String(v).trim(); } catch(_){return "";} };
  const rawName = str(row["社區(物件)"]||row["社區"]||row["物件名稱"]||"");
  const isUrl = rawName.startsWith("http")||rawName.startsWith("www");
  const district = str(row["區域"]||row["行政區"]||"");
  const type = str(row["型態"]||row["類型"]||"");
  const address = str(row["地址"]||"");
  const autoName = address?address.slice(0,15):(district&&type?district+"·"+type:"");
  const name = isUrl?autoName:(rawName||autoName);
  const url = str(row["網址"]||row["連結"]||row["url"]||(isUrl?rawName:""));
  const rawLayout = row["格局(房/廳/衛/陽)"]||row["格局"]||"";
  const layout = (rawLayout===null||rawLayout===undefined||rawLayout instanceof Date)?"":String(rawLayout).trim();
  const rooms = parseInt((layout||"0").split("/")[0])||0;
  const rawFloor = row["樓層"];
  const floor = (rawFloor===null||rawFloor===undefined||rawFloor instanceof Date)?"":String(rawFloor).trim();
  const rawParking = row["車位"];
  const parking = (rawParking===true||rawParking==="TRUE"||rawParking==="有")?"有":"無";
  const hasRent = !!(row["租金"]||row["月租"]);
  const propCategory = forceCategory||(hasRent?"出租":"售屋");
  const price = parseFloat(str(row["開價"]||row["售價"]||"0"))||0;
  const area = parseFloat(str(row["權狀坪數"]||row["坪數"]||"0"))||0;
  const rawRent = str(row["租金"]||row["月租"]||"");
  const rent = rawRent.includes("/")?rawRent.split("/")[0]:rawRent;
  const serviceFee = str(row["服務費"]||"");
  const netArea = str(row["扣車坪數"]||"");
  const unitPrice = str(row["單價/坪(扣車)"]||row["單價/坪"]||"");
  const showingType = str(row["帶看方式"]||"");
  const agent2 = str(row["開發"]||"");
  const special = str(row["特殊事項"]||"");
  const noteParts = [];
  if(agent2) noteParts.push("開發:"+agent2);
  if(netArea) noteParts.push("扣車坪:"+netArea);
  if(unitPrice) noteParts.push("單價:"+unitPrice+"萬/坪");
  return {id:uid(),name:name||"未命名",propCategory,price,area,rooms,district,floor,type,layout,parking,url,features:special,notes:noteParts.join(" · "),rent,serviceFee,address,netArea,unitPrice,agent2,showingType,_fromSheet:true};
}

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("dashboard");
  const [currentUser, setCurrentUser] = useState(null);
  const [loginName, setLoginName] = useState("");
  const [loginPass, setLoginPass] = useState("");
  const [loginInvite, setLoginInvite] = useState("");
  const [loginMode, setLoginMode] = useState("login");
  const [loginError, setLoginError] = useState("");
  const [accounts, setAccounts] = useState({});
  const [properties, setProperties] = useState([]);
  const [buyers, setBuyers] = useState([]); // shared buyers (no phone)
  const [myClients, setMyClients] = useState([]); // private clients (with phone)
  const [showings, setShowings] = useState([]);
  const [events, setEvents] = useState([]);
  const MAY_SCHEDULE = [{"id": "6115jo", "title": "勞動節", "date": "2026-05-01", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "流通會議 | 值班：采萱 | 休假：韋伶", "agent": "系統", "isSchedule": true, "duty": "采萱", "offDuty": "韋伶"}, {"id": "9555ph", "title": "休假日", "date": "2026-05-02", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：慶祥 | 休假：五哥、韋伶、采萱", "agent": "系統", "isSchedule": true, "duty": "慶祥", "offDuty": "五哥、韋伶、采萱"}, {"id": "2ybxkc", "title": "休假日", "date": "2026-05-03", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：筱涵 | 休假：五哥", "agent": "系統", "isSchedule": true, "duty": "筱涵", "offDuty": "五哥"}, {"id": "afglj7", "title": "店務", "date": "2026-05-04", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：韋伶", "agent": "系統", "isSchedule": true, "duty": "韋伶", "offDuty": ""}, {"id": "2y08jg", "title": "成長大會", "date": "2026-05-05", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "區月會 | 值班：哲嘉", "agent": "系統", "isSchedule": true, "duty": "哲嘉", "offDuty": ""}, {"id": "mgji0g", "title": "店內教育訓練", "date": "2026-05-06", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：慶祥", "agent": "系統", "isSchedule": true, "duty": "慶祥", "offDuty": ""}, {"id": "z3vip8", "title": "店務", "date": "2026-05-07", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：五哥 | 休假：慶祥、筱涵、哲嘉", "agent": "系統", "isSchedule": true, "duty": "五哥", "offDuty": "慶祥、筱涵、哲嘉"}, {"id": "zbcs6q", "title": "集體看屋", "date": "2026-05-08", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：韋伶 | 休假：慶祥、采萱", "agent": "系統", "isSchedule": true, "duty": "韋伶", "offDuty": "慶祥、采萱"}, {"id": "av6pdm", "title": "休假日", "date": "2026-05-09", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：筱涵 | 休假：五哥、韋伶、采萱", "agent": "系統", "isSchedule": true, "duty": "筱涵", "offDuty": "五哥、韋伶、采萱"}, {"id": "darp17", "title": "休假日", "date": "2026-05-10", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：哲嘉 | 休假：韋伶", "agent": "系統", "isSchedule": true, "duty": "哲嘉", "offDuty": "韋伶"}, {"id": "1z2rf1", "title": "店務", "date": "2026-05-11", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：采萱 | 休假：筱涵、哲嘉", "agent": "系統", "isSchedule": true, "duty": "采萱", "offDuty": "筱涵、哲嘉"}, {"id": "fvl13v", "title": "店務", "date": "2026-05-12", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：慶祥 | 休假：筱涵、哲嘉", "agent": "系統", "isSchedule": true, "duty": "慶祥", "offDuty": "筱涵、哲嘉"}, {"id": "qyoadb", "title": "店務", "date": "2026-05-13", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：哲嘉", "agent": "系統", "isSchedule": true, "duty": "哲嘉", "offDuty": ""}, {"id": "0awe9k", "title": "店務", "date": "2026-05-14", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：五哥 | 休假：慶祥", "agent": "系統", "isSchedule": true, "duty": "五哥", "offDuty": "慶祥"}, {"id": "d7x84e", "title": "流通會議", "date": "2026-05-15", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "店內教育訓練 | 值班：韋伶", "agent": "系統", "isSchedule": true, "duty": "韋伶", "offDuty": ""}, {"id": "o07ijg", "title": "休假日", "date": "2026-05-16", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：筱涵 | 休假：五哥、韋伶", "agent": "系統", "isSchedule": true, "duty": "筱涵", "offDuty": "五哥、韋伶"}, {"id": "6jh8ai", "title": "休假日", "date": "2026-05-17", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "南崁區教育訓練 | 值班：韋伶 | 休假：采萱", "agent": "系統", "isSchedule": true, "duty": "韋伶", "offDuty": "采萱"}, {"id": "zs7qay", "title": "店務", "date": "2026-05-18", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：慶祥 | 休假：筱涵、哲嘉", "agent": "系統", "isSchedule": true, "duty": "慶祥", "offDuty": "筱涵、哲嘉"}, {"id": "eu1gb4", "title": "店務", "date": "2026-05-19", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：采萱 | 休假：筱涵、哲嘉", "agent": "系統", "isSchedule": true, "duty": "采萱", "offDuty": "筱涵、哲嘉"}, {"id": "u2r2yr", "title": "店務", "date": "2026-05-20", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：筱涵", "agent": "系統", "isSchedule": true, "duty": "筱涵", "offDuty": ""}, {"id": "h8eal4", "title": "店務", "date": "2026-05-21", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：五哥 | 休假：慶祥", "agent": "系統", "isSchedule": true, "duty": "五哥", "offDuty": "慶祥"}, {"id": "uaujhn", "title": "店內教育訓練", "date": "2026-05-22", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：韋伶 | 休假：慶祥", "agent": "系統", "isSchedule": true, "duty": "韋伶", "offDuty": "慶祥"}, {"id": "nxemop", "title": "店務", "date": "2026-05-23", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：哲嘉 | 休假：五哥、韋伶、采萱", "agent": "系統", "isSchedule": true, "duty": "哲嘉", "offDuty": "五哥、韋伶、采萱"}, {"id": "mqv57w", "title": "休假日", "date": "2026-05-24", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "集體看屋 | 值班：慶祥", "agent": "系統", "isSchedule": true, "duty": "慶祥", "offDuty": ""}, {"id": "ogdsbi", "title": "休假日", "date": "2026-05-25", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：采萱 | 休假：筱涵、哲嘉", "agent": "系統", "isSchedule": true, "duty": "采萱", "offDuty": "筱涵、哲嘉"}, {"id": "07v7pk", "title": "店務", "date": "2026-05-26", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：五哥 | 休假：筱涵、哲嘉、采萱", "agent": "系統", "isSchedule": true, "duty": "五哥", "offDuty": "筱涵、哲嘉、采萱"}, {"id": "f7uzdc", "title": "店務", "date": "2026-05-27", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：筱涵", "agent": "系統", "isSchedule": true, "duty": "筱涵", "offDuty": ""}, {"id": "5wl3e7", "title": "店內教育訓練", "date": "2026-05-28", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：哲嘉 | 休假：慶祥", "agent": "系統", "isSchedule": true, "duty": "哲嘉", "offDuty": "慶祥"}, {"id": "ak4gwk", "title": "店務", "date": "2026-05-29", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：韋伶", "agent": "系統", "isSchedule": true, "duty": "韋伶", "offDuty": ""}, {"id": "ckjbq3", "title": "休假日", "date": "2026-05-30", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "流通會議 | 值班：慶祥 | 休假：五哥、韋伶、采萱", "agent": "系統", "isSchedule": true, "duty": "慶祥", "offDuty": "五哥、韋伶、采萱"}, {"id": "0vswpc", "title": "休假日", "date": "2026-05-31", "time": "", "buyerId": "", "buyerName": "", "propertyId": "", "propName": "", "notes": "值班：采萱 | 休假：五哥、慶祥", "agent": "系統", "isSchedule": true, "duty": "采萱", "offDuty": "五哥、慶祥"}];
  const [ready, setReady] = useState(false);
  const saveProperties = (data) => { setProperties(data); const obj={}; data.forEach(p=>{obj[p.id]=p;}); fbSave("re_props3",obj); };
  const saveBuyers = (data) => { setBuyers(data); const obj={}; data.forEach(b=>{obj[b.id]=b;}); fbSave("re_buyers",obj); };
  const saveShowings = (data) => { setShowings(data); const obj={}; data.forEach(s=>{obj[s.id]=s;}); fbSave("re_showings3",obj); };
  const saveEvents = (data) => { setEvents(data); const userEvs=data.filter(e=>!e.isSchedule); const obj={}; userEvs.forEach(e=>{obj[e.id]=e;}); fbSave("re_events",obj); };

  useEffect(()=>{
    try { const s=localStorage.getItem("re_session"); if(s){const u=JSON.parse(s);if(u&&u.user)setCurrentUser(u.user);} } catch(_){}
    const u1=fbListen("re_accounts", data=>{if(data)setAccounts(data);});
    const u2=fbListen("re_props3", data=>{setProperties(data?(Array.isArray(data)?data:Object.values(data)):SEED_PROPS);});
    const u3=fbListen("re_buyers", data=>{setBuyers(data?(Array.isArray(data)?data:Object.values(data)):[]);});
    const u4=fbListen("re_showings3", data=>{setShowings(data?(Array.isArray(data)?data:Object.values(data)):[]);});
    const u5=fbListen("re_events", data=>{
      const userEvs=data?(Array.isArray(data)?data:Object.values(data)).filter(e=>!e.isSchedule):[];
      setEvents([...MAY_SCHEDULE,...userEvs]);
    });
    setReady(true);
    return ()=>{u1();u2();u3();u4();u5();};
  },[]);

  useEffect(()=>{ if(currentUser) try{const v=localStorage.getItem("re_myclients_"+currentUser);if(v)setMyClients(JSON.parse(v));}catch(_){} },[currentUser]);
  useEffect(()=>{ if(currentUser&&ready){try{localStorage.setItem("re_myclients_"+currentUser,JSON.stringify(myClients));}catch(_){}} },[myClients,currentUser,ready]);

  const doLogin = ()=>{
    const name=loginName.trim(); const pass=loginPass.trim();
    if(!name||!pass){setLoginError("請輸入姓名和密碼");return;}
    if(loginMode==="register"){
      if(loginInvite.trim()!=="3288283"){setLoginError("邀請碼錯誤");return;}
      if(accounts[name]){setLoginError("此姓名已被註冊，請直接登入");return;}
      fbSave("re_accounts",{...accounts,[name]:pass});
      localStorage.setItem("re_session",JSON.stringify({user:name,pass}));
      setCurrentUser(name);
    } else {
      if(!accounts[name]){setLoginError("帳號不存在，請先選「首次註冊」");return;}
      if(accounts[name]!==pass){setLoginError("密碼錯誤");return;}
      localStorage.setItem("re_session",JSON.stringify({user:name,pass}));
      setCurrentUser(name);
    }
  };
  const doLogout = ()=>{ try{localStorage.removeItem("re_session");}catch(_){} setCurrentUser(null); setMyClients([]); setLoginName(""); setLoginPass(""); };

  const TABS = [
    {id:"dashboard",    icon:"📊", label:"總覽"},
    {id:"buyers",       icon:"👥", label:"買方"},
    {id:"sharedbuyers", icon:"🏪", label:"店內"},
    {id:"showings",     icon:"🏃", label:"帶看"},
    {id:"calendar",     icon:"📅", label:"行事曆"},
    {id:"match",        icon:"🤖", label:"配對"},
    {id:"properties",   icon:"🏠", label:"物件"},
  ];

  if(!ready) return <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"#ffffff",color:"#e8722a",fontFamily:"sans-serif"}}>載入中…</div>;

  if(!currentUser) return (
    <>
      <style>{CSS}</style>
      <div className="shell" style={{justifyContent:"center",alignItems:"center",padding:"40px 24px"}}>
        <div style={{width:"100%",maxWidth:360}}>
          <div style={{textAlign:"center",marginBottom:32}}>
            <div style={{fontSize:48,marginBottom:10}}>🏠</div>
            <div style={{fontFamily:"DM Serif Display,serif",fontSize:28,color:"#1a1a1a"}}>捷運<span style={{color:"#e8722a"}}>樂善店</span></div>
            <div style={{fontSize:11,color:"#aaaaaa",marginTop:6,letterSpacing:2}}>REAL ESTATE PRO</div>
          </div>
          <div style={{display:"flex",background:"#f5f0eb",borderRadius:12,padding:4,marginBottom:20}}>
            {["login","register"].map(m=>(
              <button key={m} onClick={()=>{setLoginMode(m);setLoginError("");}} style={{flex:1,padding:"10px",background:loginMode===m?"#ffffff":"none",border:"none",borderRadius:10,fontFamily:"Noto Sans TC,sans-serif",fontSize:14,fontWeight:loginMode===m?700:400,color:loginMode===m?"#e8722a":"#888888",cursor:"pointer"}}>
                {m==="login"?"登入":"首次註冊"}
              </button>
            ))}
          </div>
          <div className="field-wrap"><label className="field-label">姓名</label>
            <input placeholder="例：采萱" value={loginName} onChange={e=>{setLoginName(e.target.value);setLoginError("");}}/>
          </div>
          <div className="field-wrap"><label className="field-label">密碼</label>
            <input type="password" placeholder="輸入密碼" value={loginPass} onChange={e=>{setLoginPass(e.target.value);setLoginError("");}} onKeyDown={e=>{if(e.key==="Enter")doLogin();}}/>
          </div>
          {loginMode==="register"&&(<div className="field-wrap"><label className="field-label">店內邀請碼</label>
            <input placeholder="請輸入邀請碼" value={loginInvite} onChange={e=>{setLoginInvite(e.target.value);setLoginError("");}}/>
          </div>)}
          {loginError&&<div style={{background:"#fff3f3",border:"1px solid #ffcdd2",borderRadius:10,padding:"10px 14px",color:"#ff3b30",fontSize:13,marginBottom:12}}>❌ {loginError}</div>}
          <button className="btn-gold" onClick={doLogin} disabled={!loginName.trim()||!loginPass.trim()}>
            {loginMode==="login"?"登入 →":"註冊並登入 →"}
          </button>
          <div style={{fontSize:11,color:"#aaaaaa",textAlign:"center",marginTop:16,lineHeight:2}}>
            ✦ 我的客戶資料只有自己看得到<br/>✦ 物件庫、帶看、行事曆全員共用
          </div>
        </div>
      </div>
    </>
  );


  return (
    <>
      <style>{CSS}</style>
      <div className="shell">
        <header className="topbar">
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div>
              <div className="topbar-eyebrow">捷運樂善店</div>
              <div className="topbar-title">捷運<span>樂善店</span></div>
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:12,color:"#e8722a",marginBottom:4}}>👤 {currentUser}</div>
              <button onClick={doLogout} style={{background:"none",border:"1px solid #d5c9b8",borderRadius:8,color:"#888888",fontSize:11,padding:"4px 8px",cursor:"pointer",fontFamily:"Noto Sans TC,sans-serif"}}>登出</button>
            </div>
          </div>
        </header>
        <main className="body">
          {tab==="dashboard"  && <Dashboard properties={properties} buyers={buyers} showings={showings} events={events} setTab={setTab} currentUser={currentUser}/>}
          {tab==="buyers"     && <Buyers buyers={buyers} setBuyers={saveBuyers} myClients={myClients} setMyClients={setMyClients} properties={properties} showings={showings} setShowings={saveShowings} events={events} setEvents={saveEvents} currentUser={currentUser}/>}
          {tab==="showings"    && <ShowingsPage showings={showings} setShowings={saveShowings} buyers={buyers} properties={properties} currentUser={currentUser}/>}
          {tab==="calendar"   && <CalendarView events={events} setEvents={saveEvents} buyers={buyers} properties={properties} currentUser={currentUser}/>}
          {tab==="sharedbuyers" && <SharedBuyers buyers={buyers} setBuyers={saveBuyers} currentUser={currentUser} myClients={myClients}/>}
          {tab==="match"      && <AIMatch properties={properties} buyers={buyers}/>}
          {tab==="loan"       && <LoanCalc/>}
          {tab==="properties" && <Properties properties={properties} setProperties={saveProperties} showings={showings} buyers={buyers}/>}
        </main>
        <nav className="bottomnav">
          {TABS.map(t=>(
            <button key={t.id} className={"navbtn"+(tab===t.id?" active":"")} onClick={()=>setTab(t.id)}>
              <span className="navicon">{t.icon}</span>
              <span className="navlabel">{t.label}</span>
            </button>
          ))}
        </nav>
      </div>
    </>
  );
}

// ── Dashboard ─────────────────────────────────────────────────────────────────
function Dashboard({ properties, buyers, showings, events, setTab, currentUser }) {
  const today = new Date().toISOString().slice(0,10);
  const now = new Date().toISOString().slice(0,7);
  const todayEvents = events.filter(e=>e.date===today);
  const monthShowings = showings.filter(s=>s.date&&s.date.startsWith(now));
  const active = buyers.filter(b=>["初看","複看","斡旋中"].includes(b.status));
  const recent = [...showings].sort((a,b)=>(b.date||"").localeCompare(a.date||"")).slice(0,3);

  return (
    <div className="page">
      <div className="page-title">業務總覽</div>
      <div className="stat-grid">
        {[
          {icon:"🏠",val:properties.length,label:"在售物件",a:true},
          {icon:"👥",val:buyers.length,label:"買方總數"},
          {icon:"🔥",val:active.length,label:"洽談中",a:true},
          {icon:"🏃",val:monthShowings.length,label:"本月帶看"},
        ].map((s,i)=>(
          <div key={i} className={"stat-card"+(s.a?" stat-accent":"")}>
            <div className="stat-icon">{s.icon}</div>
            <div className="stat-val">{s.val}</div>
            <div className="stat-label">{s.label}</div>
          </div>
        ))}
      </div>

      {/* Showings stats by agent */}
      {showings.length>0 && (()=>{
        const agentMap = {};
        showings.filter(s=>s.date&&s.date.startsWith(now)).forEach(s=>{
          const a = s.agent||"未知";
          agentMap[a] = (agentMap[a]||0)+1;
        });
        const agentList = Object.entries(agentMap).sort((a,b)=>b[1]-a[1]);
        return agentList.length>0 ? (
          <>
            <div className="section-hd">本月帶看統計</div>
            <div style={{background:"#ffffff",border:"1px solid #e0d6ca",borderRadius:14,padding:"14px",marginBottom:16}}>
              {agentList.map(([agent,count])=>(
                <div key={agent} style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
                  <span style={{fontSize:13,fontWeight:600,color:"#1a1a1a",width:60}}>{agent}</span>
                  <div style={{flex:1,background:"#f5f0eb",borderRadius:20,height:20,overflow:"hidden"}}>
                    <div style={{height:"100%",background:"#e8722a",borderRadius:20,width:(count/Math.max(...agentList.map(x=>x[1]))*100)+"%",transition:"width .3s"}}/>
                  </div>
                  <span style={{fontSize:13,color:"#e8722a",fontWeight:700,width:24,textAlign:"right"}}>{count}</span>
                </div>
              ))}
            </div>
          </>
        ) : null;
      })()}

      {todayEvents.length>0 && <>
        <div className="section-hd">今日行程</div>
        {todayEvents.map(e=>(
          <div className="event-card today" key={e.id}>
            <div className="event-time">{e.time||"全天"}</div>
            <div className="event-main">
              <div className="event-title">{e.title}</div>
              {e.duty && <div className="event-sub">🧑‍💼 值班：{e.duty}</div>}
              {e.offDuty && <div className="event-sub">🏖️ 休假：{e.offDuty}</div>}
              {e.buyerName && <div className="event-sub">👤 {e.buyerName}</div>}
              {e.propName && <div className="event-sub">🏠 {e.propName}</div>}
            </div>
          </div>
        ))}
      </>}

      <div className="section-hd">最近帶看</div>
      {recent.length===0 ? <Empty icon="🏃" text="尚無帶看記錄"/> : recent.map(s=>{
        const p = properties.find(x=>x.id===s.propertyId);
        return (
          <div className="list-card" key={s.id}>
            <div className="list-card-main">
              <div className="list-card-title">{p?p.name:"已刪除物件"}</div>
              <div className="list-card-sub">{s.buyerName||"?"} · {s.date} · {s.agent||""}</div>
            </div>
            <IBadge v={s.interest}/>
          </div>
        );
      })}

      <div className="section-hd" style={{marginTop:20}}>快速入口</div>
      <div className="quick-grid">
        {[["👥","新增買方","buyers"],["📅","新增行程","calendar"],["🤖","AI配對","match"],["💰","貸款試算","loan"]].map(([icon,label,page])=>(
          <button key={page} className="quick-btn" onClick={()=>setTab(page)}>
            <span style={{fontSize:22}}>{icon}</span>
            <span style={{fontSize:12,marginTop:4}}>{label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

// ── Buyers ────────────────────────────────────────────────────────────────────
const PROP_TYPES = ["公寓","大樓","廠房","透天","土地","車位"];
const BUYER_STATUSES = ["尚未看屋","初看","複看","斡旋中","已成交"];

function Buyers({ buyers, setBuyers, myClients, setMyClients, properties, showings, setShowings, events, setEvents, currentUser }) {
  const [view, setView] = useState("list");
  const [selected, setSelected] = useState(null);
  const [form, setForm] = useState({});
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const upd = (k,v)=>setForm(p=>({...p,[k]:v}));
  const togglePropType = (t)=>{ const cur=form.propTypes||[]; upd("propTypes", cur.includes(t)?cur.filter(x=>x!==t):[...cur,t]); };

  const blank = ()=>({ id:uid(), name:"", phone:"", line:"", downPayment:"", budgetMin:"", budgetMax:"", propTypes:[], rooms1:"", rooms2:"", needs:"", status:"尚未看屋", level:"", purpose:"", motivation:"", urgentNeeds:"", pinned:false, notes:"", agent:currentUser, createdAt:new Date().toISOString().slice(0,10), shownProps:[], timeline:[] });

  const openNew = ()=>{ setForm(blank()); setSelected(null); setView("form"); };
  const openDetail = b=>{ setSelected(b); setView("detail"); };
  const openEdit = b=>{ setForm({...b, shownProps:b.shownProps||[]}); setSelected(b); setView("form"); };
  const doSave = ()=>{
    const now = new Date().toISOString().slice(0,16).replace("T"," ");
    const isNew = !buyers.find(x=>x.id===form.id);
    const newTimeline = isNew
      ? [{id:uid(), time:now, type:"create", note:"買方建立"}]
      : (form.timeline||[]);
    const existing = buyers.find(x=>x.id===form.id);
    let tl = newTimeline;
    if(!isNew && existing && form.notes !== existing.notes && form.notes) {
      tl = [...(existing.timeline||[]), {id:uid(), time:now, type:"note", note:"更新備註："+form.notes.slice(0,30)}];
    }
    // Save phone/line privately, strip from shared record
    if(form.phone || form.line) {
      save("re_contact_"+form.id, {phone:form.phone||"", line:form.line||""});
    }
    const sharedForm = {...form, phone:"", line:"", timeline: tl};
    const finalForm = {...form, timeline: tl};
    setBuyers(prev=>{ const e=prev.find(x=>x.id===sharedForm.id); return e?prev.map(x=>x.id===sharedForm.id?sharedForm:x):[...prev,sharedForm]; });
    // sync phone to myClients
    if(form.phone&&form.agent===currentUser) {
      setMyClients(prev=>{ const e=prev.find(x=>x.buyerId===form.id); const rec={id:e?e.id:uid(),buyerId:form.id,name:form.name,phone:form.phone,line:form.line,agent:currentUser}; return e?prev.map(x=>x.id===rec.id?rec:x):[...prev,rec]; });
    }
    setView(selected?"detail":"list"); if(!selected) setSelected(form);
  };
  const doDel = id=>{ setBuyers(p=>p.filter(x=>x.id!==id)); setMyClients(p=>p.filter(x=>x.buyerId!==id)); setView("list"); };

  // Add showing to a buyer
  const addShowing = (buyerId, propId, propName, feedback, interest)=>{
    const buyerName = buyers.find(b=>b.id===buyerId)?.name||"";
    const now = new Date().toISOString().slice(0,16).replace("T"," ");
    const s = { id:uid(), buyerId, buyerName, propertyId:propId, propName, feedback, interest, date:new Date().toISOString().slice(0,10), agent:currentUser };
    setShowings(prev=>[...prev, s]);
    const newShownProp = {propId, propName, feedback, interest, date:new Date().toISOString().slice(0,10)};
    const tlEntry = {id:uid(), time:now, type:"showing", note:"帶看："+propName+(feedback?" — "+feedback.slice(0,20):"")};
    setBuyers(prev=>prev.map(b=>{
      if(b.id!==buyerId) return b;
      const sp = Array.isArray(b.shownProps) ? b.shownProps : [];
      const tl = Array.isArray(b.timeline) ? b.timeline : [];
      const exists = sp.find(x=>x.propId===propId && propId);
      if(exists && propId) return {...b, shownProps:sp.map(x=>x.propId===propId?{...x,feedback,interest}:x), timeline:[...tl,tlEntry]};
      return {...b, shownProps:[...sp, newShownProp], timeline:[...tl,tlEntry]};
    }));
    setSelected(prev => {
      if(!prev || prev.id!==buyerId) return prev;
      const sp = Array.isArray(prev.shownProps)?prev.shownProps:[];
      const tl = Array.isArray(prev.timeline)?prev.timeline:[];
      return {...prev, shownProps:[...sp,newShownProp], timeline:[...tl,tlEntry]};
    });
  };

  const filtered = buyers.filter(b=>{
    if(b.agent !== currentUser) return false;
    if(filterStatus&&b.status!==filterStatus) return false;
    if(search&&!b.name.includes(search)&&!(b.needs||"").includes(search)) return false;
    return true;
  });

  if(view==="form") return <BuyerForm form={form} upd={upd} togglePropType={togglePropType} selected={selected} onSave={doSave} onDel={doDel} onBack={()=>setView(selected?"detail":"list")} currentUser={currentUser}/>;

  if(view==="detail") {
    const b = selected;
    const mc = myClients.find(x=>x.buyerId===b.id);
    const bShowings = showings.filter(s=>s.buyerId===b.id).sort((a,c)=>(c.date||"").localeCompare(a.date||""));
    return <BuyerDetail b={b} mc={mc} bShowings={bShowings} properties={properties} events={events} setEvents={saveEvents} onEdit={()=>openEdit(b)} onBack={()=>setView("list")} addShowing={addShowing} currentUser={currentUser} setSelected={setSelected} setBuyers={saveBuyers} showings={showings} setShowings={saveShowings}/>;
  }

  return (
    <div className="page">
      <div className="page-nav">
        <div className="page-title" style={{marginBottom:0}}>我的買方（{filtered.length}）</div>
        <button className="add-btn" onClick={openNew}>＋ 新增</button>
      </div>
      <div style={{display:"flex",gap:8,marginBottom:12}}>
        <input placeholder="搜尋姓名或需求…" value={search} onChange={e=>setSearch(e.target.value)} style={{flex:1,padding:"9px 12px",fontSize:13}}/>
        <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} style={{width:90,padding:"9px 8px",fontSize:12}}>
          <option value="">全部</option>
          {BUYER_STATUSES.map(s=><option key={s}>{s}</option>)}
        </select>
      </div>
      {filtered.length===0 ? <Empty icon="👥" text="沒有符合的買方"/> : filtered.map(b=>(
        <div className="list-card" key={b.id} onClick={()=>openDetail(b)}>
          <div className="list-card-main">
            <div className="list-card-title">{b.name}</div>
            <div className="list-card-sub">
              {[b.budgetMin&&b.budgetMax?b.budgetMin+"~"+b.budgetMax+"萬":b.budgetMax?b.budgetMax+"萬以內":"", b.rooms1?b.rooms1+(b.rooms2?"-"+b.rooms2:"")+"房":"", b.needs].filter(Boolean).join(" · ")}
            </div>
            {(b.propTypes||[]).length>0 && <div className="list-card-note">{(b.propTypes||[]).join("、")}</div>}
          </div>
          <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>
            <BSBadge v={b.status}/>
            <div style={{fontSize:10,color:"#4b5563"}}>{b.agent}</div>
          </div>
        </div>
      ))}
    </div>
  );
}

function BuyerForm({ form, upd, togglePropType, selected, onSave, onDel, onBack, currentUser }) {
  return (
    <div className="page">
      <div className="page-nav">
        <button className="back-btn" onClick={onBack}>← 返回</button>
        <div className="page-title" style={{marginBottom:0}}>{selected?"編輯買方":"新增買方"}</div>
      </div>
      <div className="section-hd">基本資料</div>
      <div className="field-wrap"><label className="field-label">姓名 *</label><input placeholder="王先生" value={form.name||""} onChange={e=>upd("name",e.target.value)}/></div>
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">電話（僅自己可見）</label><input placeholder="0912-345-678" value={form.phone||""} onChange={e=>upd("phone",e.target.value)}/></div>
        <div className="field-wrap"><label className="field-label">Line ID（僅自己可見）</label><input placeholder="line123" value={form.line||""} onChange={e=>upd("line",e.target.value)}/></div>
      </div>
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">買方級別</label>
          <select value={form.level||""} onChange={e=>upd("level",e.target.value)}>
            <option value="">未分級</option>
            <option value="A買">A買</option>
            <option value="B買">B買</option>
            <option value="C買">C買</option>
            <option value="其他">其他</option>
          </select>
        </div>
        <div className="field-wrap"><label className="field-label">狀態</label>
          <select value={form.status||"尚未看屋"} onChange={e=>upd("status",e.target.value)}>
            {BUYER_STATUSES.map(s=><option key={s}>{s}</option>)}
          </select>
        </div>
      </div>
      <div className="section-hd">購屋條件</div>
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">自備款（萬）</label><input type="number" placeholder="300" value={form.downPayment||""} onChange={e=>upd("downPayment",e.target.value)}/></div>
        <div className="field-wrap"><label className="field-label">狀態</label>
          <select value={form.status||"尚未看屋"} onChange={e=>upd("status",e.target.value)}>
            {BUYER_STATUSES.map(s=><option key={s}>{s}</option>)}
          </select>
        </div>
      </div>
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">預算下限（萬）</label><input type="number" placeholder="800" value={form.budgetMin||""} onChange={e=>upd("budgetMin",e.target.value)}/></div>
        <div className="field-wrap"><label className="field-label">預算上限（萬）</label><input type="number" placeholder="1500" value={form.budgetMax||""} onChange={e=>upd("budgetMax",e.target.value)}/></div>
      </div>
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">最少房數</label>
          <select value={form.rooms1||""} onChange={e=>upd("rooms1",e.target.value)}>
            <option value="">不限</option>
            {["1","2","3","4","5"].map(v=><option key={v}>{v}</option>)}
          </select>
        </div>
        <div className="field-wrap"><label className="field-label">最多房數</label>
          <select value={form.rooms2||""} onChange={e=>upd("rooms2",e.target.value)}>
            <option value="">不限</option>
            {["1","2","3","4","5"].map(v=><option key={v}>{v}</option>)}
          </select>
        </div>
      </div>
      <div className="field-wrap">
        <label className="field-label">需求形態（可複選）</label>
        <div className="chip-group">
          {PROP_TYPES.map(t=>(
            <button key={t} className={"chip"+((form.propTypes||[]).includes(t)?" active":"")} onClick={()=>togglePropType(t)}>{t}</button>
          ))}
        </div>
      </div>
      <div className="field-wrap"><label className="field-label">需求地點</label><input placeholder="大安區、信義區…" value={form.needs||""} onChange={e=>upd("needs",e.target.value)}/></div>
      <div className="field-wrap">
        <label className="field-label">明確客需（置頂顯示於店內買方）</label>
        <textarea rows={2} placeholder="例：A7 兩房以上 預算1200萬內 需車位 急！" value={form.urgentNeeds||""} onChange={e=>upd("urgentNeeds",e.target.value)}/>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12,background:form.pinned?"#fff3f0":"#f5f0eb",border:"1px solid",borderColor:form.pinned?"#e8722a":"#e0d6ca",borderRadius:10,padding:"12px 14px",cursor:"pointer"}} onClick={()=>upd("pinned",!form.pinned)}>
        <div style={{width:22,height:22,borderRadius:6,background:form.pinned?"#e8722a":"#ffffff",border:"2px solid",borderColor:form.pinned?"#e8722a":"#cccccc",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
          {form.pinned&&<span style={{color:"#fff",fontSize:14,fontWeight:700}}>✓</span>}
        </div>
        <div>
          <div style={{fontSize:13,fontWeight:700,color:form.pinned?"#e8722a":"#666666"}}>📌 置頂紅框顯示於店內買方</div>
          <div style={{fontSize:11,color:"#aaaaaa",marginTop:2}}>開啟後此買方需求會在店內買方列表最上方以紅框標示</div>
        </div>
      </div>
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">用途</label>
          <select value={form.purpose||""} onChange={e=>upd("purpose",e.target.value)}>
            <option value="">不限</option>
            <option value="住宅">住宅</option>
            <option value="辦公">辦公</option>
            <option value="店面">店面</option>
          </select>
        </div>
        <div className="field-wrap"><label className="field-label">動機</label>
          <select value={form.motivation||""} onChange={e=>upd("motivation",e.target.value)}>
            <option value="">不限</option>
            <option value="自用">自用</option>
            <option value="投資">投資</option>
          </select>
        </div>
      </div>
      <div className="field-wrap"><label className="field-label">備註</label><textarea rows={3} value={form.notes||""} onChange={e=>upd("notes",e.target.value)}/></div>
      <button className="btn-gold" onClick={onSave} disabled={!form.name}>💾 儲存</button>
      {selected && <button className="btn-danger" style={{marginTop:10}} onClick={()=>onDel(selected.id)}>🗑 刪除此買方</button>}
    </div>
  );
}

function BuyerDetail({ b, mc, bShowings, properties, events, setEvents, onEdit, onBack, addShowing, currentUser, setSelected, setBuyers, showings, setShowings }) {
  const [showAddShowing, setShowAddShowing] = useState(false);
  const [privateContact, setPrivateContact] = useState({phone:"",line:""});
  useEffect(()=>{
    try{const v=localStorage.getItem("re_contact_"+b.id);if(v)setPrivateContact(JSON.parse(v));else setPrivateContact({phone:"",line:""});}catch(_){setPrivateContact({phone:"",line:""});}
  },[b.id]);
  const [showAddEvent, setShowAddEvent] = useState(false);
  const [selPropId, setSelPropId] = useState("");
  const [feedback, setFeedback] = useState("");
  const [interest, setInterest] = useState("普通");
  const [evDate, setEvDate] = useState(new Date().toISOString().slice(0,10));
  const [evTime, setEvTime] = useState("10:00");
  const [evNote, setEvNote] = useState("");

  const [manualPropName, setManualPropName] = useState("");
  const doAddShowing = ()=>{
    const p = properties.find(x=>x.id===selPropId);
    const propName = p ? p.name : manualPropName;
    if(!propName.trim()) return;
    addShowing(b.id, selPropId||"", propName, feedback, interest);
    setSelPropId(""); setManualPropName(""); setFeedback(""); setInterest("普通"); setShowAddShowing(false);
  };

  const doAddEvent = ()=>{
    const p = selPropId ? properties.find(x=>x.id===selPropId) : null;
    const ev = { id:uid(), title:"帶看："+b.name, date:evDate, time:evTime, buyerId:b.id, buyerName:b.name, propertyId:selPropId||"", propName:p?p.name:"", notes:evNote, agent:currentUser };
    setEvents(prev=>[...prev,ev]);
    setShowAddEvent(false); setEvNote(""); setSelPropId("");
  };

  const bEvents = events.filter(e=>e.buyerId===b.id).sort((a,c)=>(a.date||"").localeCompare(c.date||""));

  return (
    <div className="page">
      <div className="page-nav">
        <button className="back-btn" onClick={onBack}>← 返回</button>
        <button className="edit-btn" onClick={onEdit}>編輯</button>
      </div>
      <div className="detail-hero">
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div className="detail-name">{b.name}</div>
          <BSBadge v={b.status}/>
        </div>
        <div style={{fontSize:12,color:"#888888"}}>負責：{b.agent} · {b.createdAt}</div>
        {mc && <>
          {mc.phone && <a href={"tel:"+mc.phone} className="phone-link">📞 {mc.phone}</a>}
          {mc.line && <div style={{fontSize:13,color:"#e8722a"}}>💬 Line: {mc.line}</div>}
        </>}
        {!mc && b.agent!==currentUser && <div style={{fontSize:12,color:"#f87171"}}>🔒 電話由 {b.agent} 保管</div>}
      </div>

      <div className="info-grid">
        {b.downPayment && <InfoItem label="自備款" val={b.downPayment+"萬"}/>}
        {(b.budgetMin||b.budgetMax) && <InfoItem label="預算" val={(b.budgetMin||"不限")+"~"+(b.budgetMax||"不限")+"萬"}/>}
        {(b.rooms1||b.rooms2) && <InfoItem label="房數" val={(b.rooms1||"不限")+"~"+(b.rooms2||"不限")+"房"}/>}
        {b.needs && <InfoItem label="需求地點" val={b.needs}/>}
      </div>
      {(b.propTypes||[]).length>0 && (
        <div className="info-box">
          <div className="info-box-label">需求形態</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:6,marginTop:4}}>
            {(b.propTypes||[]).map(t=><span key={t} className="chip active" style={{cursor:"default"}}>{t}</span>)}
          </div>
        </div>
      )}
      {b.notes && <div className="info-box"><div className="info-box-label">備註</div>{b.notes}</div>}

      {/* Timeline */}
      <TimelineSection b={b} setBuyers={saveBuyers} setSelected={setSelected} currentUser={currentUser}/>

      {/* Shown Properties */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"16px 0 8px"}}>
        <div className="section-hd" style={{margin:0}}>看過物件（{(b.shownProps||[]).length}）</div>
        <button onClick={()=>setShowAddShowing(v=>!v)} className="add-btn" style={{fontSize:12,padding:"6px 12px"}}>＋ 新增帶看</button>
      </div>

      {showAddShowing && (
        <div className="mini-form">
          {properties.length > 0 ? (
            <div className="field-wrap"><label className="field-label">選擇物件</label>
              <select value={selPropId} onChange={e=>{ setSelPropId(e.target.value); if(e.target.value) setManualPropName(""); }}>
                <option value="">從物件庫選擇…</option>
                {properties.map(p=><option key={p.id} value={p.id}>{p.name} {p.price?p.price+"萬":""}</option>)}
              </select>
            </div>
          ) : null}
          {!selPropId && (
            <div className="field-wrap"><label className="field-label">{properties.length>0?"或直接輸入物件名稱":"物件名稱"}</label>
              <input placeholder="例：信義區兩房、新潤鉑麗…" value={manualPropName} onChange={e=>setManualPropName(e.target.value)}/>
            </div>
          )}
          <div className="field-wrap"><label className="field-label">客戶回饋</label>
            <textarea rows={2} placeholder="喜歡採光，覺得坪數稍小…" value={feedback} onChange={e=>setFeedback(e.target.value)}/>
          </div>
          <div className="field-wrap"><label className="field-label">購買意願</label>
            <select value={interest} onChange={e=>setInterest(e.target.value)}>
              {["高","中","低","普通","婉拒"].map(v=><option key={v}>{v}</option>)}
            </select>
          </div>
          <div style={{display:"flex",gap:8}}>
            <button className="btn-gold" style={{flex:1}} onClick={doAddShowing} disabled={!selPropId && !manualPropName.trim()}>新增</button>
            <button onClick={()=>setShowAddShowing(false)} style={{flex:1,background:"#1e2535",border:"none",borderRadius:12,color:"#666666",fontFamily:"Noto Sans TC,sans-serif",fontSize:14,cursor:"pointer"}}>取消</button>
          </div>
        </div>
      )}

      <ShownPropsSection b={b} properties={properties} setBuyers={saveBuyers} setSelected={setSelected}/>

      {/* Events */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"16px 0 8px"}}>
        <div className="section-hd" style={{margin:0}}>行程安排（{bEvents.length}）</div>
        <button onClick={()=>setShowAddEvent(v=>!v)} className="add-btn" style={{fontSize:12,padding:"6px 12px"}}>＋ 新增行程</button>
      </div>

      {showAddEvent && (
        <div className="mini-form">
          <div className="two-col">
            <div className="field-wrap"><label className="field-label">日期</label><input type="date" value={evDate} onChange={e=>setEvDate(e.target.value)}/></div>
            <div className="field-wrap"><label className="field-label">時間</label><input type="time" value={evTime} onChange={e=>setEvTime(e.target.value)}/></div>
          </div>
          <div className="field-wrap"><label className="field-label">物件（選填）</label>
            <select value={selPropId} onChange={e=>setSelPropId(e.target.value)}>
              <option value="">不指定</option>
              {properties.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>
          <div className="field-wrap"><label className="field-label">備註</label><input placeholder="第一次帶看、複看…" value={evNote} onChange={e=>setEvNote(e.target.value)}/></div>
          <div style={{display:"flex",gap:8}}>
            <button className="btn-gold" style={{flex:1}} onClick={doAddEvent}>新增行程</button>
            <button onClick={()=>setShowAddEvent(false)} style={{flex:1,background:"#1e2535",border:"none",borderRadius:12,color:"#666666",fontFamily:"Noto Sans TC,sans-serif",fontSize:14,cursor:"pointer"}}>取消</button>
          </div>
        </div>
      )}

      {bEvents.map(e=>{
        const p = properties.find(x=>x.id===e.propertyId);
        return (
          <div className="event-card" key={e.id} style={{alignItems:"flex-start"}}>
            <div className="event-time">{e.time||"全天"}</div>
            <div className="event-main" style={{flex:1}}>
              <div className="event-title">{e.date} {e.title}</div>
              {p && <div className="event-sub">🏠 {p.name}</div>}
              {e.notes && <div className="event-sub">{e.notes}</div>}
            </div>
            <button onClick={()=>setEvents(prev=>prev.filter(x=>x.id!==e.id))} style={{background:"none",border:"none",color:"#cccccc",fontSize:16,cursor:"pointer",padding:"0 4px",flexShrink:0}}>✕</button>
          </div>
        );
      })}
    </div>
  );
}

// ── Calendar ──────────────────────────────────────────────────────────────────
function CalendarView({ events, setEvents, buyers, properties, currentUser }) {
  const today = new Date();
  const [year, setYear] = useState(today.getFullYear());
  const [month, setMonth] = useState(today.getMonth());
  const [selectedDate, setSelectedDate] = useState(today.toISOString().slice(0,10));
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({});
  const upd = (k,v)=>setForm(p=>({...p,[k]:v}));

  const daysInMonth = new Date(year, month+1, 0).getDate();
  const firstDay = new Date(year, month, 1).getDay();
  const monthStr = year+"-"+String(month+1).padStart(2,"0");
  const todayStr = today.toISOString().slice(0,10);

  const dayEvents = (d)=>{
    const ds = year+"-"+String(month+1).padStart(2,"0")+"-"+String(d).padStart(2,"0");
    return events.filter(e=>e.date===ds);
  };

  const selEvents = events.filter(e=>e.date===selectedDate).sort((a,b)=>(a.time||"").localeCompare(b.time||""));

  const openNew = ()=>{
    setForm({id:uid(),title:"",date:selectedDate,time:"10:00",buyerId:"",propertyId:"",notes:"",agent:currentUser});
    setShowForm(true);
  };
  const doSave = ()=>{
    if(!form.title) return;
    const buyer = buyers.find(x=>x.id===form.buyerId);
    const prop = properties.find(x=>x.id===form.propertyId);
    const ev = {...form, buyerName:buyer?buyer.name:"", propName:prop?prop.name:""};
    setEvents(prev=>{ const e=prev.find(x=>x.id===form.id); return e?prev.map(x=>x.id===form.id?ev:x):[...prev,ev]; });
    setShowForm(false);
  };
  const doDel = id=>setEvents(p=>p.filter(x=>x.id!==id));

  const prevMonth = ()=>{ if(month===0){setMonth(11);setYear(y=>y-1);}else setMonth(m=>m-1); };
  const nextMonth = ()=>{ if(month===11){setMonth(0);setYear(y=>y+1);}else setMonth(m=>m+1); };

  const WEEKDAYS = ["日","一","二","三","四","五","六"];

  return (
    <div className="page">
      <div className="page-nav">
        <div className="page-title" style={{marginBottom:0}}>店內行事曆</div>
        <button className="add-btn" onClick={openNew}>＋ 新增</button>
      </div>

      {/* Month nav */}
      <div className="cal-header">
        <button className="cal-nav" onClick={prevMonth}>‹</button>
        <div className="cal-month">{year}年 {month+1}月</div>
        <button className="cal-nav" onClick={nextMonth}>›</button>
      </div>

      {/* Weekday headers */}
      <div className="cal-grid">
        {WEEKDAYS.map(d=><div key={d} className="cal-wday">{d}</div>)}
        {Array(firstDay).fill(null).map((_,i)=><div key={"e"+i}/>)}
        {Array(daysInMonth).fill(null).map((_,i)=>{
          const d = i+1;
          const ds = year+"-"+String(month+1).padStart(2,"0")+"-"+String(d).padStart(2,"0");
          const de = dayEvents(d);
          const isToday = ds===todayStr;
          const isSel = ds===selectedDate;
          return (
            <div key={d} className={"cal-day"+(isToday?" today":"")+(isSel?" selected":"")} onClick={()=>setSelectedDate(ds)}>
              <span>{d}</span>
              {de.length>0 && <div className="cal-dot-row">{de.slice(0,3).map((_,i)=><span key={i} className="cal-dot"/>)}</div>}
            </div>
          );
        })}
      </div>

      {/* Selected day events */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"16px 0 8px"}}>
        <div className="section-hd" style={{margin:0}}>{selectedDate} 的行程（{selEvents.length}）</div>
      </div>

      {showForm && (
        <div className="mini-form">
          <div className="field-wrap"><label className="field-label">標題 *</label><input placeholder="帶看、回訪、說明…" value={form.title||""} onChange={e=>upd("title",e.target.value)}/></div>
          <div className="two-col">
            <div className="field-wrap"><label className="field-label">日期</label><input type="date" value={form.date||""} onChange={e=>upd("date",e.target.value)}/></div>
            <div className="field-wrap"><label className="field-label">時間</label><input type="time" value={form.time||""} onChange={e=>upd("time",e.target.value)}/></div>
          </div>
          <div className="field-wrap"><label className="field-label">買方</label>
            <select value={form.buyerId||""} onChange={e=>upd("buyerId",e.target.value)}>
              <option value="">不指定</option>
              {buyers.map(b=><option key={b.id} value={b.id}>{b.name}</option>)}
            </select>
          </div>
          <div className="field-wrap"><label className="field-label">物件</label>
            <select value={form.propertyId||""} onChange={e=>upd("propertyId",e.target.value)}>
              <option value="">不指定</option>
              {properties.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>
          <div className="field-wrap"><label className="field-label">備註</label><input value={form.notes||""} onChange={e=>upd("notes",e.target.value)}/></div>
          <div style={{display:"flex",gap:8}}>
            <button className="btn-gold" style={{flex:1}} onClick={doSave} disabled={!form.title}>儲存</button>
            <button onClick={()=>setShowForm(false)} style={{flex:1,background:"#1e2535",border:"none",borderRadius:12,color:"#666666",fontFamily:"Noto Sans TC,sans-serif",fontSize:14,cursor:"pointer"}}>取消</button>
          </div>
        </div>
      )}

      {selEvents.length===0 && !showForm && <Empty icon="📅" text="這天沒有行程"/>}
      {selEvents.map(e=>(
        <div className="event-card" key={e.id}>
          <div className="event-time">{e.time||"全天"}</div>
          <div className="event-main">
            <div className="event-title">{e.title}</div>
            {e.duty && <div className="event-sub">🧑‍💼 值班：{e.duty}</div>}
            {e.offDuty && <div className="event-sub">🏖️ 休假：{e.offDuty}</div>}
            {e.buyerName && <div className="event-sub">👤 {e.buyerName}</div>}
            {e.propName && <div className="event-sub">🏠 {e.propName}</div>}
            {!e.duty && !e.offDuty && e.notes && <div className="event-sub">{e.notes}</div>}
            {!e.isSchedule && <div className="event-sub" style={{color:"#4b5563"}}>by {e.agent}</div>}
          </div>
          {!e.isSchedule && <button onClick={()=>doDel(e.id)} style={{background:"none",border:"none",color:"#4b5563",fontSize:16,cursor:"pointer",padding:"0 4px",flexShrink:0}}>✕</button>}
        </div>
      ))}
    </div>
  );
}

// ── AI Match ──────────────────────────────────────────────────────────────────
function AIMatch({ properties, buyers }) {
  const [buyerId, setBuyerId] = useState("");
  const [budget, setBudget] = useState("");
  const [rooms, setRooms] = useState("");
  const [needs, setNeeds] = useState("");
  const [loading, setLoading] = useState(false);
  const [results, setResults] = useState(null);
  const [expandedMatch, setExpandedMatch] = useState(null);

  useEffect(()=>{
    if(!buyerId) return;
    const b = buyers.find(x=>x.id===buyerId);
    if(!b) return;
    setBudget(b.budgetMax||""); setRooms(b.rooms1||""); setNeeds(b.needs||"");
  },[buyerId]);

  const doMatch = ()=>{
    if(!properties.length) return;
    setLoading(true); setResults(null); setExpandedMatch(null);
    const budgetNum = parseFloat(budget)||999999;
    const roomsNum = parseInt(rooms)||0;
    const matches = properties.map(p=>{
      let score = 50;
      const reasons = [];
      if(p.price && p.price <= budgetNum){ score+=20; reasons.push("售價"+p.price+"萬符合預算"); }
      else if(p.price && p.price > budgetNum){ score-=25; reasons.push("超出預算"); }
      if(roomsNum>0 && p.rooms){
        if(Number(p.rooms)===roomsNum){ score+=20; reasons.push(p.rooms+"房符合"); }
        else if(Math.abs(Number(p.rooms)-roomsNum)===1){ score+=8; }
        else { score-=15; }
      }
      if(needs && p.district && (needs.includes(p.district)||p.district.includes(needs))){ score+=15; reasons.push(p.district+"符合地點"); }
      score = Math.max(0,Math.min(100,score));
      const scoreLabel = score>=80?"高度符合":score>=60?"基本符合":score>=40?"部分符合":"不符合";
      return { propertyName:p.name, score, scoreLabel, reason:reasons.length>0?reasons.join("，"):"條件可參考，建議進一步確認" };
    }).filter(m=>m.score>=35).sort((a,b)=>b.score-a.score);
    const summary = matches.length>0?"找到 "+matches.length+" 個符合物件，最推薦「"+matches[0].propertyName+"」":"目前無完全符合物件，建議調整條件";
    setTimeout(()=>{ setResults({matches,summary}); setLoading(false); },200);
  }

  return (
    <div className="page">
      <div className="page-title">AI 物件配對</div>
      {buyers.length>0 && (
        <div className="field-wrap">
          <label className="field-label">從買方帶入需求</label>
          <select value={buyerId} onChange={e=>setBuyerId(e.target.value)}>
            <option value="">— 手動填寫 —</option>
            {buyers.map(b=><option key={b.id} value={b.id}>{b.name}</option>)}
          </select>
        </div>
      )}
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">預算上限（萬）</label><input type="number" placeholder="2000" value={budget} onChange={e=>setBudget(e.target.value)}/></div>
        <div className="field-wrap"><label className="field-label">需求房數</label>
          <select value={rooms} onChange={e=>setRooms(e.target.value)}>
            <option value="">不限</option>
            {["1","2","3","4"].map(v=><option key={v} value={v}>{v}房</option>)}
          </select>
        </div>
      </div>
      <div className="field-wrap"><label className="field-label">需求地點</label><input placeholder="大安區、信義區…" value={needs} onChange={e=>setNeeds(e.target.value)}/></div>
      {properties.length===0 && (
        <div className="warn-box">
          ⚠️ 物件庫是空的！請先到「🏠 物件」頁面匯入或新增物件，才能使用 AI 配對。
        </div>
      )}
      <button className="btn-gold" onClick={doMatch} disabled={loading||!properties.length}>{loading?"分析中…":"🔍 開始配對"}</button>
      {loading && <div className="loading-box"><div className="spinner"/>分析中…</div>}
      {results && !results.error && <>
        <div className="divider"/>
        <div className="section-hd">配對結果</div>
        {results.summary && <div className="ai-hint">💡 {results.summary}</div>}
        {(results.matches||[]).map((m,i)=>{
          const p = properties.find(x=>x.name===m.propertyName);
          return (
            <div className="result-card" key={i} onClick={()=>setExpandedMatch(expandedMatch===i?null:i)} style={{cursor:"pointer"}}>
              <div className={"score-badge "+(m.score>=80?"high":m.score>=60?"mid":"low")}>{m.score}分</div>
              <div className="result-name">{m.propertyName}</div>
              <div className="result-label">{m.scoreLabel}</div>
              <div className="ai-reason">📋 {m.reason}</div>
              {expandedMatch===i && p && (
                <div style={{marginTop:12,paddingTop:12,borderTop:"1px solid rgba(201,169,110,0.2)"}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:7}}>
                    {p.price && <div className="info-item"><div className="info-item-label">開價</div><div className="info-item-val" style={{color:"#e8722a"}}>{p.price}萬</div></div>}
                    {p.area && <div className="info-item"><div className="info-item-label">坪數</div><div className="info-item-val">{p.area}坪</div></div>}
                    {p.layout && <div className="info-item"><div className="info-item-label">格局</div><div className="info-item-val">{p.layout}</div></div>}
                    {p.floor && <div className="info-item"><div className="info-item-label">樓層</div><div className="info-item-val">{p.floor}</div></div>}
                    {p.district && <div className="info-item"><div className="info-item-label">行政區</div><div className="info-item-val">{p.district}</div></div>}
                    {p.type && <div className="info-item"><div className="info-item-label">類型</div><div className="info-item-val">{p.type}</div></div>}
                  </div>
                  {p.notes && <div style={{marginTop:8,fontSize:12,color:"#666666",lineHeight:1.6}}>{p.notes}</div>}
                  {p.features && <div style={{marginTop:6,fontSize:12,color:"#e8722a"}}>{p.features}</div>}
                  {p.url && <a href={p.url} target="_blank" rel="noreferrer" onClick={e=>e.stopPropagation()} style={{display:"block",marginTop:8,fontSize:12,color:"#4ade80"}}>🔗 查看物件連結</a>}
                </div>
              )}
            </div>
          );
        })}
      </>}
      {results && results.error && <div className="warn-box">❌ 配對失敗：{results.msg||"請確認物件庫有資料"}</div>}
    </div>
  );
}





// ── BuyerShowingsSection ──────────────────────────────────────────────────────
function BuyerShowingsSection({ b, showings, setShowings, properties, setBuyers, setSelected, currentUser }) {
  const [view, setView] = useState("list"); // list | form | edit
  const [form, setForm] = useState({});
  const [editId, setEditId] = useState(null);

  // Get all showings for this buyer (new format with items[])
  const bShowings = showings.filter(s=>s.buyerId===b.id||s.buyerName===b.name)
    .sort((a,c)=>(c.date||"").localeCompare(a.date||""));

  const blankForm = () => ({
    id: uid(),
    buyerId: b.id,
    buyerName: b.name,
    date: new Date().toISOString().slice(0,10),
    time: "",
    agent: currentUser,
    items: [{id:uid(), propId:"", propName:"", feedback:"", interest:"普通"}],
    notes:"",
  });

  const updItem = (i,k,v) => setForm(p=>({...p, items:p.items.map((item,idx)=>idx===i?{...item,[k]:v}:item)}));
  const addItem = () => setForm(p=>({...p, items:[...p.items, {id:uid(),propId:"",propName:"",feedback:"",interest:"普通"}]}));
  const delItem = i => setForm(p=>({...p, items:p.items.filter((_,idx)=>idx!==i)}));

  const doSave = () => {
    const rec = {...form, items: form.items.map(item=>{
      const p = properties.find(x=>x.id===item.propId);
      return {...item, propName: p?p.name:item.propName};
    })};
    setShowings(prev=>{
      const e=prev.find(x=>x.id===rec.id);
      return e?prev.map(x=>x.id===rec.id?rec:x):[...prev,rec];
    });
    setView("list");
  };

  const doDel = id => { setShowings(p=>p.filter(x=>x.id!==id)); setView("list"); };

  if(view==="form") return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"16px 0 10px"}}>
        <div className="section-hd" style={{margin:0}}>{editId?"編輯帶看":"新增帶看"}</div>
        <button onClick={()=>{setView("list");setEditId(null);}} style={{background:"none",border:"none",color:"#888888",fontSize:13,cursor:"pointer",fontFamily:"Noto Sans TC,sans-serif"}}>取消</button>
      </div>
      <div className="mini-form">
        <div className="two-col">
          <div className="field-wrap"><label className="field-label">日期</label>
            <input type="date" value={form.date||""} onChange={e=>setForm(p=>({...p,date:e.target.value}))}/>
          </div>
          <div className="field-wrap"><label className="field-label">時間</label>
            <input type="time" value={form.time||""} onChange={e=>setForm(p=>({...p,time:e.target.value}))}/>
          </div>
        </div>
        <div className="field-wrap"><label className="field-label">負責業務</label>
          <input value={form.agent||currentUser} onChange={e=>setForm(p=>({...p,agent:e.target.value}))}/>
        </div>

        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"10px 0 8px"}}>
          <label className="field-label" style={{margin:0}}>帶看物件（可多間）</label>
          <button onClick={addItem} style={{background:"#e8722a",border:"none",borderRadius:20,padding:"4px 12px",color:"#fff",fontSize:12,cursor:"pointer",fontFamily:"Noto Sans TC,sans-serif",fontWeight:600}}>＋ 新增物件</button>
        </div>

        {form.items.map((item,i)=>(
          <div key={item.id||i} style={{background:"#ffffff",border:"1px solid #e0d6ca",borderRadius:10,padding:"10px",marginBottom:8}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
              <span style={{fontSize:12,color:"#e8722a",fontWeight:700}}>物件 {i+1}</span>
              {form.items.length>1&&<button onClick={()=>delItem(i)} style={{background:"none",border:"none",color:"#ff3b30",fontSize:14,cursor:"pointer"}}>✕</button>}
            </div>
            <div className="field-wrap">
              <select value={item.propId||""} onChange={e=>{
                const p=properties.find(x=>x.id===e.target.value);
                updItem(i,"propId",e.target.value);
                if(p) updItem(i,"propName",p.name);
              }} style={{fontSize:13}}>
                <option value="">選擇物件…</option>
                {properties.map(p=><option key={p.id} value={p.id}>{p.name}{p.price?" "+p.price+"萬":""}</option>)}
              </select>
            </div>
            {!item.propId&&<div className="field-wrap">
              <input placeholder="或直接輸入物件名稱" value={item.propName||""} onChange={e=>updItem(i,"propName",e.target.value)} style={{fontSize:13}}/>
            </div>}
            <div className="two-col">
              <div className="field-wrap">
                <select value={item.interest||"普通"} onChange={e=>updItem(i,"interest",e.target.value)} style={{fontSize:13}}>
                  {["高","中","低","普通","婉拒"].map(v=><option key={v}>{v}</option>)}
                </select>
              </div>
            </div>
            <div className="field-wrap">
              <textarea rows={2} placeholder="客戶回饋（可事後補填）" value={item.feedback||""} onChange={e=>updItem(i,"feedback",e.target.value)} style={{fontSize:13}}/>
            </div>
          </div>
        ))}

        <div className="field-wrap"><label className="field-label">整體備註</label>
          <input placeholder="今日帶看整體說明…" value={form.notes||""} onChange={e=>setForm(p=>({...p,notes:e.target.value}))} style={{fontSize:13}}/>
        </div>
        <button className="btn-gold" onClick={doSave}>💾 儲存帶看</button>
        {editId&&<button className="btn-danger" style={{marginTop:8}} onClick={()=>doDel(editId)}>🗑 刪除</button>}
      </div>
    </div>
  );

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"16px 0 8px"}}>
        <div className="section-hd" style={{margin:0}}>帶看記錄（{bShowings.length}次）</div>
        <button onClick={()=>{setForm(blankForm());setEditId(null);setView("form");}} className="add-btn" style={{fontSize:12,padding:"6px 12px"}}>＋ 新增帶看</button>
      </div>
      {bShowings.length===0&&<div style={{fontSize:13,color:"#aaaaaa",padding:"8px 0"}}>尚無帶看記錄</div>}
      {bShowings.map(s=>{
        const items = s.items&&s.items.length>0 ? s.items :
          (s.propName||s.propertyId)?[{propId:s.propertyId||"",propName:s.propName||"",feedback:s.feedback||"",interest:s.interest||"普通"}]:[];
        return (
          <div key={s.id} style={{background:"#f5f0eb",border:"1px solid #e0d6ca",borderRadius:12,padding:"12px",marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
              <div style={{fontSize:12,color:"#888888"}}>{s.date}{s.time?" "+s.time:""} · {s.agent||""}</div>
              <button onClick={()=>{
                const editItems=items.map(x=>({...x,id:x.id||uid()}));
                setForm({...s,items:editItems});setEditId(s.id);setView("form");
              }} style={{background:"none",border:"1px solid #e0d6ca",borderRadius:8,color:"#e8722a",fontSize:11,padding:"3px 8px",cursor:"pointer",fontFamily:"Noto Sans TC,sans-serif"}}>編輯</button>
            </div>
            {items.map((item,i)=>{
              const p=properties.find(x=>x.id===item.propId);
              return (
                <div key={i} style={{background:"#ffffff",borderRadius:8,padding:"8px 10px",marginBottom:i<items.length-1?6:0,borderLeft:"3px solid #e8722a"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:item.feedback?4:0}}>
                    <span style={{fontWeight:600,fontSize:13,color:"#1a1a1a"}}>{item.propName||p?.name||"物件"}</span>
                    <IBadge v={item.interest}/>
                  </div>
                  {item.feedback
                    ?<div style={{fontSize:12,color:"#666666",lineHeight:1.5}}>💬 {item.feedback}</div>
                    :<div style={{fontSize:11,color:"#cccccc"}}>尚未填寫回饋</div>
                  }
                </div>
              );
            })}
            {s.notes&&<div style={{fontSize:12,color:"#888888",marginTop:6}}>📝 {s.notes}</div>}
          </div>
        );
      })}
    </div>
  );
}

// ── ShownProps Section ────────────────────────────────────────────────────────
function ShownPropsSection({ b, properties, setBuyers, setSelected }) {
  const [editIdx, setEditIdx] = useState(null);
  const [editFeedback, setEditFeedback] = useState("");
  const [editInterest, setEditInterest] = useState("普通");

  const openEdit = (i, sp) => {
    setEditIdx(i);
    setEditFeedback(sp.feedback||"");
    setEditInterest(sp.interest||"普通");
  };

  const saveEdit = (i) => {
    const newShownProps = (b.shownProps||[]).map((sp,idx)=>
      idx===i ? {...sp, feedback:editFeedback, interest:editInterest} : sp
    );
    const updated = {...b, shownProps:newShownProps};
    setBuyers(prev=>prev.map(x=>x.id===b.id?updated:x));
    setSelected(updated);
    setEditIdx(null);
  };

  const delShown = (i) => {
    const newShownProps = (b.shownProps||[]).filter((_,idx)=>idx!==i);
    const updated = {...b, shownProps:newShownProps};
    setBuyers(prev=>prev.map(x=>x.id===b.id?updated:x));
    setSelected(updated);
    setEditIdx(null);
  };

  if((b.shownProps||[]).length===0) return <Empty icon="🏠" text="尚未看過物件"/>;

  return (
    <>
      {(b.shownProps||[]).map((sp,i)=>{
        const p = properties.find(x=>x.id===sp.propId);
        const isEdit = editIdx===i;
        return (
          <div className="shown-card" key={i}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div style={{flex:1}}>
                <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{sp.propName||p?.name||"已刪除物件"}</div>
                {p && <div style={{fontSize:11,color:"#888888",marginTop:2}}>{[p.district,p.area?(p.area+"坪"):"",p.price?(p.price+"萬"):""].filter(Boolean).join(" · ")}</div>}
                <div style={{fontSize:11,color:"#aaaaaa",marginTop:2}}>{sp.date}</div>
              </div>
              <div style={{display:"flex",gap:6,alignItems:"center"}}>
                <IBadge v={sp.interest}/>
                <button onClick={()=>isEdit?setEditIdx(null):openEdit(i,sp)} style={{background:"none",border:"1px solid #e0d6ca",borderRadius:8,color:"#e8722a",fontSize:11,padding:"3px 8px",cursor:"pointer",fontFamily:"Noto Sans TC,sans-serif"}}>
                  {isEdit?"取消":"編輯"}
                </button>
              </div>
            </div>
            {!isEdit && sp.feedback && <div className="ai-reason" style={{marginTop:8}}>💬 {sp.feedback}</div>}
            {isEdit && (
              <div style={{marginTop:10}}>
                <div className="field-wrap">
                  <label className="field-label">客戶回饋</label>
                  <textarea rows={2} value={editFeedback} onChange={e=>setEditFeedback(e.target.value)} placeholder="喜歡採光，覺得坪數稍小…"/>
                </div>
                <div className="field-wrap">
                  <label className="field-label">購買意願</label>
                  <select value={editInterest} onChange={e=>setEditInterest(e.target.value)}>
                    {["高","中","低","普通","婉拒"].map(v=><option key={v}>{v}</option>)}
                  </select>
                </div>
                <div style={{display:"flex",gap:8}}>
                  <button className="btn-gold" style={{flex:2}} onClick={()=>saveEdit(i)}>💾 儲存</button>
                  <button onClick={()=>delShown(i)} style={{flex:1,background:"#fff3f3",border:"1px solid #ffcdd2",borderRadius:12,color:"#ff3b30",fontFamily:"Noto Sans TC,sans-serif",fontSize:13,cursor:"pointer"}}>🗑 刪除</button>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </>
  );
}

// ── Timeline Section ──────────────────────────────────────────────────────────
function TimelineSection({ b, setBuyers, setSelected, currentUser }) {
  const [showAdd, setShowAdd] = useState(false);
  const [note, setNote] = useState("");
  const tl = [...(b.timeline||[])].sort((a,c)=>c.time.localeCompare(a.time));

  const addNote = ()=>{
    if(!note.trim()) return;
    const now = new Date().toISOString().slice(0,16).replace("T"," ");
    const entry = {id:uid(), time:now, type:"note", note:note.trim(), agent:currentUser};
    const updated = {...b, timeline:[...(b.timeline||[]), entry]};
    setBuyers(prev=>prev.map(x=>x.id===b.id?updated:x));
    setSelected(updated);
    setNote(""); setShowAdd(false);
  };

  const typeIcon = (type)=>({create:"🆕", note:"📝", showing:"🏠", event:"📅"}[type]||"📌");
  const typeColor = (type)=>({create:"#e8722a", note:"#666666", showing:"#2e7d32", event:"#1976d2"}[type]||"#aaaaaa");

  return (
    <>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"16px 0 8px"}}>
        <div className="section-hd" style={{margin:0}}>互動紀錄（{tl.length}）</div>
        <button onClick={()=>setShowAdd(v=>!v)} className="add-btn" style={{fontSize:12,padding:"6px 12px"}}>＋ 新增備註</button>
      </div>
      {showAdd && (
        <div className="mini-form">
          <div className="field-wrap">
            <label className="field-label">備註內容</label>
            <textarea rows={2} placeholder="回電確認、客戶反饋、跟進狀況…" value={note} onChange={e=>setNote(e.target.value)}/>
          </div>
          <div style={{display:"flex",gap:8}}>
            <button className="btn-gold" style={{flex:1}} onClick={addNote} disabled={!note.trim()}>新增</button>
            <button onClick={()=>{setShowAdd(false);setNote("");}} style={{flex:1,background:"#f5f0eb",border:"1px solid #e0d6ca",borderRadius:12,color:"#666666",fontFamily:"Noto Sans TC,sans-serif",fontSize:14,cursor:"pointer"}}>取消</button>
          </div>
        </div>
      )}
      {tl.length===0 && !showAdd && <div style={{fontSize:12,color:"#aaaaaa",padding:"12px 0"}}>尚無互動記錄</div>}
      <div style={{position:"relative"}}>
        {tl.length>0 && <div style={{position:"absolute",left:16,top:8,bottom:8,width:2,background:"#e0d6ca"}}/>}
        {tl.map((entry,i)=>(
          <div key={entry.id||i} style={{display:"flex",gap:12,marginBottom:12,position:"relative"}}>
            <div style={{width:34,height:34,borderRadius:"50%",background:"#ffffff",border:"2px solid #e0d6ca",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,flexShrink:0,zIndex:1}}>
              {typeIcon(entry.type)}
            </div>
            <div style={{flex:1,paddingTop:4}}>
              <div style={{fontSize:13,color:"#1a1a1a",lineHeight:1.5}}>{entry.note}</div>
              <div style={{fontSize:11,color:"#aaaaaa",marginTop:3}}>
                {entry.time}{entry.agent?" · "+entry.agent:""}
              </div>
            </div>
          </div>
        ))}
      </div>
    </>
  );
}


// ── Showings Page ─────────────────────────────────────────────────────────────
// Data model: one showing = { id, buyerId, buyerName, date, time, agent, items:[{propId,propName,feedback,interest}] }

function ShowingsPage({ showings, setShowings, buyers, properties, currentUser }) {
  const [view, setView] = useState("list");
  const [form, setForm] = useState({});
  const [selected, setSelected] = useState(null);
  const [filterAgent, setFilterAgent] = useState("");

  const blankForm = () => ({
    id: uid(),
    buyerId: "",
    buyerName: "",
    date: new Date().toISOString().slice(0,10),
    time: "",
    agent: currentUser,
    items: [],
    notes: "",
  });

  const openNew = () => { setForm(blankForm()); setSelected(null); setView("form"); };
  const openEdit = s => {
    // Convert old format to new if needed
    const items = s.items && s.items.length > 0 ? s.items :
      (s.propName||s.propertyId) ? [{id:uid(), propId:s.propertyId||"", propName:s.propName||"", feedback:s.feedback||"", interest:s.interest||"普通"}] : [];
    setForm({...s, items}); setSelected(s); setView("form");
  };

  const updForm = (k,v) => setForm(p=>({...p,[k]:v}));

  const addItem = () => setForm(p=>({...p, items:[...p.items, {id:uid(), propId:"", propName:"", feedback:"", interest:"普通"}]}));
  const updItem = (i,k,v) => setForm(p=>({...p, items:p.items.map((item,idx)=>idx===i?{...item,[k]:v}:item)}));
  const delItem = i => setForm(p=>({...p, items:p.items.filter((_,idx)=>idx!==i)}));

  const doSave = () => {
    const buyer = buyers.find(x=>x.id===form.buyerId);
    const rec = {
      ...form,
      buyerName: buyer ? buyer.name : form.buyerName,
      items: form.items.map(item=>{
        const p = properties.find(x=>x.id===item.propId);
        return {...item, propName: p ? p.name : item.propName};
      }),
    };
    setShowings(prev => {
      const e = prev.find(x=>x.id===rec.id);
      return e ? prev.map(x=>x.id===rec.id?rec:x) : [...prev, rec];
    });
    setView("list");
  };

  const doDel = id => { setShowings(p=>p.filter(x=>x.id!==id)); setView("list"); };

  const agents = [...new Set(showings.map(s=>s.agent).filter(Boolean))];
  const sorted = [...showings].sort((a,b)=>(b.date||"").localeCompare(a.date||""));
  const filtered = sorted.filter(s=>!filterAgent||s.agent===filterAgent);

  if(view==="form") return (
    <div className="page">
      <div className="page-nav">
        <button className="back-btn" onClick={()=>setView("list")}>← 返回</button>
        <div className="page-title" style={{marginBottom:0}}>{selected?"編輯帶看":"新增帶看"}</div>
      </div>

      <div className="section-hd">買方與時間</div>
      <div className="field-wrap">
        <label className="field-label">選擇買方</label>
        <select value={form.buyerId||""} onChange={e=>{
          const b=buyers.find(x=>x.id===e.target.value);
          updForm("buyerId",e.target.value);
          if(b) updForm("buyerName",b.name);
        }}>
          <option value="">選擇買方…</option>
          {buyers.map(b=><option key={b.id} value={b.id}>{b.name}</option>)}
        </select>
      </div>
      {!form.buyerId && (
        <div className="field-wrap">
          <label className="field-label">或直接輸入買方姓名</label>
          <input placeholder="王先生" value={form.buyerName||""} onChange={e=>updForm("buyerName",e.target.value)}/>
        </div>
      )}
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">日期</label>
          <input type="date" value={form.date||""} onChange={e=>updForm("date",e.target.value)}/>
        </div>
        <div className="field-wrap"><label className="field-label">時間</label>
          <input type="time" value={form.time||""} onChange={e=>updForm("time",e.target.value)}/>
        </div>
      </div>
      <div className="field-wrap"><label className="field-label">負責業務</label>
        <input value={form.agent||currentUser} onChange={e=>updForm("agent",e.target.value)}/>
      </div>

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"16px 0 8px"}}>
        <div className="section-hd" style={{margin:0}}>帶看物件（{form.items.length}間）</div>
        <button onClick={addItem} className="add-btn" style={{fontSize:12,padding:"6px 12px"}}>＋ 新增物件</button>
      </div>

      {form.items.length===0 && (
        <div style={{background:"#fff8f3",border:"2px dashed #e8d5c0",borderRadius:12,padding:"20px",textAlign:"center",color:"#aaaaaa",fontSize:13,marginBottom:12}}>
          點上方「＋ 新增物件」加入帶看的物件
        </div>
      )}

      {form.items.map((item,i)=>(
        <div key={item.id||i} style={{background:"#f5f0eb",border:"1px solid #e0d6ca",borderRadius:14,padding:"14px",marginBottom:12}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <span style={{fontSize:13,fontWeight:700,color:"#e8722a"}}>物件 {i+1}</span>
            <button onClick={()=>delItem(i)} style={{background:"none",border:"none",color:"#ff3b30",fontSize:16,cursor:"pointer"}}>🗑</button>
          </div>
          <div className="field-wrap">
            <label className="field-label">選擇物件</label>
            <select value={item.propId||""} onChange={e=>{
              const p=properties.find(x=>x.id===e.target.value);
              updItem(i,"propId",e.target.value);
              if(p) updItem(i,"propName",p.name);
            }}>
              <option value="">選擇物件…</option>
              {properties.map(p=><option key={p.id} value={p.id}>{p.name}{p.price?" "+p.price+"萬":""}</option>)}
            </select>
          </div>
          {!item.propId && (
            <div className="field-wrap">
              <label className="field-label">或輸入物件名稱</label>
              <input placeholder="例：新潤鉑麗 3F" value={item.propName||""} onChange={e=>updItem(i,"propName",e.target.value)}/>
            </div>
          )}
          <div className="field-wrap">
            <label className="field-label">客戶回饋（可事後補填）</label>
            <textarea rows={2} placeholder="喜歡採光，覺得坪數稍小…（帶看後再填也可以）" value={item.feedback||""} onChange={e=>updItem(i,"feedback",e.target.value)}/>
          </div>
          <div className="field-wrap">
            <label className="field-label">購買意願</label>
            <select value={item.interest||"普通"} onChange={e=>updItem(i,"interest",e.target.value)}>
              {["高","中","低","普通","婉拒"].map(v=><option key={v}>{v}</option>)}
            </select>
          </div>
        </div>
      ))}

      <div className="field-wrap"><label className="field-label">整體備註</label>
        <textarea rows={2} placeholder="今日帶看整體說明…" value={form.notes||""} onChange={e=>updForm("notes",e.target.value)}/>
      </div>

      <button className="btn-gold" onClick={doSave} disabled={!form.buyerName&&!form.buyerId}>💾 儲存</button>
      {selected && <button className="btn-danger" style={{marginTop:10}} onClick={()=>doDel(selected.id)}>🗑 刪除這筆帶看</button>}
    </div>
  );

  // Detail view - show all items with edit capability
  if(view==="detail") {
    const s = selected;
    return (
      <div className="page">
        <div className="page-nav">
          <button className="back-btn" onClick={()=>setView("list")}>← 返回</button>
          <button className="edit-btn" onClick={()=>openEdit(s)}>編輯</button>
        </div>
        <div className="detail-hero">
          <div className="detail-name">{s.buyerName||"?"}</div>
          <div style={{fontSize:13,color:"#888888"}}>{s.date} {s.time||""} · {s.agent||""}</div>
          {s.notes && <div style={{fontSize:13,color:"#666666",marginTop:4}}>{s.notes}</div>}
        </div>
        <div className="section-hd">帶看物件（{(s.items||[]).length}間）</div>
        {(s.items||[]).map((item,i)=>{
          const p = properties.find(x=>x.id===item.propId);
          return (
            <div key={i} style={{background:"#ffffff",border:"1px solid #e0d6ca",borderRadius:14,padding:"14px",marginBottom:10,borderLeft:"4px solid #e8722a"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
                <div>
                  <div style={{fontWeight:700,fontSize:15,color:"#1a1a1a"}}>{item.propName||p?.name||"物件"}</div>
                  {p&&<div style={{fontSize:11,color:"#888888",marginTop:2}}>{[p.district,p.area?(p.area+"坪"):"",p.price?(p.price+"萬"):""].filter(Boolean).join(" · ")}</div>}
                </div>
                <IBadge v={item.interest}/>
              </div>
              {item.feedback
                ? <div style={{fontSize:13,color:"#444444",background:"#fff8f3",borderRadius:8,padding:"8px 10px",borderLeft:"3px solid #e8722a",lineHeight:1.6}}>💬 {item.feedback}</div>
                : <div style={{fontSize:12,color:"#cccccc",fontStyle:"italic"}}>尚未填寫回饋</div>
              }
            </div>
          );
        })}
      </div>
    );
  }

  return (
    <div className="page">
      <div className="page-nav">
        <div className="page-title" style={{marginBottom:0}}>帶看記錄（{showings.length}）</div>
        <button className="add-btn" onClick={openNew}>＋ 新增</button>
      </div>
      {agents.length>1 && (
        <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
          <button onClick={()=>setFilterAgent("")} style={{background:!filterAgent?"#e8722a":"#f5f0eb",border:"1px solid",borderColor:!filterAgent?"#e8722a":"#e0d6ca",borderRadius:20,padding:"6px 14px",color:!filterAgent?"#fff":"#666666",fontSize:12,cursor:"pointer",fontFamily:"Noto Sans TC,sans-serif"}}>全部</button>
          {agents.map(a=>(
            <button key={a} onClick={()=>setFilterAgent(a===filterAgent?"":a)} style={{background:filterAgent===a?"#e8722a":"#f5f0eb",border:"1px solid",borderColor:filterAgent===a?"#e8722a":"#e0d6ca",borderRadius:20,padding:"6px 14px",color:filterAgent===a?"#fff":"#666666",fontSize:12,cursor:"pointer",fontFamily:"Noto Sans TC,sans-serif"}}>{a}</button>
          ))}
        </div>
      )}
      {filtered.length===0
        ? <Empty icon="🏃" text="尚無帶看記錄"/>
        : filtered.map(s=>{
          // Support both old format (propertyId) and new format (items[])
          const items = s.items && s.items.length > 0 ? s.items : 
            (s.propName||s.propertyId) ? [{propId:s.propertyId||"", propName:s.propName||"", feedback:s.feedback||"", interest:s.interest||"普通"}] : [];
          return (
            <div key={s.id} onClick={()=>{setSelected({...s,items});setView("detail");}} style={{background:"#ffffff",border:"1px solid #e0d6ca",borderRadius:14,padding:"14px",marginBottom:10,cursor:"pointer",borderLeft:"4px solid #e8722a"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                <div>
                  <span style={{fontWeight:700,fontSize:15,color:"#1a1a1a"}}>{s.buyerName||"?"}</span>
                  <span style={{fontSize:12,color:"#888888",marginLeft:8}}>{s.date}{s.time?" "+s.time:""}</span>
                </div>
                <span style={{fontSize:11,color:"#e8722a",fontWeight:600}}>{s.agent||""}</span>
              </div>
              <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:s.notes?6:0}}>
                {items.map((item,i)=>(
                  <div key={i} style={{background:"#f5f0eb",border:"1px solid #e8d5c0",borderRadius:20,padding:"4px 10px",fontSize:12,display:"flex",alignItems:"center",gap:5}}>
                    <span style={{color:"#1a1a1a"}}>{item.propName||"物件"}</span>
                    <IBadge v={item.interest}/>
                  </div>
                ))}
                {items.length===0&&<span style={{fontSize:12,color:"#cccccc"}}>尚未加入物件</span>}
              </div>
              {s.notes&&<div style={{fontSize:12,color:"#888888",marginTop:4}}>{s.notes}</div>}
            </div>
          );
        })
      }
    </div>
  );
}

// ── Shared Buyers (店內買方) ───────────────────────────────────────────────────
function SharedBuyers({ buyers, setBuyers, currentUser, myClients }) {
  const [view, setView] = useState("list");
  const [selected, setSelected] = useState(null);
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("");

  const filtered = buyers.filter(b=>{
    if(filterStatus && b.status!==filterStatus) return false;
    if(search && !b.name.includes(search) && !(b.needs||"").includes(search) && !(b.urgentNeeds||"").includes(search)) return false;
    return true;
  }).sort((a,b)=>(b.pinned?1:0)-(a.pinned?1:0));

  if(view==="detail") {
    const b = selected;
    // Only show phone if current user is the agent
    const isOwner = b.agent === currentUser;
    const mc = isOwner ? myClients.find(x=>x.buyerId===b.id) : null;
    return (
      <div className="page">
        <div className="page-nav">
          <button className="back-btn" onClick={()=>setView("list")}>← 返回</button>
        </div>
        <div className="detail-hero">
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div className="detail-name">{b.name}</div>
            <BSBadge v={b.status}/>
          </div>
          <div style={{fontSize:12,color:"#888888"}}>負責：{b.agent} · {b.createdAt}</div>
          {isOwner && mc && mc.phone && <a href={"tel:"+mc.phone} className="phone-link">📞 {mc.phone}</a>}
          {isOwner && mc && mc.line && <div style={{fontSize:13,color:"#e8722a"}}>💬 Line: {mc.line}</div>}
          {!isOwner && <div className="info-box" style={{background:"#1a0808",borderColor:"#3a1010",color:"#f87171",fontSize:12,marginTop:4}}>🔒 電話由 {b.agent} 保管，請直接聯絡</div>}
        </div>
        <div className="info-grid">
          {b.level && <InfoItem label="買方級別" val={b.level}/>}
          {b.downPayment && <InfoItem label="自備款" val={b.downPayment+"萬"}/>}
          {(b.budgetMin||b.budgetMax) && <InfoItem label="預算" val={(b.budgetMin||"不限")+"~"+(b.budgetMax||"不限")+"萬"}/>}
          {(b.rooms1||b.rooms2) && <InfoItem label="房數" val={(b.rooms1||"不限")+"~"+(b.rooms2||"不限")+"房"}/>}
          {b.needs && <InfoItem label="需求地點" val={b.needs}/>}
          {b.purpose && <InfoItem label="用途" val={b.purpose}/>}
          {b.motivation && <InfoItem label="動機" val={b.motivation}/>}
        </div>
        {(b.propTypes||[]).length>0 && (
          <div className="info-box">
            <div className="info-box-label">需求形態</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:6,marginTop:4}}>
              {(b.propTypes||[]).map(t=><span key={t} className="chip active" style={{cursor:"default"}}>{t}</span>)}
            </div>
          </div>
        )}
        {b.notes && <div className="info-box"><div className="info-box-label">備註</div>{b.notes}</div>}
        {(b.shownProps||[]).length>0 && <>
          <div className="section-hd">看過物件（{(b.shownProps||[]).length}）</div>
          {(b.shownProps||[]).map((sp,i)=>(
            <div className="shown-card" key={i}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{sp.propName}</div>
                  <div style={{fontSize:11,color:"#666666",marginTop:2}}>{sp.date}</div>
                </div>
                <IBadge v={sp.interest}/>
              </div>
              {sp.feedback && <div className="ai-reason" style={{marginTop:8}}>💬 {sp.feedback}</div>}
            </div>
          ))}
        </>}
      </div>
    );
  }

  return (
    <div className="page">
      <div className="page-nav">
        <div className="page-title" style={{marginBottom:0}}>店內買方（{buyers.length}）</div>
      </div>
      <div className="info-box" style={{marginBottom:12,fontSize:12,color:"#666666"}}>
        📋 到「買方」頁新增後，電話由各業務保管
      </div>
      <div style={{display:"flex",gap:8,marginBottom:12}}>
        <input placeholder="搜尋姓名或需求…" value={search} onChange={e=>setSearch(e.target.value)} style={{flex:1,padding:"9px 12px",fontSize:13}}/>
        <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} style={{width:90,padding:"9px 8px",fontSize:12}}>
          <option value="">全部</option>
          {BUYER_STATUSES.map(s=><option key={s}>{s}</option>)}
        </select>
      </div>
      {filtered.length===0
        ? <Empty icon="🏪" text="尚無買方資料"/>
        : filtered.map(b=>(
          <div key={b.id} onClick={()=>{setSelected(b);setView("detail");}}
            style={{
              background: b.pinned ? "#fff8f3" : "#ffffff",
              border: b.pinned ? "2px solid #e8722a" : "1px solid #e0d6ca",
              borderRadius:14, padding:"14px", marginBottom:10, cursor:"pointer"
            }}>
            {/* Pinned badge */}
            {b.pinned && (
              <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:8}}>
                <span style={{background:"#e8722a",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>📌 急需配對</span>
                <span style={{fontSize:11,color:"#888888"}}>負責：{b.agent}</span>
              </div>
            )}
            {/* Name + Status */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}>
              <div>
                {b.level&&<span style={{background:"#e8722a",color:"#fff",borderRadius:4,padding:"1px 6px",fontSize:11,fontWeight:700,marginRight:6}}>{b.level}</span>}
                <span style={{fontWeight:700,fontSize:15,color:"#1a1a1a"}}>{b.name}</span>
              </div>
              <BSBadge v={b.status}/>
            </div>
            {/* urgentNeeds - prominent display */}
            {b.urgentNeeds && (
              <div style={{background:b.pinned?"#fff0eb":"#f5f0eb",border:"1px solid",borderColor:b.pinned?"#e8722a":"#e8d5c0",borderRadius:8,padding:"8px 12px",margin:"6px 0",fontSize:13,fontWeight:600,color:"#1a1a1a",lineHeight:1.5}}>
                🎯 {b.urgentNeeds}
              </div>
            )}
            {/* Sub info */}
            <div style={{fontSize:12,color:"#888888",marginTop:4}}>
              {[b.budgetMin&&b.budgetMax?b.budgetMin+"~"+b.budgetMax+"萬":b.budgetMax?b.budgetMax+"萬以內":"", b.rooms1?b.rooms1+(b.rooms2?"-"+b.rooms2:"")+"房":"", b.needs].filter(Boolean).join(" · ")}
            </div>
            {(b.propTypes||[]).length>0&&<div style={{fontSize:11,color:"#aaaaaa",marginTop:2}}>{(b.propTypes||[]).join("、")}</div>}
            {!b.pinned&&<div style={{fontSize:11,color:"#aaaaaa",marginTop:2}}>負責：{b.agent} · 看過 {(b.shownProps||[]).length} 間</div>}
          </div>
        ))
      }
    </div>
  );
}

// ── Loan Calc ─────────────────────────────────────────────────────────────────
function LoanCalc() {
  const [price,setPrice]=useState("1500");
  const [ratio,setRatio]=useState("70");
  const [rate,setRate]=useState("2.1");
  const [years,setYears]=useState("30");
  const [result,setResult]=useState(null);
  const calc=()=>{
    const p=Number(price)*10000*Number(ratio)/100;
    const mr=Number(rate)/100/12;
    const n=Number(years)*12;
    const m=p*mr*Math.pow(1+mr,n)/(Math.pow(1+mr,n)-1);
    const t=m*n;
    setResult({monthly:Math.round(m),total:Math.round(t),interest:Math.round(t-p),down:Math.round(Number(price)*10000-p),principal:Math.round(p)});
  };
  const fmtW=n=>n?(n/10000).toFixed(1)+"萬":"";
  return (
    <div className="page">
      <div className="page-title">貸款試算</div>
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">售價（萬）</label><input type="number" value={price} onChange={e=>setPrice(e.target.value)}/></div>
        <div className="field-wrap"><label className="field-label">貸款成數（%）</label><input type="number" value={ratio} onChange={e=>setRatio(e.target.value)}/></div>
      </div>
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">年利率（%）</label><input type="number" step="0.01" value={rate} onChange={e=>setRate(e.target.value)}/></div>
        <div className="field-wrap"><label className="field-label">貸款年期</label>
          <select value={years} onChange={e=>setYears(e.target.value)}>
            {["10","15","20","25","30","35","40"].map(v=><option key={v} value={v}>{v}年</option>)}
          </select>
        </div>
      </div>
      <button className="btn-gold" onClick={calc}>💰 計算</button>
      {result && (
        <div className="loan-result">
          <div className="loan-main"><div className="loan-main-label">每月還款</div><div className="loan-main-val">NT$ {result.monthly.toLocaleString()}</div></div>
          <div className="loan-grid">
            {[["頭期款",fmtW(result.down)],["貸款金額",fmtW(result.principal)],["利息總額",fmtW(result.interest)],["還款總額",fmtW(result.total)]].map(([l,v])=>(
              <div key={l} className="loan-item"><div className="loan-item-label">{l}</div><div className="loan-item-val">{v}</div></div>
            ))}
          </div>
          <div className="loan-note">※ 以上為試算參考，實際以銀行核定為準</div>
        </div>
      )}
    </div>
  );
}

// ── Properties ────────────────────────────────────────────────────────────────
function Properties({ properties, setProperties, showings, buyers }) {
  const [view, setView] = useState("list");
  const [form, setForm] = useState({});
  const [selected, setSelected] = useState(null);
  const [showImport, setShowImport] = useState(false);
  const [pasteText, setPasteText] = useState("");
  const [importMsg, setImportMsg] = useState("");
  const [importStats, setImportStats] = useState(null);
  const [importCategory, setImportCategory] = useState("售屋");
  const [filterCategory, setFilterCategory] = useState("全部");
  useEffect(()=>{ if(filterCategory!=="全部") setImportCategory(filterCategory); },[filterCategory]);
  const upd=(k,v)=>setForm(p=>({...p,[k]:v}));
  const openDetail=p=>{ setSelected(p); setView("detail"); };
  const openNew=(cat="售屋")=>{ setForm({id:uid(),name:"",propCategory:cat,price:"",area:"",rooms:"",district:"",floor:"",type:"",layout:"",parking:"",features:"",notes:"",_fromSheet:false}); setSelected(null); setView("form"); };
  const openEdit=p=>{ setForm({...p, propCategory:p.propCategory||"售屋"}); setSelected(p); setView("form"); };
  const doSave=()=>{
    const rec={...form,price:Number(form.price),area:Number(form.area),rooms:Number(form.rooms)};
    setProperties(prev=>{ const e=prev.find(x=>x.id===rec.id); return e?prev.map(x=>x.id===rec.id?rec:x):[...prev,rec]; });
    setView("list");
  };
  const doDel=id=>{ setProperties(p=>p.filter(x=>x.id!==id)); setView("list"); };
  const doImport=()=>{
    if(!pasteText.trim()){ setImportMsg("⚠️ 請先貼上資料"); return; }
    setImportMsg(""); setImportStats(null);
    const rows=parseSheetText(pasteText);
    if(rows.length===0){ setImportMsg("⚠️ 沒有讀到資料"); return; }
    const imported=rows.map(row=>rowToProperty(row,importCategory)).filter(p=>p.name&&p.name!=="未命名");
    if(imported.length===0){ setImportMsg("⚠️ 無法辨識欄位"); return; }
    let added=0, updated=0, unchanged=0;
    // Start with ALL existing properties untouched
    const result = [...properties];
    imported.forEach(np=>{
      // Only match within same category and name
      const exIdx = result.findIndex(e=>e.name===np.name&&(e.propCategory||"售屋")===importCategory);
      if(exIdx===-1){
        // New property - just add it
        result.push(np);
        added++;
      } else {
        const ex = result[exIdx];
        if(ex.price!==np.price||ex.area!==np.area||ex.floor!==np.floor||ex.rent!==np.rent){
          result[exIdx]={...np,id:ex.id}; // keep same id, update data
          updated++;
        } else {
          unchanged++;
        }
      }
    });
    setProperties(result);
    setImportStats({added,updated,unchanged}); setImportMsg("ok"); setPasteText("");
  };

  if(view==="form") return (
    <div className="page">
      <div className="page-nav">
        <button className="back-btn" onClick={()=>setView("list")}>← 返回</button>
        <div className="page-title" style={{marginBottom:0}}>{selected?"編輯物件":"新增物件"}</div>
      </div>

      {/* 物件類型 - 新增時可選擇，編輯時只顯示目前類型 */}
      <div className="field-wrap">
        <label className="field-label">物件類型</label>
        {selected ? (
          <div style={{display:"inline-flex",alignItems:"center",gap:8,background:(form.propCategory||"售屋")==="出租"?"#eff6ff":(form.propCategory||"售屋")==="口約"?"#f5f3ff":"#fff8f3",border:"2px solid",borderColor:(form.propCategory||"售屋")==="出租"?"#3b82f6":(form.propCategory||"售屋")==="口約"?"#8b5cf6":"#e8722a",borderRadius:10,padding:"10px 16px",fontSize:14,fontWeight:700,color:(form.propCategory||"售屋")==="出租"?"#3b82f6":(form.propCategory||"售屋")==="口約"?"#8b5cf6":"#e8722a"}}>
            {(form.propCategory||"售屋")==="出租"?"🔑 出租物件":(form.propCategory||"售屋")==="口約"?"🤝 口約物件":"🏠 售屋物件"}
          </div>
        ) : (
          <div style={{display:"flex",gap:8}}>
            {["售屋","出租","口約"].map(t=>(
              <button key={t} onClick={()=>upd("propCategory",t)} style={{flex:1,padding:"10px",background:(form.propCategory||"售屋")===t?"#e8722a":"#f5f0eb",border:"1px solid",borderColor:(form.propCategory||"售屋")===t?"#e8722a":"#e0d6ca",borderRadius:10,color:(form.propCategory||"售屋")===t?"#fff":"#666666",fontFamily:"Noto Sans TC,sans-serif",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                {t==="售屋"?"🏠":t==="出租"?"🔑":"🤝"} {t}
              </button>
            ))}
          </div>
        )}
      </div>

      {/* 共用基本欄位 */}
      <div className="field-wrap"><label className="field-label">物件名稱（社區）*</label><input value={form.name||""} onChange={e=>upd("name",e.target.value)}/></div>
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">區域</label><input placeholder="A7、信義區…" value={form.district||""} onChange={e=>upd("district",e.target.value)}/></div>
        <div className="field-wrap"><label className="field-label">型態</label>
          <select value={form.type||""} onChange={e=>upd("type",e.target.value)}>
            <option value="">選擇</option>
            {["電梯大樓","公寓","透天厝","套房","店面","土地","廠房"].map(v=><option key={v}>{v}</option>)}
          </select>
        </div>
      </div>
      <div className="two-col">
        <div className="field-wrap"><label className="field-label">格局（房/廳/衛/陽）</label><input placeholder="2/2/1/1" value={form.layout||""} onChange={e=>upd("layout",e.target.value)}/></div>
        <div className="field-wrap"><label className="field-label">車位</label>
          <select value={form.parking||""} onChange={e=>upd("parking",e.target.value)}>
            <option value="">無</option>
            {["有","平面","機械","坡道"].map(v=><option key={v}>{v}</option>)}
          </select>
        </div>
      </div>
      <div className="field-wrap"><label className="field-label">樓層</label><input placeholder="例：3/15、頂樓" value={form.floor||""} onChange={e=>upd("floor",e.target.value)}/></div>

      {/* ── 售屋專用 ── */}
      {(form.propCategory||"售屋")==="售屋" && <>
        <div className="section-hd">售屋資訊</div>
        <div className="field-wrap"><label className="field-label">開價（萬）</label><input type="number" value={form.price||""} onChange={e=>upd("price",e.target.value)}/></div>
        <div className="two-col">
          <div className="field-wrap"><label className="field-label">主建物（坪）</label><input type="number" placeholder="0.00" value={form.areaMain||""} onChange={e=>upd("areaMain",e.target.value)}/></div>
          <div className="field-wrap"><label className="field-label">附屬建物（坪）</label><input type="number" placeholder="0.00" value={form.areaSub||""} onChange={e=>upd("areaSub",e.target.value)}/></div>
        </div>
        <div className="two-col">
          <div className="field-wrap"><label className="field-label">公設（坪）</label><input type="number" placeholder="0.00" value={form.areaCommon||""} onChange={e=>upd("areaCommon",e.target.value)}/></div>
          <div className="field-wrap"><label className="field-label">車位（坪）</label><input type="number" placeholder="0.00" value={form.areaParking||""} onChange={e=>upd("areaParking",e.target.value)}/></div>
        </div>
        <div className="two-col">
          <div className="field-wrap"><label className="field-label">帶看方式</label><input placeholder="自由帶看、配合帶…" value={form.showingType||""} onChange={e=>upd("showingType",e.target.value)}/></div>
          <div className="field-wrap"><label className="field-label">開發業務</label><input placeholder="采萱" value={form.agent2||""} onChange={e=>upd("agent2",e.target.value)}/></div>
        </div>
      </>}

      {/* ── 出租專用 ── */}
      {form.propCategory==="出租" && <>
        <div className="section-hd">出租資訊</div>
        <div className="two-col">
          <div className="field-wrap"><label className="field-label">租金（元/月）</label><input type="number" placeholder="25000" value={form.rent||""} onChange={e=>upd("rent",e.target.value)}/></div>
          <div className="field-wrap"><label className="field-label">服務費（元）</label><input type="number" placeholder="3000" value={form.serviceFee||""} onChange={e=>upd("serviceFee",e.target.value)}/></div>
        </div>
        <div className="field-wrap"><label className="field-label">地址</label><input placeholder="完整地址" value={form.address||""} onChange={e=>upd("address",e.target.value)}/></div>
        <div className="two-col">
          <div className="field-wrap"><label className="field-label">帶看方式</label><input placeholder="自由帶看、配合帶…" value={form.showingType||""} onChange={e=>upd("showingType",e.target.value)}/></div>
          <div className="field-wrap"><label className="field-label">開發業務</label><input placeholder="采萱" value={form.agent2||""} onChange={e=>upd("agent2",e.target.value)}/></div>
        </div>
      </>}

      {/* ── 口約專用 ── */}
      {form.propCategory==="口約" && <>
        <div className="section-hd">口約資訊</div>
        <div className="two-col">
          <div className="field-wrap"><label className="field-label">開價（萬）</label><input type="number" value={form.price||""} onChange={e=>upd("price",e.target.value)}/></div>
          <div className="field-wrap"><label className="field-label">權狀坪數</label><input type="number" placeholder="0.00" value={form.area||""} onChange={e=>upd("area",e.target.value)}/></div>
        </div>
        <div className="two-col">
          <div className="field-wrap"><label className="field-label">扣車坪數</label><input type="number" placeholder="0.00" value={form.netArea||""} onChange={e=>upd("netArea",e.target.value)}/></div>
          <div className="field-wrap"><label className="field-label">單價/坪（扣車）</label><input type="number" placeholder="0.00" value={form.unitPrice||""} onChange={e=>upd("unitPrice",e.target.value)}/></div>
        </div>
        <div className="two-col">
          <div className="field-wrap"><label className="field-label">帶看方式</label><input placeholder="自由帶看、配合帶…" value={form.showingType||""} onChange={e=>upd("showingType",e.target.value)}/></div>
          <div className="field-wrap"><label className="field-label">開發業務</label><input placeholder="采萱" value={form.agent2||""} onChange={e=>upd("agent2",e.target.value)}/></div>
        </div>
      </>}

      {/* 共用下方欄位 */}
      <div className="field-wrap"><label className="field-label">特殊事項</label><input placeholder="近捷運、有管理員…" value={form.features||""} onChange={e=>upd("features",e.target.value)}/></div>
      <div className="field-wrap"><label className="field-label">物件連結</label><input placeholder="https://..." value={form.url||""} onChange={e=>upd("url",e.target.value)}/></div>
      <div className="field-wrap">
        <label className="field-label">上傳物調資料（PDF / 圖片）</label>
        <label style={{display:"flex",alignItems:"center",gap:10,background:"#fff8f3",border:"2px dashed #e8722a",borderRadius:12,padding:"14px 16px",cursor:"pointer",color:"#e8722a",fontSize:14,fontWeight:600}}>
          📎 點擊選擇檔案（PDF / PNG / JPG）
          <input type="file" accept="application/pdf,image/*" multiple onChange={e=>{
            const files=Array.from(e.target.files);
            files.forEach(file=>{
              const reader=new FileReader();
              reader.onload=ev=>{ const isPdf=file.type==="application/pdf"; upd("docs",[...(form.docs||[]),{id:uid(),name:file.name,data:ev.target.result,isPdf}]); };
              reader.readAsDataURL(file);
            });
          }} style={{display:"none"}}/>
        </label>
        {(form.docs||[]).length>0&&(
          <div style={{display:"flex",flexDirection:"column",gap:8,marginTop:10}}>
            {(form.docs||[]).map((doc,i)=>(
              <div key={doc.id||i} style={{background:"#f5f0eb",border:"1px solid #e0d6ca",borderRadius:12,overflow:"hidden"}}>
                {!doc.isPdf&&<img src={doc.data} alt={doc.name} style={{width:"100%",maxHeight:200,objectFit:"contain",background:"#ffffff",display:"block"}}/>}
                <div style={{display:"flex",alignItems:"center",gap:8,padding:"10px 12px"}}>
                  <span style={{fontSize:16}}>{doc.isPdf?"📄":"🖼️"}</span>
                  <span style={{flex:1,fontSize:12,color:"#444444",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{doc.name}</span>
                  <button onClick={()=>upd("docs",(form.docs||[]).filter((_,j)=>j!==i))} style={{background:"none",border:"none",color:"#ff3b30",cursor:"pointer",fontSize:18}}>✕</button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
      <div className="field-wrap"><label className="field-label">備註</label><textarea rows={2} value={form.notes||""} onChange={e=>upd("notes",e.target.value)}/></div>
      <button className="btn-gold" onClick={doSave} disabled={!form.name}>💾 儲存</button>
      {selected&&<button className="btn-danger" style={{marginTop:10}} onClick={()=>doDel(selected.id)}>🗑 刪除</button>}
    </div>
  );


  if(view==="detail") {
    const p = selected;
    const propShowings = (()=>{
      const norm = s => (s||"").trim().replace(/\s+/g," ");
      // From showings array
      const fromShowings = (showings||[]).filter(s=>
        s.propertyId===p.id ||
        (norm(s.propName) && norm(s.propName)===norm(p.name))
      );
      // From buyers shownProps
      const fromBuyers = [];
      (buyers||[]).forEach(b=>{
        (b.shownProps||[]).forEach((sp,i)=>{
          const nameMatch = norm(sp.propName) && norm(sp.propName)===norm(p.name);
          const idMatch = sp.propId && sp.propId===p.id;
          if(nameMatch || idMatch) {
            // Only skip if EXACT same buyer+date+feedback already exists
            const alreadyIn = fromShowings.find(s=>
              s.buyerId===b.id &&
              s.date===sp.date &&
              (s.feedback||"")===(sp.feedback||"")
            );
            if(!alreadyIn) {
              fromBuyers.push({
                id: "sp_"+b.id+"_"+sp.date+"_"+i,
                buyerId: b.id,
                buyerName: b.name||"",
                propertyId: p.id,
                propName: p.name,
                feedback: sp.feedback||"",
                interest: sp.interest||"普通",
                date: sp.date||"",
                agent: b.agent||"",
                time: "",
              });
            }
          }
        });
      });
      const all = [...fromShowings, ...fromBuyers];
      all.sort((a,b)=>(b.date||"").localeCompare(a.date||""));
      return all;
    })();
    return (
      <div className="page">
        <div className="page-nav">
          <button className="back-btn" onClick={()=>setView("list")}>← 返回</button>
          <button className="edit-btn" onClick={()=>openEdit(p)}>編輯</button>
        </div>

        <div className="detail-hero">
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div className="detail-name" style={{fontSize:20}}>{p.name}</div>
            <div className="price-tag" style={{fontSize:20}}>{p.price?p.price+"萬":""}</div>
          </div>
          <div style={{fontSize:13,color:"#888888"}}>{[p.district,p.layout||(p.rooms?(p.rooms+"房"):""),p.type,p.floor].filter(Boolean).join(" · ")}</div>
        </div>

        {/* Category badge */}
        {p.propCategory&&p.propCategory!=="售屋"&&(
          <div style={{display:"inline-flex",alignItems:"center",gap:6,background:p.propCategory==="出租"?"#eff6ff":"#f5f3ff",border:"1px solid",borderColor:p.propCategory==="出租"?"#bfdbfe":"#ddd6fe",borderRadius:20,padding:"4px 12px",marginBottom:12,fontSize:13,fontWeight:700,color:p.propCategory==="出租"?"#3b82f6":"#8b5cf6"}}>
            {p.propCategory==="出租"?"🔑 出租物件":"🤝 口約物件"}
          </div>
        )}

        {/* Rental specific */}
        {p.propCategory==="出租"&&(
          <div className="info-box" style={{marginBottom:12}}>
            <div className="info-box-label">出租資訊</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginTop:6}}>
              {p.rent&&<div style={{fontSize:14,fontWeight:700,color:"#e8722a"}}>{Number(p.rent).toLocaleString()} 元/月</div>}
              {p.serviceFee&&<div style={{fontSize:13}}>服務費：{Number(p.serviceFee).toLocaleString()} 元</div>}
              {p.address&&<div style={{fontSize:13,gridColumn:"1/-1"}}>📍 {p.address}</div>}
              {p.showingType&&<div style={{fontSize:13}}>帶看：{p.showingType}</div>}
              {p.agent2&&<div style={{fontSize:13}}>開發：{p.agent2}</div>}
            </div>
          </div>
        )}

        {/* Verbal specific */}
        {p.propCategory==="口約"&&(
          <div className="info-box" style={{marginBottom:12}}>
            <div className="info-box-label">口約資訊</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginTop:6}}>
              {p.price&&<div style={{fontSize:14,fontWeight:700,color:"#e8722a"}}>{p.price} 萬</div>}
              {p.area&&<div style={{fontSize:13}}>權狀：{p.area} 坪</div>}
              {p.netArea&&<div style={{fontSize:13}}>扣車：{p.netArea} 坪</div>}
              {p.unitPrice&&<div style={{fontSize:13}}>單價：{p.unitPrice} 萬/坪</div>}
              {p.showingType&&<div style={{fontSize:13}}>帶看：{p.showingType}</div>}
              {p.agent2&&<div style={{fontSize:13}}>開發：{p.agent2}</div>}
            </div>
          </div>
        )}

        {/* Area breakdown */}
        {(p.areaMain||p.areaSub||p.areaCommon||p.areaParking) && (
          <div className="info-box" style={{marginBottom:12}}>
            <div className="info-box-label">坪數明細</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginTop:6}}>
              {p.areaMain&&<div style={{fontSize:13}}><span style={{color:"#888888"}}>主建物 </span><span style={{fontWeight:700}}>{p.areaMain}坪</span></div>}
              {p.areaSub&&<div style={{fontSize:13}}><span style={{color:"#888888"}}>附屬建物 </span><span style={{fontWeight:700}}>{p.areaSub}坪</span></div>}
              {p.areaCommon&&<div style={{fontSize:13}}><span style={{color:"#888888"}}>公設 </span><span style={{fontWeight:700}}>{p.areaCommon}坪</span></div>}
              {p.areaParking&&<div style={{fontSize:13}}><span style={{color:"#888888"}}>車位 </span><span style={{fontWeight:700}}>{p.areaParking}坪</span></div>}
            </div>
            {(p.areaMain||p.areaSub||p.areaCommon) && (
              <div style={{marginTop:6,paddingTop:6,borderTop:"1px solid #e0d6ca",fontSize:13}}>
                <span style={{color:"#888888"}}>權狀總坪 </span>
                <span style={{fontWeight:700,color:"#e8722a"}}>{((Number(p.areaMain)||0)+(Number(p.areaSub)||0)+(Number(p.areaCommon)||0)+(Number(p.areaParking)||0)).toFixed(2)}坪</span>
              </div>
            )}
          </div>
        )}
        {!p.areaMain && p.area && <div className="info-box" style={{marginBottom:12}}><div className="info-box-label">坪數</div>{p.area}坪</div>}

        <div className="info-grid">
          {p.district&&<InfoItem label="行政區" val={p.district}/>}
          {p.floor&&<InfoItem label="樓層" val={p.floor}/>}
          {p.type&&<InfoItem label="類型" val={p.type}/>}
          {p.layout&&<InfoItem label="格局" val={p.layout}/>}
          {p.parking&&p.parking!=="無"&&<InfoItem label="車位" val={p.parking}/>}
          {p.showingType&&<InfoItem label="帶看方式" val={p.showingType}/>}
          {p.agent2&&<InfoItem label="開發業務" val={p.agent2}/>}
        </div>
        {p.features&&<div className="info-box"><div className="info-box-label">特殊事項</div>{p.features}</div>}
        {p.notes&&<div className="info-box"><div className="info-box-label">備註</div>{p.notes}</div>}

        {/* Links */}
        {p.url&&<a href={p.url} target="_blank" rel="noreferrer" style={{display:"flex",alignItems:"center",gap:6,background:"#fff8f3",border:"1px solid #e8d5c0",borderRadius:12,padding:"12px 14px",color:"#e8722a",fontSize:14,textDecoration:"none",fontWeight:600,marginBottom:10}}>🔗 查看物件連結</a>}

        {/* Docs: PDF + Images */}
        {(p.docs||[]).length>0&&(
          <>
            <div className="section-hd">物調資料（{p.docs.length}份）</div>
            {p.docs.map((doc,i)=>(
              <div key={doc.id||i} style={{marginBottom:12}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                  <span style={{fontSize:16}}>{doc.isPdf?"📄":"🖼️"}</span>
                  <span style={{fontSize:12,color:"#666666",fontWeight:600}}>{doc.name}</span>
                </div>
                {doc.isPdf ? (
                  <a href={doc.data} download={doc.name} style={{display:"inline-flex",alignItems:"center",gap:6,background:"#fff3f3",border:"1px solid #ffcdd2",borderRadius:10,padding:"10px 14px",color:"#ff3b30",fontSize:13,textDecoration:"none",fontWeight:600}}>
                    📥 下載 PDF
                  </a>
                ) : (
                  <img src={doc.data} alt={doc.name} style={{width:"100%",borderRadius:12,border:"1px solid #e0d6ca",maxHeight:500,objectFit:"contain",background:"#f5f0eb",cursor:"pointer"}} onClick={()=>window.open(doc.data,"_blank")}/>
                )}
              </div>
            ))}
          </>
        )}
        {/* Fallback for old images field */}
        {(p.images||[]).length>0&&!(p.docs||[]).length&&(
          <>
            <div className="section-hd">物調圖片</div>
            {p.images.map((img,i)=>(
              <img key={i} src={img.data} alt={img.name} style={{width:"100%",borderRadius:12,border:"1px solid #e0d6ca",marginBottom:10,objectFit:"contain"}} onClick={()=>window.open(img.data,"_blank")}/>
            ))}
          </>
        )}

        {/* Showings Stats + List */}
        <div className="section-hd">帶看紀錄（{propShowings.length}次）</div>
        {propShowings.length===0
          ? <div style={{fontSize:13,color:"#aaaaaa",padding:"12px 0"}}>尚無帶看紀錄</div>
          : <>
            {/* Stats */}
            {(()=>{
              const agentMap = {};
              const interestMap = {};
              propShowings.forEach(s=>{
                const a = s.agent||"未知";
                agentMap[a] = (agentMap[a]||0)+1;
                const v = s.interest||"普通";
                interestMap[v] = (interestMap[v]||0)+1;
              });
              const agentList = Object.entries(agentMap).sort((a,b)=>b[1]-a[1]);
              return (
                <div style={{background:"#fff8f3",border:"1px solid #e8d5c0",borderRadius:12,padding:"14px",marginBottom:14}}>
                  <div style={{fontSize:12,color:"#888888",marginBottom:10,fontWeight:600}}>帶看統計</div>
                  <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:10}}>
                    {Object.entries(interestMap).map(([k,v])=>(
                      <div key={k} style={{textAlign:"center"}}>
                        <div style={{fontSize:18,fontWeight:700,color:"#e8722a"}}>{v}</div>
                        <div style={{fontSize:11,color:"#888888"}}>{k}</div>
                      </div>
                    ))}
                  </div>
                  {agentList.map(([agent,count])=>(
                    <div key={agent} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                      <span style={{fontSize:12,fontWeight:600,color:"#1a1a1a",width:50,flexShrink:0}}>{agent}</span>
                      <div style={{flex:1,background:"#f5f0eb",borderRadius:20,height:16,overflow:"hidden"}}>
                        <div style={{height:"100%",background:"#e8722a",borderRadius:20,width:(count/Math.max(...agentList.map(x=>x[1]))*100)+"%"}}/>
                      </div>
                      <span style={{fontSize:12,color:"#e8722a",fontWeight:700,width:20,textAlign:"right"}}>{count}</span>
                    </div>
                  ))}
                </div>
              );
            })()}
            {/* List */}
            {propShowings.map(s=>(
              <div key={s.id} style={{background:"#f5f0eb",border:"1px solid #e0d6ca",borderRadius:12,padding:"12px 14px",marginBottom:9}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:s.feedback?8:0}}>
                  <div>
                    {s.agent&&<span style={{fontWeight:700,fontSize:14,color:"#e8722a"}}>{s.agent}</span>}
                    <span style={{fontSize:13,color:"#1a1a1a",marginLeft:s.agent?6:0,fontWeight:600}}>{s.buyerName||"?"}</span>
                    <span style={{fontSize:12,color:"#888888",marginLeft:6}}>{s.date}{s.time?" "+s.time:""}</span>
                  </div>
                  <IBadge v={s.interest}/>
                </div>
                {s.feedback
                  ? <div style={{fontSize:13,color:"#444444",background:"#ffffff",borderRadius:8,padding:"8px 10px",borderLeft:"3px solid #e8722a",lineHeight:1.6}}>💬 {s.feedback}</div>
                  : <div style={{fontSize:12,color:"#cccccc"}}>（無回饋）</div>
                }
              </div>
            ))}
          </>
        }
      </div>
    );
  }

  const CATS = ["全部","售屋","出租","口約"];
  const CAT_COLORS = {"售屋":"#e8722a","出租":"#3b82f6","口約":"#8b5cf6"};
  const [filterPrice, setFilterPrice] = useState({min:"",max:""});
  const [filterRooms, setFilterRooms] = useState("");
  const [filterFloor, setFilterFloor] = useState("");
  const [showFilter, setShowFilter] = useState(false);
  const [sortBy, setSortBy] = useState("default"); // default | price_asc | price_desc | rent_asc | rent_desc
  const [selectMode, setSelectMode] = useState(false);
  const [selectedIds, setSelectedIds] = useState([]);

  const filteredProps = properties.filter(p=>{
    if(filterCategory!=="全部" && (p.propCategory||"售屋")!==filterCategory) return false;
    // Price filter (售屋/口約)
    if((p.propCategory||"售屋")!=="出租") {
      if(filterPrice.min && p.price < Number(filterPrice.min)) return false;
      if(filterPrice.max && p.price > Number(filterPrice.max)) return false;
    }
    // Rooms filter
    if(filterRooms) {
      const r = parseInt(p.layout?.split("/")[0]||p.rooms||0);
      if(r !== parseInt(filterRooms)) return false;
    }
    // Floor filter
    if(filterFloor && p.floor && !p.floor.includes(filterFloor)) return false;
    return true;
  });
  const hasFilter = filterPrice.min||filterPrice.max||filterRooms||filterFloor;

  // Sort
  if(sortBy==="price_asc") filteredProps.sort((a,b)=>(a.price||0)-(b.price||0));
  else if(sortBy==="price_desc") filteredProps.sort((a,b)=>(b.price||0)-(a.price||0));
  else if(sortBy==="rent_asc") filteredProps.sort((a,b)=>parseInt(a.rent||0)-parseInt(b.rent||0));
  else if(sortBy==="rent_desc") filteredProps.sort((a,b)=>parseInt(b.rent||0)-parseInt(a.rent||0));

  return (
    <div className="page">
      <div className="page-nav">
        <div className="page-title" style={{marginBottom:0}}>物件庫（{filteredProps.length}）</div>
        <div style={{display:"flex",gap:8}}>
          {selectMode ? (
            <>
              <button onClick={()=>{
                if(selectedIds.length>0){
                  setProperties(prev=>prev.filter(p=>!selectedIds.includes(p.id)));
                  setSelectedIds([]); setSelectMode(false);
                }
              }} style={{background:"#ff3b30",border:"none",borderRadius:20,padding:"8px 14px",color:"#fff",fontFamily:"Noto Sans TC,sans-serif",fontSize:13,fontWeight:700,cursor:"pointer"}} disabled={selectedIds.length===0}>
                🗑 刪除（{selectedIds.length}）
              </button>
              <button onClick={()=>{setSelectMode(false);setSelectedIds([]);}} style={{background:"#f5f0eb",border:"1px solid #e0d6ca",borderRadius:20,padding:"8px 14px",color:"#666666",fontFamily:"Noto Sans TC,sans-serif",fontSize:13,cursor:"pointer"}}>取消</button>
            </>
          ) : (
            <>
              <button onClick={()=>setSelectMode(true)} style={{background:"#f5f0eb",border:"1px solid #e0d6ca",borderRadius:20,padding:"8px 12px",color:"#666666",fontFamily:"Noto Sans TC,sans-serif",fontSize:13,cursor:"pointer"}}>☑ 多選</button>
              <button className="add-btn" onClick={()=>openNew(filterCategory==="全部"?"售屋":filterCategory)}>＋ 新增</button>
            </>
          )}
        </div>
      </div>
      {selectMode && selectedIds.length>0 && (
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:"#fff8f3",border:"1px solid #e8d5c0",borderRadius:10,padding:"8px 14px",marginBottom:10}}>
          <span style={{fontSize:13,color:"#e8722a",fontWeight:700}}>已選 {selectedIds.length} 筆</span>
          <button onClick={()=>setSelectedIds(filteredProps.map(p=>p.id))} style={{background:"none",border:"none",color:"#e8722a",fontSize:13,cursor:"pointer",fontFamily:"Noto Sans TC,sans-serif"}}>全選（{filteredProps.length}）</button>
        </div>
      )}

      {/* Category filter tabs */}
      <div style={{display:"flex",gap:6,marginBottom:12,overflowX:"auto",paddingBottom:2}}>
        {CATS.map(cat=>{
          const count = cat==="全部" ? properties.length : properties.filter(p=>(p.propCategory||"售屋")===cat).length;
          const isActive = filterCategory===cat;
          const color = CAT_COLORS[cat]||"#e8722a";
          return (
            <button key={cat} onClick={()=>{setFilterCategory(cat);setSelectMode(false);setSelectedIds([]);setSortBy("default");}} style={{
              flexShrink:0, padding:"8px 14px", borderRadius:20,
              background: isActive ? color : "#f5f0eb",
              border: "1px solid", borderColor: isActive ? color : "#e0d6ca",
              color: isActive ? "#fff" : "#666666",
              fontFamily:"Noto Sans TC,sans-serif", fontSize:13, fontWeight:700, cursor:"pointer"
            }}>
              {cat==="售屋"?"🏠":cat==="出租"?"🔑":cat==="口約"?"🤝":"📋"} {cat} {count>0?`(${count})`:""}
            </button>
          );
        })}
      </div>

      {/* Filter panel */}
      {(filterCategory==="售屋"||filterCategory==="口約"||filterCategory==="全部") && (
        <div style={{marginBottom:10}}>
          <button onClick={()=>setShowFilter(v=>!v)} style={{display:"flex",alignItems:"center",gap:6,background:hasFilter?"#fff8f3":"#f5f0eb",border:"1px solid",borderColor:hasFilter?"#e8722a":"#e0d6ca",borderRadius:20,padding:"7px 14px",color:hasFilter?"#e8722a":"#666666",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"Noto Sans TC,sans-serif"}}>
            🔍 篩選條件 {hasFilter?"（已設定）":""} {showFilter?"▲":"▼"}
          </button>
          {showFilter && (
            <div style={{background:"#f5f0eb",border:"1px solid #e0d6ca",borderRadius:14,padding:"14px",marginTop:8}}>
              <div className="two-col" style={{marginBottom:8}}>
                <div className="field-wrap" style={{marginBottom:0}}>
                  <label className="field-label">最低價（萬）</label>
                  <input type="number" placeholder="500" value={filterPrice.min} onChange={e=>setFilterPrice(p=>({...p,min:e.target.value}))} style={{fontSize:13}}/>
                </div>
                <div className="field-wrap" style={{marginBottom:0}}>
                  <label className="field-label">最高價（萬）</label>
                  <input type="number" placeholder="2000" value={filterPrice.max} onChange={e=>setFilterPrice(p=>({...p,max:e.target.value}))} style={{fontSize:13}}/>
                </div>
              </div>
              <div className="two-col" style={{marginBottom:8}}>
                <div className="field-wrap" style={{marginBottom:0}}>
                  <label className="field-label">房數</label>
                  <select value={filterRooms} onChange={e=>setFilterRooms(e.target.value)} style={{fontSize:13}}>
                    <option value="">不限</option>
                    {["1","2","3","4","5"].map(v=><option key={v} value={v}>{v}房</option>)}
                  </select>
                </div>
                <div className="field-wrap" style={{marginBottom:0}}>
                  <label className="field-label">樓層含</label>
                  <input placeholder="例：3、頂樓、1F" value={filterFloor} onChange={e=>setFilterFloor(e.target.value)} style={{fontSize:13}}/>
                </div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}>
                <label className="field-label" style={{margin:0,whiteSpace:"nowrap"}}>排序</label>
                <select value={sortBy} onChange={e=>setSortBy(e.target.value)} style={{flex:1,padding:"6px 10px",fontSize:12,borderRadius:8,border:"1px solid #e0d6ca",background:"#ffffff",color:"#1a1a1a",fontFamily:"Noto Sans TC,sans-serif"}}>
                  <option value="default">預設順序</option>
                  {(filterCategory==="出租") ? <>
                    <option value="rent_asc">租金 低→高</option>
                    <option value="rent_desc">租金 高→低</option>
                  </> : <>
                    <option value="price_asc">總價 低→高</option>
                    <option value="price_desc">總價 高→低</option>
                  </>}
                </select>
              </div>
              <button onClick={()=>{setFilterPrice({min:"",max:""});setFilterRooms("");setFilterFloor("");setSortBy("default");}} style={{background:"none",border:"1px solid #e0d6ca",borderRadius:8,padding:"6px 12px",color:"#888888",fontSize:12,cursor:"pointer",fontFamily:"Noto Sans TC,sans-serif"}}>
                ✕ 清除篩選
              </button>
              <div style={{fontSize:11,color:"#e8722a",marginTop:6}}>找到 {filteredProps.length} 筆{sortBy!=="default"?" · 已排序":""}</div>
            </div>
          )}
        </div>
      )}

      {/* Quick add buttons when filtered */}
      {filterCategory!=="全部" && (
        <button onClick={()=>openNew(filterCategory)} style={{
          width:"100%", padding:"12px", marginBottom:12,
          background:"#fff8f3", border:"2px dashed",
          borderColor: CAT_COLORS[filterCategory]||"#e8722a",
          borderRadius:12, color: CAT_COLORS[filterCategory]||"#e8722a",
          fontFamily:"Noto Sans TC,sans-serif", fontSize:14, fontWeight:700, cursor:"pointer"
        }}>
          ＋ 新增{filterCategory==="售屋"?"🏠":filterCategory==="出租"?"🔑":"🤝"}{filterCategory}物件
        </button>
      )}

      <div className="import-panel">
        <button className="import-toggle" onClick={()=>{setShowImport(v=>!v);setImportMsg("");setImportStats(null);}}>
          <span>📊 從 Google Sheets 匯入</span>
          <span style={{fontSize:11,color:"#666666"}}>{showImport?"▲":"▼"}</span>
        </button>
        {showImport && (
          <div className="import-body">
            {/* Excel 上傳 */}
            <div style={{marginBottom:14}}>
              <div style={{fontSize:12,fontWeight:700,color:"#1a1a1a",marginBottom:8}}>方法一：直接上傳 Excel 檔案</div>
              <label style={{display:"flex",alignItems:"center",gap:10,background:"#fff8f3",border:"2px dashed #e8722a",borderRadius:12,padding:"14px 16px",cursor:"pointer",color:"#e8722a",fontSize:14,fontWeight:600}}>
                📂 上傳 .xlsx 檔案（自動辨識分頁）
                <input type="file" accept=".xlsx,.xls" onChange={e=>{
                  const file=e.target.files[0];
                  if(!file) return;
                  setImportMsg("讀取中…"); setImportStats(null);
                  const reader=new FileReader();
                  reader.onload=ev=>{
                    try {
                      // Use SheetJS to read Excel
                      const wb = XLSX.read(ev.target.result, {type:"array", cellDates:true});
                      const SKIP = ["主攻開發信","買方","樂善店","買方斡旋中","買方斡","班表","行事曆"];
                      const CATEGORY_MAP = {
                        "口約":"口約", "租件":"出租", "租":"出租",
                        "0-999":"售屋","1000":"售屋","1500":"售屋","2000":"售屋","3000":"售屋",
                      };
                      let totalAdded=0, totalUpdated=0, totalUnchanged=0;
                      const allImported = [];
                      wb.SheetNames.forEach(sheetName=>{
                        if(SKIP.some(s=>sheetName.includes(s))) return;
                        const ws = wb.Sheets[sheetName];
                        const rows = XLSX.utils.sheet_to_json(ws, {defval:""});
                        if(rows.length===0) return;
                        const firstRow = rows[0];
                        const hasRent = "租金" in firstRow;
                        const hasName = "社區(物件)" in firstRow || "社區" in firstRow;
                        if(!hasName) return; // skip non-property sheets
                        let cat = "售屋";
                        if(hasRent || sheetName.includes("租")) cat="出租";
                        else if(sheetName.includes("口約")) cat="口約";
                        const imported = rows
                          .filter(r=>(r["社區(物件)"]||r["社區"]||"").toString().trim())
                          .map(r=>rowToProperty(r, cat));
                        allImported.push(...imported);
                      });
                      // Merge with existing
                      setProperties(currentProps=>{
                        const result=[...currentProps];
                        allImported.forEach(np=>{
                          const exIdx=result.findIndex(e=>e.name===np.name&&(e.propCategory||"售屋")===np.propCategory);
                          if(exIdx===-1){result.push(np);totalAdded++;}
                          else{
                            const ex=result[exIdx];
                            if(ex.price!==np.price||ex.rent!==np.rent||ex.floor!==np.floor){result[exIdx]={...np,id:ex.id};totalUpdated++;}
                            else totalUnchanged++;
                          }
                        });
                        const obj={}; result.forEach(p=>{obj[p.id]=p;}); fbSave("re_props3",obj);
                        return result;
                      });
                      setImportStats({added:totalAdded,updated:totalUpdated,unchanged:totalUnchanged});
                      setImportMsg("ok");
                    } catch(err) {
                      setImportMsg("❌ 讀取失敗："+err.message);
                    }
                  };
                  reader.readAsArrayBuffer(file);
                }} style={{display:"none"}}/>
              </label>
              {importMsg==="讀取中…"&&<div style={{fontSize:12,color:"#e8722a",marginTop:6}}>⏳ 讀取中…</div>}
            </div>
            <div style={{borderTop:"1px solid #e0d6ca",paddingTop:14,marginBottom:10}}>
              <div style={{fontSize:12,fontWeight:700,color:"#1a1a1a",marginBottom:8}}>方法二：從 Google Sheets 複製貼上</div>
            </div>
            <div className="import-steps">
              <div className="import-step" style={{color:"#e8722a",fontWeight:700}}>① 先選擇物件類型（重要！）</div>
              <div className="import-step">② 開啟 Google Sheets → 選分頁 → 全選複製</div>
              <div className="import-step">③ 貼到下方 → 智能匯入</div>
            </div>
            <div style={{display:"flex",gap:6,margin:"8px 0"}}>
              {["售屋","出租","口約"].map(cat=>(
                <button key={cat} onClick={()=>setImportCategory(cat)} style={{flex:1,padding:"8px",background:importCategory===cat?"#e8722a":"#f5f0eb",border:"1px solid",borderColor:importCategory===cat?"#e8722a":"#e0d6ca",borderRadius:8,color:importCategory===cat?"#fff":"#666666",fontFamily:"Noto Sans TC,sans-serif",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  {cat==="售屋"?"🏠":cat==="出租"?"🔑":"🤝"} {cat}
                </button>
              ))}
            </div>
            <div className="field-wrap" style={{marginTop:10}}>
              <label className="field-label">貼上資料</label>
              <textarea rows={5} placeholder="從 Google Sheets 複製貼上…" value={pasteText} onChange={e=>{setPasteText(e.target.value);setImportMsg("");setImportStats(null);}} style={{fontSize:12,fontFamily:"monospace"}}/>
            </div>
            <button className="btn-gold" onClick={doImport} disabled={!pasteText.trim()}>⬇️ 智能匯入</button>
            {importMsg==="ok"&&importStats&&<div className="import-msg ok">✅ 新增 {importStats.added} 筆，更新 {importStats.updated} 筆，未變動 {importStats.unchanged} 筆</div>}
            {importMsg&&importMsg!=="ok"&&<div className="import-msg">{importMsg}</div>}
          </div>
        )}
      </div>
      {filteredProps.length===0 ? <Empty icon="🏠" text={filterCategory==="全部"?"尚無物件":"尚無"+filterCategory+"物件"}/> : filteredProps.map(p=>(
        <div className="list-card" key={p.id} onClick={()=>{ if(selectMode){ setSelectedIds(prev=>prev.includes(p.id)?prev.filter(x=>x!==p.id):[...prev,p.id]); } else openDetail(p); }} style={{flexDirection:"column",gap:0,alignItems:"stretch",background:selectedIds.includes(p.id)?"#fff8f3":"#ffffff",borderColor:selectedIds.includes(p.id)?"#e8722a":"#e0d6ca",position:"relative"}}>
            {selectMode && <div style={{position:"absolute",top:12,right:12,width:22,height:22,borderRadius:6,background:selectedIds.includes(p.id)?"#e8722a":"#fff",border:"2px solid",borderColor:selectedIds.includes(p.id)?"#e8722a":"#cccccc",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2}}>{selectedIds.includes(p.id)&&<span style={{color:"#fff",fontSize:13,fontWeight:700}}>✓</span>}</div>}
          <div style={{display:"flex",alignItems:"flex-start",gap:10}}>
            <div className="list-card-main" style={{flex:1}}>
              <div className="list-card-title">
                {p._fromSheet&&<span className="sheet-dot">●</span>}
                {p.propCategory&&p.propCategory!=="售屋"&&(
                  <span style={{background:p.propCategory==="出租"?"#3b82f6":"#8b5cf6",color:"#fff",borderRadius:4,padding:"1px 6px",fontSize:10,fontWeight:700,marginRight:6}}>{p.propCategory==="出租"?"🔑 出租":"🤝 口約"}</span>
                )}
                {p.name}
              </div>
              <div className="list-card-sub">
                {(p.propCategory||"售屋")==="出租"
                  ? [p.district, p.layout||(p.rooms?(p.rooms+"房"):""), p.type, p.floor, p.rent?(Number(p.rent).toLocaleString()+"元/月"):""].filter(Boolean).join(" · ")
                  : [p.district, p.areaMain?(((Number(p.areaMain)||0)+(Number(p.areaSub)||0)+(Number(p.areaCommon)||0)+(Number(p.areaParking)||0)).toFixed(1)+"坪"):p.area?(p.area+"坪"):"", p.layout||(p.rooms?(p.rooms+"房"):""), p.type, p.floor].filter(Boolean).join(" · ")
                }
              </div>
              {p.notes&&<div className="list-card-note">{p.notes.slice(0,35)}{p.notes.length>35?"…":""}</div>}
            </div>
            <div style={{textAlign:"right",flexShrink:0}}>
              {(p.propCategory||"售屋")==="出租"
                ? <div className="price-tag" style={{fontSize:13}}>{p.rent?(Number(p.rent).toLocaleString()+"元/月"):""}</div>
                : <div className="price-tag">{p.price?(p.price+"萬"):""}</div>
              }
            </div>
          </div>
          {(()=>{
            const n=s=>(s||"").trim();
            const fromS=(showings||[]).filter(s=>s.propertyId===p.id||n(s.propName)===n(p.name));
            const fromB=[];
            (buyers||[]).forEach(b=>{(b.shownProps||[]).forEach(sp=>{if(sp.propId===p.id||n(sp.propName)===n(p.name)){if(!fromS.find(s=>s.buyerId===b.id&&s.date===sp.date&&(s.feedback||"")===(sp.feedback||"")))fromB.push({agent:b.agent||"",buyerName:b.name});}}); });
            const all=[...fromS,...fromB];
            if(all.length===0) return null;
            return (
              <div style={{marginTop:8,paddingTop:8,borderTop:"1px solid #e8d5c0"}}>
                <span style={{fontSize:12,color:"#e8722a",fontWeight:700}}>🏃 帶看 {all.length} 次</span>
              </div>
            );
          })()}
        </div>
      ))}
    </div>
  );
}

// ── Shared UI ─────────────────────────────────────────────────────────────────
function Empty({icon,text}){ return <div className="empty"><div style={{fontSize:36,marginBottom:10}}>{icon}</div>{text}</div>; }
function InfoItem({label,val}){ return <div className="info-item"><div className="info-item-label">{label}</div><div className="info-item-val">{val}</div></div>; }
function IBadge({v}){
  const m={"高":"#4ade80","中":"#fbbf24","低":"#f87171","普通":"#6b7280","婉拒":"#ef4444"};
  return <span className="badge" style={{background:m[v]||"#6b7280",color:"#0c1118"}}>{v}</span>;
}
function BSBadge({v}){
  const m={"尚未看屋":"#374151","初看":"#3b82f6","複看":"#8b5cf6","斡旋中":"#f59e0b","已成交":"#4ade80"};
  return <span className="badge" style={{background:m[v]||"#374151",color:v==="尚未看屋"?"#9ca3af":"#0c1118"}}>{v}</span>;
}

// ── CSS ───────────────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&family=DM+Serif+Display&display=swap');
*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent;}
body{font-family:'Noto Sans TC',sans-serif;background:#ffffff;color:#1a1a1a;overscroll-behavior:none;}
.shell{max-width:480px;margin:0 auto;min-height:100vh;display:flex;flex-direction:column;background:#ffffff;}
.topbar{background:linear-gradient(135deg,#f5f0eb,#ffffff);border-bottom:1px solid #e0d6ca;padding:14px 18px 12px;flex-shrink:0;}
.topbar-eyebrow{font-size:9px;letter-spacing:3px;color:#e8722a;margin-bottom:2px;font-weight:500;}
.topbar-title{font-family:'DM Serif Display',serif;font-size:22px;color:#1a1a1a;line-height:1.1;}
.topbar-title span{color:#e8722a;}
.body{flex:1;overflow-y:auto;padding-bottom:72px;}
.bottomnav{position:fixed;bottom:0;left:50%;transform:translateX(-50%);width:100%;max-width:480px;background:#ffffff;border-top:1px solid #e0d6ca;display:flex;padding:6px 2px calc(6px + env(safe-area-inset-bottom));z-index:20;}
.navbtn{flex:1;display:flex;flex-direction:column;align-items:center;gap:1px;background:none;border:none;color:#aaaaaa;cursor:pointer;padding:3px 1px;border-radius:8px;transition:color .2s;}
.navbtn.active{color:#e8722a;}
.navicon{font-size:18px;line-height:1;}
.navlabel{font-size:9px;font-family:'Noto Sans TC',sans-serif;font-weight:500;}
.page{padding:16px 14px;}
.page-title{font-family:'DM Serif Display',serif;font-size:22px;color:#1a1a1a;margin-bottom:16px;}
.page-nav{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;}
.back-btn{background:none;border:none;color:#e8722a;font-family:'Noto Sans TC',sans-serif;font-size:14px;cursor:pointer;padding:4px 0;}
.edit-btn,.add-btn{background:#e8722a;border:none;color:#0c1118;font-family:'Noto Sans TC',sans-serif;font-size:13px;font-weight:700;padding:8px 14px;border-radius:20px;cursor:pointer;}
.section-hd{font-size:10px;letter-spacing:2px;color:#e8722a;text-transform:uppercase;margin:14px 0 8px;font-weight:500;}
.stat-grid{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:18px;}
.stat-card{background:#f5f0eb;border:1px solid #e0d6ca;border-radius:14px;padding:14px;text-align:center;}
.stat-accent{background:linear-gradient(135deg,#f0f5e8,#f5f0eb);border-color:#2a3a18;}
.stat-icon{font-size:20px;margin-bottom:5px;}
.stat-val{font-family:'DM Serif Display',serif;font-size:26px;color:#1a1a1a;line-height:1;}
.stat-label{font-size:10px;color:#888888;margin-top:3px;}
.quick-grid{display:grid;grid-template-columns:1fr 1fr;gap:9px;}
.quick-btn{background:#f5f0eb;border:1px solid #e0d6ca;border-radius:14px;padding:16px 10px;display:flex;flex-direction:column;align-items:center;cursor:pointer;color:#1a1a1a;font-family:'Noto Sans TC',sans-serif;}
.quick-btn:active{background:#1e2535;}
.list-card{background:#f5f0eb;border:1px solid #e0d6ca;border-radius:14px;padding:13px 14px;margin-bottom:9px;display:flex;align-items:flex-start;gap:10px;cursor:pointer;transition:border-color .2s;}
.list-card:active{border-color:#e8722a;}
.list-card-main{flex:1;min-width:0;}
.list-card-title{font-size:15px;font-weight:700;color:#1a1a1a;margin-bottom:3px;}
.list-card-sub{font-size:12px;color:#888888;}
.list-card-note{font-size:11px;color:#666666;margin-top:3px;}
.badge{border-radius:20px;padding:4px 9px;font-size:11px;font-weight:700;white-space:nowrap;flex-shrink:0;}
.price-tag{color:#e8722a;font-size:14px;font-weight:700;white-space:nowrap;flex-shrink:0;}
.field-wrap{margin-bottom:11px;}
.field-label{display:block;font-size:11px;color:#666666;margin-bottom:5px;font-weight:500;}
input,select,textarea{width:100%;background:#ffffff;border:1px solid #e0d6ca;border-radius:10px;padding:11px 13px;color:#1a1a1a;font-family:'Noto Sans TC',sans-serif;font-size:15px;outline:none;transition:border-color .2s;-webkit-appearance:none;appearance:none;}
input:focus,select:focus,textarea:focus{border-color:#e8722a;}
select option{background:#f5f0eb;}
textarea{resize:none;}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:9px;}
.btn-gold{width:100%;background:linear-gradient(135deg,#e8722a,#d4611a);border:none;border-radius:12px;padding:15px;color:#0c1118;font-family:'Noto Sans TC',sans-serif;font-size:15px;font-weight:700;cursor:pointer;letter-spacing:1px;transition:opacity .2s;}
.btn-gold:active{opacity:.85;}
.btn-gold:disabled{opacity:.4;}
.btn-danger{width:100%;background:#fff3f3;border:1px solid #ffcdd2;border-radius:12px;padding:13px;color:#f87171;font-family:'Noto Sans TC',sans-serif;font-size:14px;font-weight:600;cursor:pointer;}
.warn-box{background:#fffde7;border:1px solid #ffe082;border-radius:10px;padding:11px 13px;color:#fbbf24;font-size:13px;margin-bottom:11px;line-height:1.5;}
.loading-box{display:flex;flex-direction:column;align-items:center;padding:36px 20px;gap:14px;color:#666666;font-size:13px;}
.spinner{width:34px;height:34px;border:3px solid #e0d6ca;border-top-color:#e8722a;border-radius:50%;animation:spin .8s linear infinite;}
@keyframes spin{to{transform:rotate(360deg);}}
.divider{height:1px;background:#1e2535;margin:18px 0;}
.result-card{background:linear-gradient(135deg,#f0f5e8,#f5f0eb);border:1px solid #1e3a1e;border-radius:14px;padding:15px;margin-bottom:11px;position:relative;overflow:hidden;}
.result-card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,#4ade80,#e8722a);}
.score-badge{position:absolute;top:13px;right:13px;border-radius:20px;padding:4px 11px;font-size:13px;font-weight:700;color:#0c1118;}
.score-badge.high{background:#4ade80;}.score-badge.mid{background:#e8722a;}.score-badge.low{background:#f87171;}
.result-name{font-size:15px;font-weight:700;padding-right:68px;margin-bottom:4px;}
.result-label{font-size:11px;color:#666666;margin-bottom:7px;}
.ai-reason{background:rgba(201,169,110,.08);border-left:3px solid #e8722a;border-radius:0 8px 8px 0;padding:9px 11px;font-size:13px;color:#d1c9be;line-height:1.6;}
.ai-hint{background:#e8f5e9;border:1px solid #1e3a1e;border-radius:10px;padding:11px 13px;color:#2e7d32;font-size:13px;margin-bottom:13px;line-height:1.6;}
.empty{text-align:center;color:#cccccc;padding:36px 20px;font-size:14px;}
.detail-hero{background:#f5f0eb;border:1px solid #e0d6ca;border-radius:14px;padding:16px;margin-bottom:14px;display:flex;flex-direction:column;gap:7px;}
.detail-name{font-family:'DM Serif Display',serif;font-size:24px;color:#1a1a1a;}
.phone-link{color:#e8722a;font-size:15px;text-decoration:none;}
.info-grid{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-bottom:11px;}
.info-item{background:#f5f0eb;border:1px solid #e0d6ca;border-radius:10px;padding:9px 11px;}
.info-item-label{font-size:10px;color:#888888;margin-bottom:3px;}
.info-item-val{font-size:13px;color:#1a1a1a;font-weight:600;}
.info-box{background:#f5f0eb;border:1px solid #e0d6ca;border-radius:10px;padding:11px 13px;margin-bottom:9px;font-size:14px;color:#d1c9be;line-height:1.6;}
.info-box-label{font-size:10px;color:#888888;margin-bottom:3px;}
.loan-result{background:#f5f0eb;border:1px solid #e0d6ca;border-radius:16px;padding:18px;margin-top:14px;}
.loan-main{text-align:center;padding-bottom:14px;border-bottom:1px solid #e0d6ca;margin-bottom:14px;}
.loan-main-label{font-size:12px;color:#888888;margin-bottom:5px;letter-spacing:1px;}
.loan-main-val{font-family:'DM Serif Display',serif;font-size:30px;color:#e8722a;}
.loan-grid{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:11px;}
.loan-item{background:#ffffff;border-radius:10px;padding:11px;}
.loan-item-label{font-size:11px;color:#888888;margin-bottom:3px;}
.loan-item-val{font-size:15px;font-weight:700;color:#1a1a1a;}
.loan-note{font-size:11px;color:#aaaaaa;text-align:center;}
.import-panel{background:#f5f0eb;border:1px solid #e0d6ca;border-radius:14px;margin-bottom:13px;overflow:hidden;}
.import-toggle{width:100%;display:flex;justify-content:space-between;align-items:center;padding:13px 15px;background:none;border:none;color:#1a1a1a;font-family:'Noto Sans TC',sans-serif;font-size:14px;font-weight:600;cursor:pointer;}
.import-body{padding:0 15px 15px;}
.import-steps{background:#ffffff;border-radius:10px;padding:11px;margin-bottom:4px;}
.import-step{font-size:12px;color:#666666;line-height:2;}
.import-msg{margin-top:9px;padding:9px 11px;border-radius:10px;font-size:13px;background:#fffde7;color:#fbbf24;border:1px solid #ffe082;}
.import-msg.ok{background:#e8f5e9;color:#2e7d32;border-color:#a5d6a7;}
.sheet-dot{color:#2e7d32;font-size:8px;margin-right:4px;vertical-align:middle;}
.chip-group{display:flex;flex-wrap:wrap;gap:7px;margin-top:4px;}
.chip{background:#ffffff;border:1px solid #d5c9b8;border-radius:20px;padding:6px 14px;color:#888888;font-family:'Noto Sans TC',sans-serif;font-size:13px;cursor:pointer;transition:all .2s;}
.chip.active{background:#e8722a;border-color:#e8722a;color:#0c1118;font-weight:700;}
.shown-card{background:#f5f0eb;border:1px solid #e0d6ca;border-radius:12px;padding:13px;margin-bottom:9px;}
.mini-form{background:#f5f0eb;border:1px solid #d5c9b8;border-radius:14px;padding:14px;margin-bottom:14px;}
.event-card{background:#f5f0eb;border:1px solid #e0d6ca;border-radius:12px;padding:12px 14px;margin-bottom:8px;display:flex;align-items:flex-start;gap:10px;}
.event-card.today{border-color:#e8722a;background:linear-gradient(135deg,#1a1810,#141b27);}
.event-time{font-size:13px;font-weight:700;color:#e8722a;white-space:nowrap;min-width:40px;margin-top:2px;}
.event-main{flex:1;}
.event-title{font-size:14px;font-weight:700;color:#1a1a1a;margin-bottom:3px;}
.event-sub{font-size:12px;color:#888888;margin-top:2px;}
.cal-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;}
.cal-month{font-family:'DM Serif Display',serif;font-size:18px;color:#1a1a1a;}
.cal-nav{background:none;border:1px solid #d5c9b8;border-radius:8px;color:#e8722a;font-size:20px;width:36px;height:36px;cursor:pointer;display:flex;align-items:center;justify-content:center;}
.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;margin-bottom:16px;}
.cal-wday{text-align:center;font-size:11px;color:#aaaaaa;padding:4px 0;font-weight:600;}
.cal-day{aspect-ratio:1;display:flex;flex-direction:column;align-items:center;justify-content:center;border-radius:10px;cursor:pointer;font-size:13px;color:#666666;gap:2px;transition:all .15s;}
.cal-day:active{background:#1e2535;}
.cal-day.today{color:#e8722a;font-weight:700;}
.cal-day.selected{background:#1e2535;color:#1a1a1a;}
.cal-day.today.selected{background:rgba(201,169,110,.15);}
.cal-dot-row{display:flex;gap:2px;}
.cal-dot{width:4px;height:4px;border-radius:50%;background:#e8722a;}
`;
