import { useState, useEffect, useRef, useMemo } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid } from "recharts";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://rmokkujpnptiugqlovvf.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJtb2trdWpwbnB0aXVncWxvdnZmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg5MjkzNzQsImV4cCI6MjA5NDUwNTM3NH0.4CIPkHwqTn37hnXryIN_ebJKAZ9P1Oc6JyzEnHzo1vA";
const TABLE = "inventory_batches";

const sb = {
  async select(table) {
    if (!SUPABASE_URL) return { data: [], error: null };
    const limit = 1000;
    let all = [], offset = 0;
    while (true) {
      const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}?order=expiry_date.asc&limit=${limit}&offset=${offset}`, {
        headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}`, Prefer: "count=none" },
      });
      if (!res.ok) return { data: [], error: await res.text() };
      const batch = await res.json();
      all = [...all, ...batch];
      if (batch.length < limit) break;
      offset += limit;
    }
    return { data: all, error: null };
  },
  async upsert(table, row) {
    if (!SUPABASE_URL) return { error: null };
    const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}`, {
      method: "POST",
      headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}`, "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" },
      body: JSON.stringify(row),
    });
    if (res.ok) return { error: null };
    const patchRes = await fetch(`${SUPABASE_URL}/rest/v1/${table}?id=eq.${row.id}`, {
      method: "PATCH",
      headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}`, "Content-Type": "application/json", Prefer: "return=minimal" },
      body: JSON.stringify(row),
    });
    return { error: patchRes.ok ? null : await patchRes.text() };
  },
  async delete(table, id) {
    if (!SUPABASE_URL) return { error: null };
    const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}?id=eq.${id}`, {
      method: "DELETE",
      headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` },
    });
    return { error: res.ok ? null : await res.text() };
  },
};

let _idb = null;
function openIDB() {
  return new Promise((res, rej) => {
    if (_idb) return res(_idb);
    const r = indexedDB.open("InvProV4", 1);
    r.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains("batches")) db.createObjectStore("batches", { keyPath: "id" });
      if (!db.objectStoreNames.contains("settings")) db.createObjectStore("settings", { keyPath: "k" });
    };
    r.onsuccess = e => { _idb = e.target.result; res(_idb); };
    r.onerror = () => rej(r.error);
  });
}
async function idbAll() { const db = await openIDB(); return new Promise((res,rej) => { const r=db.transaction("batches","readonly").objectStore("batches").getAll(); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error); }); }
async function idbPut(item) { const db = await openIDB(); return new Promise((res,rej) => { const r=db.transaction("batches","readwrite").objectStore("batches").put(item); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
async function idbDel(id) { const db = await openIDB(); return new Promise((res,rej) => { const r=db.transaction("batches","readwrite").objectStore("batches").delete(id); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); }); }
async function idbGet(k) { try { const db=await openIDB(); return new Promise(res => { const r=db.transaction("settings","readonly").objectStore("settings").get(k); r.onsuccess=()=>res(r.result?.v); r.onerror=()=>res(null); }); } catch { return null; } }
async function idbSet(k,v) { const db=await openIDB(); return new Promise((res,rej)=>{ const r=db.transaction("settings","readwrite").objectStore("settings").put({k,v}); r.onsuccess=()=>res(); r.onerror=()=>rej(); }); }


function daysLeft(d) { if(!d)return null; const t=new Date();t.setHours(0,0,0,0);const e=new Date(d);e.setHours(0,0,0,0);return Math.floor((e-t)/86400000); }
function tierOf(days) { if(days===null)return"none";if(days<0)return"expired";if(days<30)return"red";if(days<90)return"yellow";if(days<=180)return"green";return"safe"; }

const TIER = {
  expired:{ label:"已過期",   color:"#dc2626", bg:"#fef2f2", border:"#fca5a5" },
  red:    { label:"緊急<30天", color:"#ea580c", bg:"#fff7ed", border:"#fdba74" },
  yellow: { label:"促銷規劃",  color:"#ca8a04", bg:"#fefce8", border:"#fde047" },
  green:  { label:"正常銷售",  color:"#16a34a", bg:"#f0fdf4", border:"#86efac" },
  safe:   { label:"安全>180",  color:"#2563eb", bg:"#eff6ff", border:"#93c5fd" },
  none:   { label:"無效期",    color:"#9ca3af", bg:"#f9fafb", border:"#e5e7eb" },
};

function genId() { return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g,c=>{const r=Math.random()*16|0;return(c==="x"?r:(r&0x3|0x8)).toString(16);}); }
function fmtMoney(n) { return n?`$${Number(n).toLocaleString()}`:"—"; }
function fmtDate(d) { return d?d.slice(0,10):""; }
function newBatch(o={}) { return { id:genId(),barcode:"",product_no:"",name:"",category:"食品",expiry_date:"",qty:0,unit:"個",cost:0,price:0,location:"",supplier:"",note:"",created_at:new Date().toISOString(),...o }; }
function toSB(r) { return {...r, expiry_date:r.expiry_date||null, qty:Number(r.qty)||0, cost:Number(r.cost)||0, price:Number(r.price)||0 }; }

function buildChart(items) {
  return Array.from({length:4},(_,i)=>{
    const now=new Date(); const start=new Date(now.getFullYear(),now.getMonth()+i,1); const end=new Date(now.getFullYear(),now.getMonth()+i+1,0);
    const label=`${start.getMonth()+1}月`;
    let red=0,yellow=0,green=0;
    items.forEach(b=>{ if(!b.expiry_date)return; const exp=new Date(b.expiry_date); if(exp>=start&&exp<=end){ const v=(b.qty||0)*(b.cost||0); const t=tierOf(daysLeft(b.expiry_date)); if(t==="red"||t==="expired")red+=v; else if(t==="yellow")yellow+=v; else green+=v; } });
    return { label, red:Math.round(red), yellow:Math.round(yellow), green:Math.round(green) };
  });
}

const ROW_H=38, HDR_H=42, SCAN=8;
function useVScroll(n, ref) {
  const [top,setTop]=useState(0),[vh,setVh]=useState(500);
  useEffect(()=>{ const el=ref.current; if(!el)return; const ro=new ResizeObserver(()=>setVh(el.clientHeight)); ro.observe(el); const h=()=>setTop(el.scrollTop); el.addEventListener("scroll",h,{passive:true}); return()=>{ro.disconnect();el.removeEventListener("scroll",h);}; },[]);
  const s=Math.max(0,Math.floor(top/ROW_H)-SCAN), e=Math.min(n-1,Math.ceil((top+vh)/ROW_H)+SCAN);
  return {s,e,totalH:n*ROW_H,offsetY:s*ROW_H};
}

const COLS=[
  {k:"barcode",    lbl:"條碼",     w:120,ed:true},
  {k:"product_no",  lbl:"產品編號",  w:130,ed:true},
  {k:"name",       lbl:"商品名稱", w:170,ed:true},
  {k:"category",   lbl:"類別",     w:85, ed:true,type:"sel"},
  {k:"expiry_date",lbl:"有效日期", w:115,ed:true,type:"date"},
  {k:"_days",      lbl:"剩餘天數", w:85, ed:false},
  {k:"_tier",      lbl:"燈號",     w:95, ed:false},
  {k:"qty",        lbl:"庫存量",   w:70, ed:true,type:"num"},
  {k:"unit",       lbl:"單位",     w:55, ed:true},
  {k:"cost",       lbl:"成本",     w:85, ed:true,type:"num"},
  {k:"price",      lbl:"售價",     w:85, ed:true,type:"num"},
  {k:"location",   lbl:"儲位",     w:80, ed:true},
  {k:"supplier",   lbl:"供應商",   w:120,ed:true},
  {k:"note",       lbl:"備註",     w:140,ed:true},
];

// Light theme style constants
const fldSt = { width:"100%", background:"#fff", border:"1px solid #d1d5db", color:"#111827", padding:"8px 10px", borderRadius:6, fontSize:13, fontFamily:"inherit" };
const aBox  = { background:"#fff", border:"1px solid #e5e7eb", borderRadius:8, padding:16, marginBottom:12, boxShadow:"0 1px 3px rgba(0,0,0,0.06)" };
const aH1   = { fontSize:16, fontWeight:700, color:"#111827", marginBottom:14 };
const aH2   = { fontSize:11, color:"#6b7280", letterSpacing:0.5, marginBottom:8, fontWeight:600, textTransform:"uppercase" };
const codeS = { background:"#f8fafc", border:"1px solid #e2e8f0", borderRadius:6, padding:"12px", fontSize:11, color:"#475569", lineHeight:1.7, fontFamily:"monospace", whiteSpace:"pre-wrap", wordBreak:"break-all" };
const cpyBtn = { background:"#f3f4f6", border:"1px solid #d1d5db", color:"#374151", padding:"5px 12px", borderRadius:5, cursor:"pointer", fontSize:11, fontFamily:"inherit" };

const SCHEMA = [
  "CREATE TABLE IF NOT EXISTS inventory_batches (",
  "  id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),",
  "  barcode       TEXT,",
  "  product_no    TEXT,",
  "  name         TEXT NOT NULL,",
  "  batch_no     TEXT,",
  "  category     TEXT DEFAULT '食品',",
  "  expiry_date  DATE,",
  "  qty          INTEGER DEFAULT 0,",
  "  unit         TEXT DEFAULT '個',",
  "  cost         NUMERIC(12,2) DEFAULT 0,",
  "  price        NUMERIC(12,2) DEFAULT 0,",
  "  location     TEXT,",
  "  supplier     TEXT,",
  "  note         TEXT,",
  "  created_at   TIMESTAMPTZ DEFAULT NOW()",
  ");",
  "CREATE INDEX IF NOT EXISTS idx_exp ON inventory_batches(expiry_date);",
  "CREATE INDEX IF NOT EXISTS idx_barcode ON inventory_batches(barcode);",
  "ALTER TABLE inventory_batches ENABLE ROW LEVEL SECURITY;",
  "CREATE POLICY allow_all ON inventory_batches FOR ALL USING (true) WITH CHECK (true);",
].join("\n");

const AUTO_GUIDE = [
  "每天自動 LINE 推播 — 設定概覽",
  "",
  "工具：Supabase Edge Function + pg_cron",
  "",
  "步驟：",
  "1. 安裝 Supabase CLI (npm i -g supabase)",
  "2. supabase login",
  "3. supabase functions new send-expiry-alert",
  "4. 把推播邏輯程式碼放入 index.ts",
  "5. 在 Supabase Dashboard > Settings > Secrets",
  "   新增 LINE_NOTIFY_TOKEN",
  "6. supabase functions deploy send-expiry-alert",
  "7. Supabase Dashboard > Database > Extensions",
  "   啟用 pg_cron，設定每天 00:00 UTC 執行",
  "   （= 台灣時間早上 8:00）",
  "",
  "不想自己架伺服器可用：",
  "· Make.com 無程式碼串接 LINE Notify",
  "· Zapier 串接 LINE Notify",
].join("\n");

export default function App() {
  const [items,setItems]         = useState([]);
  const [loaded,setLoaded]       = useState(false);
  const [tab,setTab]             = useState("dashboard");
  const [adminTab,setAdminTab]   = useState("overview");
  const [search,setSearch]       = useState("");
  const [filterTier,setFT]       = useState("all");
  const [filterCat,setFC]        = useState("all");
  const [dynCats,setDynCats]     = useState([]);
  const [sortK,setSortK]         = useState("expiry_date");
  const [sortD,setSortD]         = useState(1);
  const [editCell,setEditCell]   = useState(null);
  const [editVal,setEditVal]     = useState("");
  const [selected,setSelected]   = useState(new Set());
  const [colW,setColW]           = useState(()=>Object.fromEntries(COLS.map(c=>[c.k,c.w])));
  const [showForm,setShowForm]   = useState(false);
  const [form,setForm]           = useState(newBatch());
  const [toast,setToast]         = useState(null);
  const [botLog,setBotLog]       = useState([]);
  const [lineToken,setLineToken] = useState("");
  const [lineMsg,setLineMsg]     = useState("");
  const [importLog,setImportLog] = useState(null);
  const [importing,setImporting] = useState(false);
  const [importProgress,setImportProgress] = useState({current:0,total:0});
  const [fifoSku,setFifoSku]     = useState("");
  const [isOnline,setIsOnline]   = useState(false);
  const [syncing,setSyncing]     = useState(false);
  const [stRecords,setStRecords] = useState([]);
  const [stBarcodes,setStBarcodes] = useState("");
  const [stOperator,setStOp]     = useState("");
  const [stPending,setStPending] = useState(null);
  const [page,setPage]           = useState(1);
  const [expandedRows,setExpandedRows] = useState(new Set());
  const PAGE_SIZE = 1000;

  const gridRef = useRef(null);
  const editRef = useRef(null);
  const fileRef = useRef(null);

  useEffect(()=>{
    (async()=>{
      let data=[];
      if(SUPABASE_URL){ setSyncing(true); const {data:d,error}=await sb.select(TABLE); if(!error){data=d||[];setIsOnline(true);}else data=await idbAll(); setSyncing(false); }
    
      setItems(data); setLoaded(true);
      // 從資料自動抓取所有類別
      const cats=[...new Set(data.map(b=>b.category).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"zh"));
      setDynCats(cats);
    })();
    idbGet("botLog").then(v=>v&&setBotLog(v));
    idbGet("lineToken").then(v=>v&&setLineToken(v));
  },[]);

  useEffect(()=>{ if(loaded&&items.length>0)runBot(items,false); },[loaded]);

  const showToast=(msg,type="ok")=>{ setToast({msg,type}); setTimeout(()=>setToast(null),3000); };

  async function saveItem(item) {
    // 如果是新類別，加入動態類別清單
    if(item.category&&!dynCats.includes(item.category)){
      setDynCats(prev=>[...new Set([...prev,item.category])].sort((a,b)=>a.localeCompare(b,"zh")));
    }
    const clean={...item}; if(!clean.expiry_date)clean.expiry_date=null;
    await idbPut(clean);
    setItems(prev=>{ const i=prev.findIndex(x=>x.id===clean.id); if(i>=0){const n=[...prev];n[i]=clean;return n;} return [...prev,clean]; });
    if(SUPABASE_URL){ const {error}=await sb.upsert(TABLE,toSB(clean)); if(error)showToast("雲端同步失敗","error"); }
  }

  async function deleteItems(ids) {
    for(const id of ids){ await idbDel(id); if(SUPABASE_URL)await sb.delete(TABLE,id); }
    setItems(prev=>prev.filter(x=>!ids.has(x.id))); setSelected(new Set());
    showToast(`已刪除 ${ids.size} 筆`,"warning");
  }

  const rows=useMemo(()=>{
    let r=items.map(b=>({...b,_days:daysLeft(b.expiry_date),_tier:tierOf(daysLeft(b.expiry_date))}));
    if(search){const q=search.toLowerCase();r=r.filter(x=>x.name?.toLowerCase().includes(q)||x.barcode?.toLowerCase().includes(q)||x.batch_no?.toLowerCase().includes(q)||x.supplier?.toLowerCase().includes(q));}
    if(filterTier!=="all")r=r.filter(x=>x._tier===filterTier);
    if(filterCat!=="all")r=r.filter(x=>x.category===filterCat);
    r.sort((a,b)=>{ let av=a[sortK]??"",bv=b[sortK]??""; if(sortK==="_days"){av=a._days??9999;bv=b._days??9999;} return(typeof av==="number"?av-bv:String(av).localeCompare(String(bv),"zh"))*sortD; });
    // 標記每個條碼的批次資訊
    const barcodeMap={};
    r.forEach(row=>{
      const key=String(row.barcode||"").trim()||row.id;
      if(!barcodeMap[key])barcodeMap[key]=[];
      barcodeMap[key].push(row);
    });
    return r.map(row=>{
      const key=String(row.barcode||"").trim()||row.id;
      const batches=barcodeMap[key];
      return {...row,_batches:batches,_batchCount:batches.length,_isFirstBatch:batches[0].id===row.id};
    });
  },[items,search,filterTier,filterCat,sortK,sortD]);

  const totalPages=Math.ceil(rows.length/PAGE_SIZE);
  const pagedRows=rows.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);
  useEffect(()=>setPage(1),[search,filterTier,filterCat]);
  const {s:vs,e:ve,totalH,offsetY}=useVScroll(pagedRows.length,gridRef);
  const visRows=pagedRows.slice(vs,ve+1);

  const stats=useMemo(()=>{
    const c={total:items.length,expired:0,red:0,yellow:0,green:0,safe:0};let tv=0,rv=0;
    items.forEach(b=>{ const t=tierOf(daysLeft(b.expiry_date));if(c[t]!==undefined)c[t]++;else c.safe++;const v=(b.qty||0)*(b.cost||0);tv+=v;if(t==="red"||t==="expired")rv+=v; });
    return {...c,totalVal:tv,riskVal:rv};
  },[items]);

  const chartData=useMemo(()=>buildChart(items),[items]);

  function startEdit(id,k,val){ const col=COLS.find(c=>c.k===k);if(!col?.ed)return;setEditCell({id,k});setEditVal(String(val??""));setTimeout(()=>editRef.current?.focus(),10); }
  async function commitEdit(){ if(!editCell)return;const b=items.find(x=>x.id===editCell.id);if(!b){setEditCell(null);return;}const col=COLS.find(c=>c.k===editCell.k);let val=editVal;if(col?.type==="num")val=parseFloat(val)||0;await saveItem({...b,[editCell.k]:val});setEditCell(null); }
  function toggleSort(k){ if(sortK===k)setSortD(d=>-d);else{setSortK(k);setSortD(1);} }

  const fifoRows=useMemo(()=>{ if(!fifoSku.trim())return[];return items.filter(b=>b.barcode===fifoSku.trim()&&b.qty>0).sort((a,b)=>new Date(a.expiry_date)-new Date(b.expiry_date)); },[items,fifoSku]);

  async function runBot(data,notify=true){
    const today=new Date().toLocaleDateString("zh-TW");const alerts=[];
    data.forEach(b=>{ const d=daysLeft(b.expiry_date);const t=tierOf(d);
      if(t==="expired")alerts.push({t:"expired",msg:`🔴 ${b.name}（${b.batch_no||"—"}）已過期 ${Math.abs(d)} 天`});
      else if(t==="red")alerts.push({t:"red",msg:`🟠 ${b.name}（${b.batch_no||"—"}）剩 ${d} 天，請特賣`});
      else if(t==="yellow")alerts.push({t:"yellow",msg:`🟡 ${b.name}（${b.batch_no||"—"}）剩 ${d} 天，規劃促銷`});
    });
    const entry={date:today,count:alerts.length,alerts};const log=[entry,...botLog.slice(0,19)];
    setBotLog(log);await idbSet("botLog",log);
    if(notify&&alerts.length>0)showToast(`機器人：${alerts.length} 筆異常`,"warning");
  }

  function genLineMsg(){
    const today=new Date().toLocaleDateString("zh-TW");
    const list=items.filter(b=>{const d=daysLeft(b.expiry_date);return d!==null&&d<90;}).sort((a,b)=>daysLeft(a.expiry_date)-daysLeft(b.expiry_date)).slice(0,25);
    if(!list.length){setLineMsg(`【庫存效期報告】${today}\n✅ 全部正常`);return;}
    const lines=list.map(b=>{ const d=daysLeft(b.expiry_date);const t=tierOf(d);const icon=t==="expired"?"🔴":t==="red"?"🟠":"🟡";return `${icon} ${b.name}（${b.batch_no||"—"}）剩${d<0?`過期${Math.abs(d)}天`:d+"天"} ×${b.qty}${b.unit}`; });
    setLineMsg(`【庫存效期報告】${today}\n共 ${list.length} 筆需注意：\n\n${lines.join("\n")}\n\n⚙️ QYIM 庫存有效管理`);
  }

  async function handleImport(e){
    const file=e.target.files[0];if(!file)return;
    const reader=new FileReader();
    reader.onload=async ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary",cellDates:true});
        const raw=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});
        const MAP={"SKU":"barcode","sku":"barcode","barcode":"barcode","條碼":"barcode","產品編號":"product_no","product_no":"product_no","商品名稱":"name","name":"name","批號":"batch_no","有效日期":"expiry_date","到期日":"expiry_date","類別":"category","庫存量":"qty","qty":"qty","數量":"qty","單位":"unit","成本":"cost","成本價":"cost","售價":"price","儲位":"location","供應商":"supplier","備註":"note"};
        let ok=0,updated=0;
        setImporting(true);
        setImportProgress({current:0,total:raw.length});
        for(const row of raw){
          const b=newBatch();
          Object.entries(row).forEach(([k,v])=>{ const mk=MAP[k.trim()];if(!mk)return;if(mk==="expiry_date"){const toLocal=d=>{const y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,"0"),dd=String(d.getDate()).padStart(2,"0");return y+"-"+m+"-"+dd;};b[mk]=v instanceof Date?toLocal(v):typeof v==="string"&&v?toLocal(new Date(v)):"";}else b[mk]=v; });
          if(!b.name&&!b.barcode)continue;
          // 判斷邏輯：條碼 + 有效日期 組合 → 同批；條碼相同+效期不同 → 新批次
          const barcode=String(b.barcode||"").trim();
          const productNo=String(b.product_no||"").trim();
          const expiryDate=String(b.expiry_date||"").trim();
          // 先用條碼+效期找完全相同的批次
          const exactMatch=barcode&&expiryDate
            ?items.find(x=>String(x.barcode||"").trim()===barcode&&String(x.expiry_date||"").trim()===expiryDate)
            :null;
          // 若沒有完全相同，且沒有效期，再用產品編號找
          const noDateMatch=!exactMatch&&!expiryDate&&productNo
            ?items.find(x=>String(x.product_no||"").trim()===productNo&&!x.expiry_date)
            :null;
          const existing=exactMatch||noDateMatch;
          if(existing){
            // 只用 Excel 有填的欄位覆蓋，空白欄位保留原本的值
            const merged={...existing};
            Object.entries(b).forEach(([k,v])=>{
              if(k==="id"||k==="created_at")return;
              const isEmpty=v===null||v===undefined||String(v).trim()===""||((k==="cost"||k==="price")?`${v}`==="0":false);
              if(!isEmpty)merged[k]=v;
            });
            await saveItem({...merged,id:existing.id,created_at:existing.created_at});
            updated++;
          } else {
            // 條碼相同但效期不同 → 新批次，複製基本資料
            if(barcode&&!exactMatch){
              const sameBarcode=items.find(x=>String(x.barcode||"").trim()===barcode);
              if(sameBarcode){
                // 繼承基本資料（名稱、類別、單位、成本、售價、供應商等），只有效期和庫存用新的
                const inherited={...sameBarcode,id:genId(),created_at:new Date().toISOString()};
                Object.entries(b).forEach(([k,v])=>{
                  if(k==="id"||k==="created_at")return;
                  const isEmpty=v===null||v===undefined||String(v).trim()===""||((k==="cost"||k==="price")?`${v}`==="0":false);
                  if(!isEmpty)inherited[k]=v;
                });
                await saveItem(inherited);
                ok++;setImportProgress({current:ok,total:raw.length});
                continue;
              }
            }
            await saveItem(b);
          }
          ok++;
          setImportProgress({current:ok,total:raw.length});
        }
        setImporting(false);
        setImportLog({ok,updated,added:ok-updated,total:raw.length,file:file.name});showToast(`✅ 匯入完成：新增 ${ok-updated} 筆，更新 ${updated} 筆`);
      }catch(err){setImporting(false);showToast("匯入失敗："+err.message,"error");}
    };
    reader.readAsBinaryString(file);e.target.value="";
  }

  function downloadTemplate(){ const ws=XLSX.utils.aoa_to_sheet([["條碼","商品名稱","批號","有效日期","類別","庫存量","單位","成本價","售價","儲位","供應商","備註"],["4710000000001","示範商品","LOT001","2026-06-30","食品",100,"個",50,120,"A-01","供應商甲","範例"]]);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"庫存");XLSX.writeFile(wb,"inventory_template.xlsx"); }

  function startResize(k,e){ e.preventDefault();const sx=e.clientX,sw=colW[k];const mv=ev=>setColW(p=>({...p,[k]:Math.max(50,sw+ev.clientX-sx)}));const up=()=>{window.removeEventListener("mousemove",mv);window.removeEventListener("mouseup",up);};window.addEventListener("mousemove",mv);window.addEventListener("mouseup",up); }

  const totalW=COLS.reduce((s,c)=>s+colW[c.k],0)+46;

  const TABS=[
    {id:"dashboard",icon:"📊",lbl:"看板"},
    {id:"grid",     icon:"⊞", lbl:"庫存總表"},
    {id:"fifo",     icon:"⇄", lbl:"FIFO"},
    {id:"bot",      icon:"🤖",lbl:"機器人"},
    {id:"stocktake",icon:"📋",lbl:"盤點"},
    {id:"import",   icon:"⬆", lbl:"匯入"},
    {id:"line",     icon:"💬",lbl:"LINE"},
    {id:"admin",    icon:"⚙️", lbl:"後台"},
  ];

  return (
    <div style={{height:"100vh",display:"flex",flexDirection:"column",background:"#f3f4f6",color:"#111827",fontFamily:"'Noto Sans TC',Arial,sans-serif",overflow:"hidden"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        ::-webkit-scrollbar{width:6px;height:6px;}
        ::-webkit-scrollbar-track{background:#f1f5f9;}
        ::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:3px;}
        ::-webkit-scrollbar-thumb:hover{background:#94a3b8;}
        .rh:hover{background:#f8fafc!important;}
        .hb:hover{background:#f9fafb!important;}
        input:focus,select:focus,textarea:focus{outline:none!important;border-color:#3b82f6!important;box-shadow:0 0 0 3px rgba(59,130,246,0.1)!important;}
        @keyframes fadeUp{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
        @keyframes spin{to{transform:rotate(360deg)}}
        .fade{animation:fadeUp .18s ease;}
        textarea{resize:vertical;}
        .tab-active{border-bottom:2px solid #2563eb;color:#2563eb;font-weight:600;}

        /* 全站膠囊按鈕樣式 */
        button{transition:all 0.18s ease!important;cursor:pointer!important;}
        .btn-primary{border-radius:999px!important;background:#2563eb!important;color:#fff!important;border:none!important;font-weight:600!important;box-shadow:0 2px 8px rgba(37,99,235,0.25)!important;}
        .btn-primary:hover{background:#1d4ed8!important;box-shadow:0 4px 16px rgba(37,99,235,0.4)!important;transform:translateY(-1px)!important;}
        .btn-secondary{border-radius:999px!important;background:#f3f4f6!important;color:#374151!important;border:1px solid #e5e7eb!important;font-weight:500!important;}
        .btn-secondary:hover{background:#e5e7eb!important;border-color:#d1d5db!important;transform:translateY(-1px)!important;}
        .btn-danger{border-radius:999px!important;background:#fee2e2!important;color:#dc2626!important;border:1px solid #fca5a5!important;font-weight:600!important;}
        .btn-danger:hover{background:#dc2626!important;color:#fff!important;transform:translateY(-1px)!important;}
        .btn-tab:hover{background:#eff6ff!important;color:#2563eb!important;border-radius:999px!important;}
        .btn-sm{border-radius:999px!important;font-size:11px!important;font-weight:600!important;}
        .btn-sm:hover{transform:translateY(-1px)!important;box-shadow:0 3px 10px rgba(0,0,0,0.12)!important;}
      `}</style>

      {/* TOP BAR */}
      <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",height:56,display:"flex",alignItems:"center",padding:"0 16px",gap:8,flexShrink:0,overflowX:"auto",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
        <div style={{display:"inline-flex",alignItems:"center",marginRight:8,flexShrink:0,background:"linear-gradient(135deg,#2563eb,#7c3aed)",borderRadius:999,padding:"6px 18px",boxShadow:"0 2px 8px rgba(37,99,235,0.25)",gap:6}}>
          <span style={{fontWeight:800,fontSize:13,color:"rgba(255,255,255,0.7)",letterSpacing:1}}>QYIM</span>
          <span style={{width:1,height:14,background:"rgba(255,255,255,0.35)",display:"inline-block"}}/>
          <span style={{fontWeight:700,fontSize:14,color:"#fff",letterSpacing:0.5,whiteSpace:"nowrap"}}>庫存有效管理</span>
        </div>
        <div style={{width:1,height:24,background:"#e5e7eb",margin:"0 4px",flexShrink:0}}/>
        {/* Connection badge */}
        <div style={{display:"flex",alignItems:"center",gap:5,padding:"3px 10px",borderRadius:20,background:isOnline?"#f0fdf4":"#fff7ed",border:`1px solid ${isOnline?"#86efac":"#fdba74"}`,color:isOnline?"#16a34a":"#ea580c",fontSize:11,fontWeight:600,flexShrink:0}}>
          {syncing?<span style={{animation:"spin 1s linear infinite",display:"inline-block",fontSize:12}}>↻</span>:<span style={{width:7,height:7,borderRadius:4,background:"currentColor",display:"inline-block"}}/>}
          {syncing?"同步中":isOnline?"雲端連線中":"本機模式"}
        </div>
        <div style={{width:1,height:24,background:"#e5e7eb",margin:"0 4px",flexShrink:0}}/>
        {/* Tabs */}
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{background:tab===t.id?"#1d4ed8":"#e5e7eb",border:"none",borderRadius:999,color:tab===t.id?"#fff":"#374151",padding:"6px 14px",fontSize:12,fontFamily:"inherit",fontWeight:tab===t.id?600:400,flexShrink:0,transition:"all 0.18s",cursor:"pointer",boxShadow:tab===t.id?"0 4px 12px rgba(37,99,235,0.35)":"none"}} onMouseEnter={e=>{if(tab!==t.id){e.currentTarget.style.background="#1d4ed8";e.currentTarget.style.color="#fff";e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 4px 12px rgba(37,99,235,0.35)";}}} onMouseLeave={e=>{if(tab!==t.id){e.currentTarget.style.background="#e5e7eb";e.currentTarget.style.color="#374151";e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="none";}}}>
            {t.icon} {t.lbl}
          </button>
        ))}
        <div style={{flex:1}}/>
        <button onClick={()=>{setForm(newBatch());setShowForm(true);}} style={{background:"#800020",color:"#fff",border:"none",borderRadius:999,padding:"9px 22px",fontSize:13,fontFamily:"inherit",fontWeight:600,flexShrink:0,cursor:"pointer",transition:"all 0.18s",boxShadow:"0 2px 8px rgba(128,0,32,0.3)"}} onMouseEnter={e=>{e.currentTarget.style.background="#1C1C1E";e.currentTarget.style.color="#80DEEA";e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 4px 16px rgba(0,0,0,0.35)";}} onMouseLeave={e=>{e.currentTarget.style.background="#800020";e.currentTarget.style.color="#fff";e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="0 2px 8px rgba(128,0,32,0.3)";}}>＋ 新增批號</button>
      </div>

      {/* STATS BAR */}
      <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",display:"flex",flexShrink:0,overflowX:"auto",padding:"0 8px"}}>
        {[
          {key:"all",    lbl:"全部",    val:stats.total,   col:"#2563eb", bg:"#eff6ff"},
          {key:"expired",lbl:"已過期",  val:stats.expired, col:"#dc2626", bg:"#fef2f2"},
          {key:"red",    lbl:"緊急<30", val:stats.red,     col:"#ea580c", bg:"#fff7ed"},
          {key:"yellow", lbl:"促銷規劃",val:stats.yellow,  col:"#ca8a04", bg:"#fefce8"},
          {key:"green",  lbl:"正常",    val:stats.green,   col:"#16a34a", bg:"#f0fdf4"},
          {key:"safe",   lbl:"安全>180",val:stats.safe,    col:"#2563eb", bg:"#eff6ff"},
        ].map(s=>(
          <div key={s.key} onClick={()=>setFT(s.key)} style={{padding:"8px 16px",borderBottom:filterTier===s.key?`3px solid ${s.col}`:"3px solid transparent",cursor:"pointer",flexShrink:0,transition:"all 0.15s",background:filterTier===s.key?s.bg:"transparent"}}>
            <div style={{color:s.col,fontSize:20,fontWeight:700,lineHeight:1}}>{s.val}</div>
            <div style={{color:"#9ca3af",fontSize:10,marginTop:2,whiteSpace:"nowrap"}}>{s.lbl}</div>
          </div>
        ))}
        <div style={{padding:"8px 16px",borderRight:"1px solid #f3f4f6",flexShrink:0}}>
          <div style={{color:"#ea580c",fontSize:20,fontWeight:700,lineHeight:1}}>{fmtMoney(stats.riskVal)}</div>
          <div style={{color:"#9ca3af",fontSize:10,marginTop:2}}>風險金額</div>
        </div>
        <div style={{padding:"8px 16px",flexShrink:0}}>
          <div style={{color:"#2563eb",fontSize:20,fontWeight:700,lineHeight:1}}>{fmtMoney(stats.totalVal)}</div>
          <div style={{color:"#9ca3af",fontSize:10,marginTop:2}}>庫存總值</div>
        </div>
        <div style={{flex:1}}/>
        <div style={{display:"flex",gap:8,alignItems:"center",padding:"0 8px"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 搜尋條碼 / 商品 / 批號" style={{background:"#f9fafb",border:"1px solid #e5e7eb",color:"#374151",padding:"6px 12px",borderRadius:7,fontSize:12,width:200,fontFamily:"inherit"}}/>
          <select value={filterCat} onChange={e=>setFC(e.target.value)} style={{background:"#f9fafb",border:"1px solid #e5e7eb",color:"#374151",padding:"6px 10px",borderRadius:7,fontSize:12,fontFamily:"inherit"}}>
            <option value="all">所有類別</option>{dynCats.map(c=><option key={c}>{c}</option>)}
          </select>
          {selected.size>0&&<button onClick={()=>deleteItems(selected)} style={{background:"#fef2f2",border:"1px solid #fca5a5",color:"#dc2626",padding:"6px 12px",borderRadius:6,cursor:"pointer",fontSize:12,fontFamily:"inherit",fontWeight:600}}>🗑 刪除 {selected.size} 筆</button>}
        </div>
      </div>

      {/* ══ DASHBOARD ══ */}
      {tab==="dashboard"&&(
        <div style={{flex:1,overflow:"auto",padding:20,display:"flex",flexDirection:"column",gap:16}} className="fade">
          <div style={{fontSize:13,fontWeight:700,color:"#374151",letterSpacing:0.5}}>📈 未來 4 個月到期庫存金額</div>
          <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:"16px 8px 8px",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={chartData} barGap={4}>
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false}/>
                <XAxis dataKey="label" tick={{fill:"#9ca3af",fontSize:12}} axisLine={false} tickLine={false}/>
                <YAxis tick={{fill:"#9ca3af",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>v>=1000?`${(v/1000).toFixed(0)}K`:v}/>
                <Tooltip contentStyle={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:8,fontSize:12,boxShadow:"0 4px 12px rgba(0,0,0,0.1)"}} formatter={(v,n)=>[`$${v.toLocaleString()}`,n==="red"?"緊急":n==="yellow"?"促銷規劃":"正常"]}/>
                <Bar dataKey="red"    fill="#ef4444" radius={[4,4,0,0]}/>
                <Bar dataKey="yellow" fill="#f59e0b" radius={[4,4,0,0]}/>
                <Bar dataKey="green"  fill="#22c55e" radius={[4,4,0,0]}/>
              </BarChart>
            </ResponsiveContainer>
            <div style={{display:"flex",gap:16,justifyContent:"center",marginTop:6}}>
              {[["#ef4444","緊急<30天"],["#f59e0b","促銷60-90天"],["#22c55e","正常≤180天"]].map(([c,l])=>(
                <div key={l} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:"#6b7280"}}><div style={{width:10,height:10,background:c,borderRadius:2}}/>{l}</div>
              ))}
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))",gap:14}}>
            <div style={aBox}>
              <div style={aH2}>⚠️ 需要行動的品項</div>
              {["expired","red","yellow"].flatMap(t=>items.filter(b=>tierOf(daysLeft(b.expiry_date))===t).slice(0,3).map(b=>(
                <div key={b.id} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"6px 0",borderBottom:"1px solid #f3f4f6",alignItems:"center"}}>
                  <span style={{background:TIER[t].bg,color:TIER[t].color,border:`1px solid ${TIER[t].border}`,padding:"1px 6px",borderRadius:4,fontSize:10,fontWeight:600,flexShrink:0,marginRight:8}}>{TIER[t].label}</span>
                  <span style={{flex:1,color:"#374151",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{b.name}</span>
                  <span style={{color:TIER[t].color,flexShrink:0,marginLeft:8,fontWeight:600}}>×{b.qty}</span>
                </div>
              )))}
              {items.filter(b=>["expired","red","yellow"].includes(tierOf(daysLeft(b.expiry_date)))).length===0&&<div style={{color:"#22c55e",fontSize:13,fontWeight:600}}>✅ 全部庫存正常</div>}
            </div>
            <div style={aBox}>
              <div style={aH2}>⚡ 快捷操作</div>
              {[
                {lbl:"⊞ 庫存總表",        act:()=>setTab("grid"),  color:"#2563eb"},
                {lbl:"⇄ FIFO 出貨查詢",   act:()=>setTab("fifo"),  color:"#7c3aed"},
                {lbl:"💬 產生 LINE 推播", act:()=>{genLineMsg();setTab("line");}, color:"#16a34a"},
                {lbl:"⬆ 匯入 Excel",     act:()=>setTab("import"), color:"#ea580c"},
                {lbl:"⚙️ 後台管理介面",   act:()=>setTab("admin"),  color:"#6b7280"},
              ].map((a,i)=>(
                <button key={i} onClick={a.act} className="hb" style={{display:"flex",alignItems:"center",gap:8,width:"100%",textAlign:"left",background:"#e5e7eb",color:"#374151",border:"none",padding:"10px 16px",borderRadius:999,cursor:"pointer",fontSize:13,fontFamily:"inherit",marginBottom:6,transition:"all 0.18s"}} onMouseEnter={e=>{e.currentTarget.style.background="#1d4ed8";e.currentTarget.style.color="#fff";e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 4px 12px rgba(37,99,235,0.35)";}} onMouseLeave={e=>{e.currentTarget.style.background="#e5e7eb";e.currentTarget.style.color="#374151";e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="none";}}>
                  {a.lbl}
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ══ GRID ══ */}
      {tab==="grid"&&(
        <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
          <div style={{flex:1,overflow:"auto",background:"#fff"}} ref={gridRef}>
            <div style={{minWidth:totalW}}>
              {/* Header */}
              <div style={{display:"flex",background:"#f8fafc",borderBottom:"2px solid #e2e8f0",position:"sticky",top:0,zIndex:50,height:HDR_H}}>
                <div style={{width:46,minWidth:46,padding:"0 12px",display:"flex",alignItems:"center",borderRight:"1px solid #e2e8f0"}}>
                  <input type="checkbox" checked={selected.size===rows.length&&rows.length>0} onChange={e=>setSelected(e.target.checked?new Set(rows.map(r=>r.id)):new Set())} style={{accentColor:"#2563eb",cursor:"pointer",width:15,height:15}}/>
                </div>
                {COLS.map(col=>(
                  <div key={col.k} style={{width:colW[col.k],minWidth:colW[col.k],padding:"0 10px",display:"flex",alignItems:"center",justifyContent:"space-between",borderRight:"1px solid #e2e8f0",cursor:"pointer",userSelect:"none",position:"relative",color:sortK===col.k?"#2563eb":"#6b7280",fontSize:11,fontWeight:600,letterSpacing:0.5,textTransform:"uppercase"}}
                    className="hb" onClick={()=>toggleSort(col.k)}>
                    <span>{col.lbl}</span>
                    <span style={{opacity:sortK===col.k?1:0.2,fontSize:9}}>{sortD===1?"▲":"▼"}</span>
                    <div onMouseDown={e=>startResize(col.k,e)} style={{position:"absolute",right:0,top:0,width:4,height:"100%",cursor:"col-resize"}} onClick={e=>e.stopPropagation()}/>
                  </div>
                ))}
              </div>
              {/* Rows */}
              <div style={{height:totalH,position:"relative"}}>
                <div style={{position:"absolute",top:offsetY,left:0,right:0}}>
                  {visRows.map((row,vi)=>{
                    const isSel=selected.has(row.id); const tm=TIER[row._tier];
                    return (
                      <div key={row.id} className="rh" style={{display:"flex",height:ROW_H,alignItems:"center",background:isSel?"#eff6ff":(vs+vi)%2===0?"#fff":"#fafafa",borderBottom:"1px solid #f1f5f9",borderLeft:isSel?"3px solid #2563eb":"3px solid transparent"}}>
                        <div style={{width:46,minWidth:46,padding:"0 12px",display:"flex",alignItems:"center",borderRight:"1px solid #f1f5f9"}}>
                          <input type="checkbox" checked={isSel} onChange={()=>setSelected(prev=>{const n=new Set(prev);n.has(row.id)?n.delete(row.id):n.add(row.id);return n;})} style={{accentColor:"#2563eb",cursor:"pointer",width:15,height:15}}/>
                        </div>
                        {COLS.map(col=>{
                          const isEd=editCell?.id===row.id&&editCell?.k===col.k;
                          let disp=row[col.k]??"";
                          if(col.k==="_days")disp=row._days===null?"—":row._days<0?`${row._days}天`:`+${row._days}天`;
                          if(col.k==="_tier")disp=tm.label;
                          return (
                            <div key={col.k} style={{width:colW[col.k],minWidth:colW[col.k],height:"100%",padding:"0 10px",display:"flex",alignItems:"center",borderRight:"1px solid #f1f5f9",cursor:col.ed?"cell":"default",fontSize:12,
                              color:col.k==="_days"?tm.color:col.k==="name"?"#111827":col.k==="barcode"?"#2563eb":"#4b5563",
                              background:isEd?"#eff6ff":"transparent",outline:isEd?"2px solid #3b82f6":"none",outlineOffset:"-2px"}}
                              onDoubleClick={()=>col.ed&&startEdit(row.id,col.k,row[col.k])}>
                              {isEd?(col.type==="sel"?(<select autoFocus value={editVal} onChange={e=>setEditVal(e.target.value)} onBlur={commitEdit} style={{width:"100%",background:"#fff",border:"none",color:"#111827",fontSize:12,fontFamily:"inherit"}}>{dynCats.map(c=><option key={c}>{c}</option>)}</select>):(<input ref={editRef} value={editVal} type={col.type==="num"?"number":col.type==="date"?"date":"text"} onChange={e=>setEditVal(e.target.value)} onBlur={commitEdit} onKeyDown={e=>{if(e.key==="Enter")commitEdit();if(e.key==="Escape")setEditCell(null);}} style={{width:"100%",background:"transparent",border:"none",color:"#111827",fontSize:12,fontFamily:"inherit"}}/>)):(
                                col.k==="_tier"&&row._tier!=="none"?(
                                  <span style={{background:tm.bg,color:tm.color,border:`1px solid ${tm.border}`,padding:"2px 7px",borderRadius:4,fontSize:10,fontWeight:600,whiteSpace:"nowrap"}}>{disp}</span>
                                ):col.k==="expiry_date"&&row._batchCount>1?(
                                  <div style={{width:"100%",position:"relative"}}>
                                    <button onClick={e=>{e.stopPropagation();setExpandedRows(prev=>{const n=new Set(prev);const key=String(row.barcode||"").trim();n.has(key)?n.delete(key):n.add(key);return n;});}} style={{background:"none",border:"none",cursor:"pointer",display:"flex",alignItems:"center",gap:4,fontSize:12,color:"#2563eb",fontWeight:500,padding:0,width:"100%"}}>
                                      <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1,textAlign:"left"}}>{disp||"—"}</span>
                                      <span style={{fontSize:10,background:"#eff6ff",color:"#2563eb",borderRadius:999,padding:"1px 6px",flexShrink:0,fontWeight:700}}>{row._batchCount}批</span>
                                      <span style={{fontSize:10,transition:"transform 0.2s",transform:expandedRows.has(String(row.barcode||"").trim())?"rotate(180deg)":"rotate(0deg)"}}>▼</span>
                                    </button>
                                    {expandedRows.has(String(row.barcode||"").trim())&&(
                                      <div style={{position:"absolute",top:"100%",left:0,zIndex:200,background:"#fff",border:"1px solid #e5e7eb",borderRadius:8,boxShadow:"0 8px 24px rgba(0,0,0,0.12)",minWidth:220,padding:8}}>
                                        <div style={{fontSize:10,color:"#9ca3af",fontWeight:600,marginBottom:6,letterSpacing:0.5}}>所有批次（依效期排序）</div>
                                        {row._batches.sort((a,b)=>(a.expiry_date||"")>(b.expiry_date||"")?1:-1).map((bt,bi)=>{
                                          const btDays=daysLeft(bt.expiry_date);const btTier=tierOf(btDays);const btTm=TIER[btTier];
                                          return(
                                            <div key={bt.id} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 6px",borderRadius:5,background:bi%2===0?"#f8fafc":"#fff",marginBottom:2}}>
                                              <span style={{fontSize:10,color:"#9ca3af",minWidth:20,fontWeight:600}}>#{bi+1}</span>
                                              <span style={{fontSize:11,color:"#374151",fontWeight:500,flex:1}}>{bt.expiry_date?bt.expiry_date.slice(0,10):"無效期"}</span>
                                              <span style={{fontSize:11,color:"#111827",fontWeight:600,minWidth:40,textAlign:"right"}}>{bt.qty}{bt.unit}</span>
                                              <span style={{background:btTm.bg,color:btTm.color,border:`1px solid ${btTm.border}`,padding:"1px 5px",borderRadius:4,fontSize:9,fontWeight:600,whiteSpace:"nowrap"}}>{btTm.label}</span>
                                            </div>
                                          );
                                        })}
                                      </div>
                                    )}
                                  </div>
                                ):(
                                  <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",width:"100%",fontWeight:col.k==="name"?500:400,...(col.k==="cost"||col.k==="price"?{color:"#16a34a",fontWeight:500}:{})}}>
                                    {col.k==="cost"||col.k==="price"?fmtMoney(disp):String(disp)}
                                  </span>
                                )
                              )}
                            </div>
                          );
                        })}
                      </div>
                    );
                  })}
                </div>
              </div>
              {rows.length===0&&loaded&&<div style={{padding:60,textAlign:"center",color:"#9ca3af",fontSize:14}}>📦 {items.length===0?"尚無資料，請點「＋ 新增批號」或匯入 Excel":"沒有符合條件的批號"}</div>}
            </div>
          </div>
          <div style={{background:"#f8fafc",borderTop:"1px solid #e2e8f0",padding:"5px 16px",fontSize:11,color:"#9ca3af",display:"flex",gap:12,alignItems:"center",justifyContent:"space-between"}}>
            <div style={{display:"flex",gap:12,alignItems:"center"}}>
              <span>顯示 {(page-1)*PAGE_SIZE+1}–{Math.min(page*PAGE_SIZE,rows.length)} / {rows.length} 筆（共 {items.length} 筆）</span><span>·</span><span>{isOnline?"✅ 已同步 Supabase":"💾 IndexedDB 本機"}</span>
            </div>
            {totalPages>1&&<div style={{display:"flex",gap:6,alignItems:"center"}}>
              <button onClick={()=>setPage(p=>Math.max(1,p-1))} disabled={page===1} style={{background:page===1?"#f3f4f6":"#fff",border:"1px solid #e5e7eb",color:page===1?"#d1d5db":"#374151",padding:"3px 10px",borderRadius:5,cursor:page===1?"default":"pointer",fontSize:11,fontFamily:"inherit"}}>‹ 上一頁</button>
              {Array.from({length:totalPages},(_,i)=>i+1).map(p=>(
                <button key={p} onClick={()=>setPage(p)} style={{background:p===page?"#2563eb":"#fff",border:"1px solid #e5e7eb",color:p===page?"#fff":"#374151",padding:"3px 10px",borderRadius:5,cursor:"pointer",fontSize:11,fontFamily:"inherit",fontWeight:p===page?700:400}}>{p}</button>
              ))}
              <button onClick={()=>setPage(p=>Math.min(totalPages,p+1))} disabled={page===totalPages} style={{background:page===totalPages?"#f3f4f6":"#fff",border:"1px solid #e5e7eb",color:page===totalPages?"#d1d5db":"#374151",padding:"3px 10px",borderRadius:5,cursor:page===totalPages?"default":"pointer",fontSize:11,fontFamily:"inherit"}}>下一頁 ›</button>
            </div>}
            <button onClick={()=>{const r=rows.map(b=>({條碼:b.barcode,商品名稱:b.name,批號:b.batch_no,有效日期:b.expiry_date?b.expiry_date.slice(0,10):"",類別:b.category||"",庫存量:b.qty,單位:b.unit||"",成本:b.cost||"",售價:b.price||"",儲位:b.location||"",供應商:b.supplier||"",備註:b.note||"",剩餘天數:daysLeft(b.expiry_date)??"",狀態:TIER[tierOf(daysLeft(b.expiry_date))].label}));const ws=XLSX.utils.json_to_sheet(r);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"庫存");XLSX.writeFile(wb,`inventory_filter_${new Date().toISOString().slice(0,10)}.xlsx`);showToast("匯出完成 ✅");}} style={{background:"#e5e7eb",color:"#374151",border:"none",padding:"5px 14px",borderRadius:999,cursor:"pointer",fontSize:11,fontWeight:600,fontFamily:"inherit",whiteSpace:"nowrap",transition:"all 0.18s"}} onMouseEnter={e=>{e.currentTarget.style.background="#1d4ed8";e.currentTarget.style.color="#fff";e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 4px 12px rgba(37,99,235,0.35)";}} onMouseLeave={e=>{e.currentTarget.style.background="#e5e7eb";e.currentTarget.style.color="#374151";e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="none";}}>⬇ 匯出篩選結果</button>
          </div>
        </div>
      )}

      {/* ══ FIFO ══ */}
      {tab==="fifo"&&(
        <div style={{flex:1,overflow:"auto",padding:20}} className="fade">
          <div style={{maxWidth:640}}>
            <div style={aH1}>⇄ FIFO 先進先出出貨查詢</div>
            <div style={{color:"#6b7280",fontSize:13,marginBottom:16}}>依到期日由近到遠排序，第 1 批為建議優先出貨批次。</div>
            <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap"}}>
              <input value={fifoSku} onChange={e=>setFifoSku(e.target.value)} placeholder="輸入條碼查詢…" style={{flex:1,minWidth:140,...fldSt}}/>
              {[...new Set(items.map(b=>b.barcode).filter(Boolean))].slice(0,10).map(s=>(<button key={s} onClick={()=>setFifoSku(s)} style={{background:"#eff6ff",border:"1px solid #bfdbfe",color:"#2563eb",padding:"5px 12px",borderRadius:5,cursor:"pointer",fontSize:12,fontFamily:"inherit",fontWeight:500}}>{s}</button>))}
            </div>
            {fifoRows.length>0?fifoRows.map((b,i)=>{ const d=daysLeft(b.expiry_date);const tm=TIER[tierOf(d)];return (
              <div key={b.id} style={{...aBox,display:"flex",alignItems:"center",gap:12,borderColor:i===0?"#bfdbfe":"#e5e7eb",background:i===0?"#eff6ff":"#fff"}}>
                <div style={{width:30,height:30,borderRadius:15,background:i===0?"#2563eb":"#f3f4f6",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,color:i===0?"#fff":"#9ca3af",fontWeight:700,flexShrink:0}}>{i+1}</div>
                <div style={{flex:1}}>
                  <div style={{color:"#111827",fontSize:13,fontWeight:600}}>{b.name} <span style={{color:"#6b7280",fontSize:11,fontWeight:400}}>批號：{b.batch_no||"—"}</span></div>
                  <div style={{display:"flex",gap:12,marginTop:4,fontSize:12,flexWrap:"wrap"}}>
                    <span style={{color:"#6b7280"}}>到期：{fmtDate(b.expiry_date)}</span>
                    <span style={{color:tm.color,fontWeight:600}}>{d===null?"—":d<0?`過期${Math.abs(d)}天`:`剩${d}天`}</span>
                    <span style={{color:"#16a34a"}}>庫存：{b.qty}{b.unit}</span>
                  </div>
                </div>
                {i===0&&<div style={{background:"#2563eb",color:"#fff",padding:"4px 10px",borderRadius:5,fontSize:11,fontWeight:600,flexShrink:0}}>出貨優先</div>}
              </div>
            );}):fifoSku.trim()?<div style={{color:"#9ca3af",fontSize:13}}>找不到條碼「{fifoSku}」的批號</div>:<div style={{color:"#9ca3af",fontSize:13}}>請輸入條碼開始查詢</div>}
          </div>
        </div>
      )}

      {/* ══ BOT ══ */}
      {tab==="bot"&&(
        <div style={{flex:1,overflow:"auto",padding:20}} className="fade">
          <div style={{maxWidth:640}}>
            <div style={aH1}>🤖 提醒機器人</div>
            <div style={{color:"#6b7280",fontSize:13,marginBottom:16}}>每次開啟自動掃描全庫存，以「異常管理」為核心。</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:16}}>
              {[{t:"expired",icon:"🔴",action:"立即下架"},{t:"red",icon:"🟠",action:"即期特賣"},{t:"yellow",icon:"🟡",action:"規劃促銷"}].map(r=>{
                const cnt=items.filter(b=>tierOf(daysLeft(b.expiry_date))===r.t).length;
                return <div key={r.t} style={{...aBox,textAlign:"center",borderColor:cnt>0?TIER[r.t].border:"#e5e7eb",background:cnt>0?TIER[r.t].bg:"#fff"}}>
                  <div style={{fontSize:24}}>{r.icon}</div>
                  <div style={{color:TIER[r.t].color,fontSize:24,fontWeight:700,margin:"4px 0"}}>{cnt}</div>
                  <div style={{color:"#6b7280",fontSize:11}}>建議：{r.action}</div>
                </div>;
              })}
            </div>
            <button onClick={()=>runBot(items,true)} style={{background:"#2563eb",border:"none",color:"#fff",padding:"10px 24px",borderRadius:7,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:600,marginBottom:16,boxShadow:"0 1px 3px rgba(37,99,235,0.3)"}}>▶ 立即掃描全庫存</button>
            <div style={{fontSize:13,fontWeight:600,color:"#374151",marginBottom:10}}>掃描記錄</div>
            {botLog.map((log,i)=>(
              <div key={i} style={aBox}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:6,fontSize:12}}>
                  <span style={{color:"#374151",fontWeight:600}}>{log.date}</span>
                  <span style={{color:log.count>0?"#ea580c":"#16a34a",fontWeight:600}}>{log.count>0?`⚠ ${log.count} 筆異常`:"✅ 全部正常"}</span>
                </div>
                {log.alerts.slice(0,5).map((a,j)=><div key={j} style={{color:TIER[a.t]?.color||"#6b7280",fontSize:12,padding:"2px 0"}}>{a.msg}</div>)}
                {log.alerts.length>5&&<div style={{color:"#9ca3af",fontSize:11,marginTop:4}}>…還有 {log.alerts.length-5} 筆</div>}
              </div>
            ))}
            {botLog.length===0&&<div style={{color:"#9ca3af",fontSize:13}}>尚無掃描記錄</div>}
          </div>
        </div>
      )}


      {/* ══ STOCKTAKE 盤點 ══ */}
      {tab==="stocktake"&&(
        <div style={{flex:1,overflow:"auto",padding:20}} className="fade">
          <div style={{maxWidth:800}}>
            <div style={aH1}>📋 庫存盤點</div>
            <div style={aBox}>
              <div style={aH2}>① 選擇盤點商品並匯出盤點表</div>
              <div style={{color:"#6b7280",fontSize:13,marginBottom:10}}>輸入國際條碼（每行一個，可多個），留空則匯出全部商品。</div>
              <textarea value={stBarcodes} onChange={e=>setStBarcodes(e.target.value)} placeholder={"4710000000001\n4710000000002\n4710000000003"} rows={4} style={{...fldSt,marginBottom:10,fontFamily:"monospace",fontSize:13}}/>
              <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
                <button onClick={()=>{const barcodeList=stBarcodes.trim().split("\n").map(s=>s.trim()).filter(Boolean);const filtered=barcodeList.length>0?items.filter(b=>barcodeList.includes(String(b.barcode||"").trim())):items;if(filtered.length===0){showToast("找不到符合條碼的商品","error");return;}const r=filtered.map(b=>({條碼:b.barcode||"",商品名稱:b.name,批號:b.batch_no,類別:b.category||"",有效日期:b.expiry_date?b.expiry_date.slice(0,10):"",帳面數量:b.qty,實際數量:"",單位:b.unit||"",儲位:b.location||"",備註:""}));const ws=XLSX.utils.json_to_sheet(r);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"盤點表");XLSX.writeFile(wb,`stocktake_${new Date().toISOString().slice(0,10)}.xlsx`);showToast(`✅ 已匯出 ${filtered.length} 筆盤點表`);}} style={{background:"#2563eb",border:"none",color:"#fff",padding:"8px 18px",borderRadius:6,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:600}}>⬇ 匯出盤點表</button>
                <span style={{color:"#9ca3af",fontSize:12}}>留空 = 匯出全部 {items.length} 筆</span>
              </div>
            </div>
            <div style={aBox}>
              <div style={aH2}>② 匯入盤點結果</div>
              <div style={{color:"#6b7280",fontSize:13,marginBottom:10}}>盤點人員填寫完成後，請匯入已填好的 Excel 檔案。</div>
              <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
                <input placeholder="盤點人員姓名" value={stOperator} onChange={e=>setStOp(e.target.value)} style={{...fldSt,width:160}}/>
                <button onClick={()=>document.getElementById("stFileInput").click()} style={{background:"#16a34a",border:"none",color:"#fff",padding:"8px 18px",borderRadius:6,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:600}}>⬆ 匯入盤點結果</button>
                <input id="stFileInput" type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={async e=>{const file=e.target.files[0];if(!file)return;try{const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellDates:true});const raw=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});const diffs=[];raw.forEach(row=>{const batchNo=String(row["批號"]||"").trim();const actualQty=row["實際數量"];if(!batchNo||actualQty===""||actualQty===null)return;const found=items.find(b=>String(b.batch_no||"").trim()===batchNo);if(!found)return;const actual=Number(actualQty)||0;const diff=actual-found.qty;diffs.push({id:found.id,barcode:found.barcode,name:found.name,batch_no:batchNo,bookQty:found.qty,actualQty:actual,diff,unit:found.unit||""});});setStPending({diffs,operator:stOperator||"未填寫",date:new Date().toLocaleString("zh-TW"),file:file.name});e.target.value="";showToast(`✅ 讀取完成，共 ${diffs.length} 筆`);}catch(err){showToast("匯入失敗："+err.message,"error");}}}/>
              </div>
            </div>
            {stPending&&(
              <div style={aBox}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                  <div style={aH2}>③ 差異確認 — {stPending.operator} · {stPending.date}</div>
                  <button onClick={()=>setStPending(null)} style={{background:"none",border:"none",color:"#9ca3af",cursor:"pointer",fontSize:18}}>✕</button>
                </div>
                <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap"}}>
                  {[{lbl:"盤盈",count:stPending.diffs.filter(d=>d.diff>0).length,col:"#16a34a",bg:"#f0fdf4"},{lbl:"盤虧",count:stPending.diffs.filter(d=>d.diff<0).length,col:"#dc2626",bg:"#fef2f2"},{lbl:"相符",count:stPending.diffs.filter(d=>d.diff===0).length,col:"#2563eb",bg:"#eff6ff"}].map(s=>(
                    <div key={s.lbl} style={{background:s.bg,borderRadius:8,padding:"10px 18px",textAlign:"center"}}>
                      <div style={{color:s.col,fontSize:20,fontWeight:700}}>{s.count}</div>
                      <div style={{color:s.col,fontSize:11,fontWeight:600}}>{s.lbl}</div>
                    </div>
                  ))}
                </div>
                <div style={{overflowX:"auto",marginBottom:14}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead><tr style={{background:"#f8fafc"}}>{["條碼","商品名稱","批號","帳面數量","實際數量","差異","單位","狀態"].map(h=><th key={h} style={{padding:"8px 10px",borderBottom:"1px solid #e5e7eb",textAlign:"left",color:"#6b7280",fontWeight:600,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                    <tbody>{stPending.diffs.map((d,i)=>(<tr key={i} style={{borderBottom:"1px solid #f1f5f9",background:d.diff>0?"#f0fdf4":d.diff<0?"#fef2f2":"#fff"}}><td style={{padding:"7px 10px",color:"#2563eb"}}>{d.barcode||"—"}</td><td style={{padding:"7px 10px"}}>{d.name}</td><td style={{padding:"7px 10px"}}>{d.batch_no}</td><td style={{padding:"7px 10px",textAlign:"right"}}>{d.bookQty}</td><td style={{padding:"7px 10px",textAlign:"right"}}>{d.actualQty}</td><td style={{padding:"7px 10px",textAlign:"right",fontWeight:700,color:d.diff>0?"#16a34a":d.diff<0?"#dc2626":"#6b7280"}}>{d.diff>0?"+":""}{d.diff}</td><td style={{padding:"7px 10px",color:"#6b7280"}}>{d.unit}</td><td style={{padding:"7px 10px"}}><span style={{background:d.diff>0?"#dcfce7":d.diff<0?"#fee2e2":"#f1f5f9",color:d.diff>0?"#16a34a":d.diff<0?"#dc2626":"#6b7280",padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:600}}>{d.diff>0?"盤盈":d.diff<0?"盤虧":"相符"}</span></td></tr>))}</tbody>
                  </table>
                </div>
                <div style={{display:"flex",gap:10}}>
                  <button onClick={()=>setStPending(null)} style={{flex:1,padding:"9px",background:"#f9fafb",border:"1px solid #e5e7eb",color:"#4b5563",borderRadius:6,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:500}}>取消</button>
                  <button onClick={async()=>{for(const d of stPending.diffs){if(d.diff===0)continue;const orig=items.find(b=>b.id===d.id);if(orig)await saveItem({...orig,qty:d.actualQty});}const rec={id:genId(),date:stPending.date,operator:stPending.operator,file:stPending.file,total:stPending.diffs.length,gain:stPending.diffs.filter(d=>d.diff>0).length,loss:stPending.diffs.filter(d=>d.diff<0).length,match:stPending.diffs.filter(d=>d.diff===0).length,diffs:stPending.diffs};setStRecords(prev=>[rec,...prev]);setStPending(null);showToast("✅ 庫存已更新");}} style={{flex:2,padding:"9px",background:"#16a34a",border:"none",color:"#fff",borderRadius:6,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:700}}>✅ 確認並更新庫存</button>
                </div>
              </div>
            )}
            <div style={aBox}>
              <div style={aH2}>📜 歷史盤點紀錄</div>
              {stRecords.length===0&&<div style={{color:"#9ca3af",fontSize:13,textAlign:"center",padding:20}}>尚無盤點紀錄</div>}
              {stRecords.map(rec=>(
                <div key={rec.id} style={{border:"1px solid #e5e7eb",borderRadius:8,padding:14,marginBottom:10,background:"#fafafa"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:8}}>
                    <div>
                      <div style={{fontWeight:600,fontSize:13,color:"#111827",marginBottom:4}}>🗓 {rec.date}　👤 {rec.operator}</div>
                      <div style={{fontSize:12,color:"#6b7280"}}>檔案：{rec.file}　共 {rec.total} 筆</div>
                    </div>
                    <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                      <span style={{background:"#f0fdf4",color:"#16a34a",padding:"2px 10px",borderRadius:12,fontSize:11,fontWeight:600}}>盈 {rec.gain}</span>
                      <span style={{background:"#fef2f2",color:"#dc2626",padding:"2px 10px",borderRadius:12,fontSize:11,fontWeight:600}}>虧 {rec.loss}</span>
                      <span style={{background:"#eff6ff",color:"#2563eb",padding:"2px 10px",borderRadius:12,fontSize:11,fontWeight:600}}>符 {rec.match}</span>
                      <button onClick={()=>{const printWin=window.open("","_blank","width=900,height=700");const rows=rec.diffs.map(d=>`<tr style="background:${d.diff>0?"#f0fdf4":d.diff<0?"#fef2f2":"#fff"}"><td>${d.barcode||"—"}</td><td>${d.name}</td><td>${d.batch_no}</td><td style="text-align:right">${d.bookQty}</td><td style="text-align:right">${d.actualQty}</td><td style="text-align:right;font-weight:700;color:${d.diff>0?"#16a34a":d.diff<0?"#dc2626":"#6b7280"}">${d.diff>0?"+":""}${d.diff}</td><td>${d.unit}</td><td>${d.diff>0?"盤盈":d.diff<0?"盤虧":"相符"}</td></tr>`).join("");printWin.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>盤點報告</title><style>body{font-family:Arial,sans-serif;padding:24px;font-size:13px;}table{width:100%;border-collapse:collapse;}th,td{border:1px solid #e5e7eb;padding:7px 10px;text-align:left;}th{background:#f8fafc;}@media print{button{display:none}}</style></head><body><h2>📋 庫存盤點報告</h2><p>盤點日期：${rec.date}　盤點人員：${rec.operator}</p><p>盤盈：${rec.gain} 筆　盤虧：${rec.loss} 筆　相符：${rec.match} 筆</p><table><thead><tr><th>條碼</th><th>商品名稱</th><th>批號</th><th>帳面數量</th><th>實際數量</th><th>差異</th><th>單位</th><th>狀態</th></tr></thead><tbody>${rows}</tbody></table><br><button onclick="window.print()">🖨 列印</button></body></html>`);printWin.document.close();}} style={{background:"#f3f4f6",border:"1px solid #d1d5db",color:"#374151",padding:"4px 12px",borderRadius:5,cursor:"pointer",fontSize:11,fontFamily:"inherit",fontWeight:600}}>🖨 列印</button>
                      <button onClick={()=>{const r=rec.diffs.map(d=>({條碼:d.barcode||"",商品名稱:d.name,批號:d.batch_no,帳面數量:d.bookQty,實際數量:d.actualQty,差異:d.diff,單位:d.unit,狀態:d.diff>0?"盤盈":d.diff<0?"盤虧":"相符"}));const ws=XLSX.utils.json_to_sheet(r);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"盤點結果");XLSX.writeFile(wb,`stocktake_result_${rec.date.slice(0,10)}.xlsx`);showToast("✅ 匯出完成");}} style={{background:"#eff6ff",border:"1px solid #bfdbfe",color:"#2563eb",padding:"4px 12px",borderRadius:5,cursor:"pointer",fontSize:11,fontFamily:"inherit",fontWeight:600}}>⬇ 匯出</button>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ══ IMPORT ══ */}
      {tab==="import"&&(
        <div style={{flex:1,overflow:"auto",padding:20}} className="fade">
          <div style={{maxWidth:580}}>
            <div style={aH1}>⬆ 批次匯入 Excel</div>
            <div style={aBox}>
              <div style={aH2}>步驟 1 — 下載範本</div>
              <button onClick={downloadTemplate} style={{background:"#e5e7eb",color:"#374151",border:"none",padding:"9px 20px",borderRadius:999,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:600,transition:"all 0.18s"}} onMouseEnter={e=>{e.currentTarget.style.background="#1d4ed8";e.currentTarget.style.color="#fff";e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 4px 12px rgba(37,99,235,0.35)";}} onMouseLeave={e=>{e.currentTarget.style.background="#e5e7eb";e.currentTarget.style.color="#374151";e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="none";}}>📥 下載 Excel 範本</button>
              <div style={{color:"#9ca3af",fontSize:12,marginTop:8}}>欄位：條碼、商品名稱、批號、有效日期、類別、庫存量、單位、成本價、售價、儲位、供應商、備註</div>
            </div>
            <div style={{...aBox,border:"2px dashed #d1d5db",textAlign:"center",cursor:"pointer",padding:32}} onClick={()=>fileRef.current?.click()}>
              <div style={{fontSize:36,marginBottom:8}}>📂</div>
              <div style={{color:"#374151",fontSize:14,fontWeight:600,marginBottom:4}}>點擊選擇檔案</div>
              <div style={{color:"#9ca3af",fontSize:12}}>支援 .xlsx  .xls  .csv</div>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleImport}/>
            </div>
            {importLog&&<div style={{background:"#f0fdf4",border:"1px solid #86efac",borderRadius:8,padding:14}}>
              <div style={{color:"#16a34a",fontSize:13,fontWeight:700}}>✅ 匯入完成</div>
              <div style={{color:"#6b7280",fontSize:12,marginTop:4}}>新增 {importLog.added} 筆，更新 {importLog.updated} 筆，共 {importLog.ok} / {importLog.total} 筆　檔案：{importLog.file}</div>
            </div>}
          </div>
        </div>
      )}

      {/* ══ LINE ══ */}
      {tab==="line"&&(
        <div style={{flex:1,overflow:"auto",padding:20}} className="fade">
          <div style={{maxWidth:580}}>
            <div style={{...aH1,color:"#16a34a"}}>💬 LINE 推播</div>
            <div style={aBox}>
              <div style={aH2}>LINE NOTIFY TOKEN（選填）</div>
              <input value={lineToken} onChange={e=>{setLineToken(e.target.value);idbSet("lineToken",e.target.value);}} placeholder="貼入 LINE Notify Token…" style={fldSt}/>
              <div style={{color:"#9ca3af",fontSize:11,marginTop:6}}>前往 notify-bot.line.me → 登入 → 個人頁面 → 發行存取權杖</div>
            </div>
            <button onClick={genLineMsg} style={{background:"#f0fdf4",border:"1px solid #86efac",color:"#16a34a",padding:"9px 18px",borderRadius:6,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:600,marginBottom:10}}>🔄 產生即期清單訊息</button>
            {lineMsg&&<>
              <textarea value={lineMsg} onChange={e=>setLineMsg(e.target.value)} rows={10} style={{...fldSt,marginBottom:10,lineHeight:1.6}}/>
              <div style={{display:"flex",gap:10}}>
                <button onClick={()=>{navigator.clipboard?.writeText(lineMsg);showToast("已複製 ✓");}} style={{flex:1,background:"#eff6ff",border:"1px solid #bfdbfe",color:"#2563eb",padding:"11px",borderRadius:7,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:600}}>📋 複製訊息</button>
                <button onClick={()=>window.open("https://line.me","_blank")} style={{flex:1,background:"#06c755",border:"none",color:"#fff",padding:"11px",borderRadius:7,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:600}}>↗ 開啟 LINE</button>
              </div>
            </>}
          </div>
        </div>
      )}

      {/* ══ ADMIN ══ */}
      {tab==="admin"&&(
        <div style={{flex:1,overflow:"hidden",display:"flex"}}>
          {/* Side nav */}
          <div style={{width:180,background:"#fff",borderRight:"1px solid #e5e7eb",padding:"12px 0",flexShrink:0,overflowY:"auto"}}>
            <div style={{padding:"0 14px 10px",fontSize:10,color:"#9ca3af",letterSpacing:1,fontWeight:600,textTransform:"uppercase"}}>後台管理</div>
            {[
              {id:"overview",icon:"📋",lbl:"系統總覽"},
              {id:"supabase",icon:"🔌",lbl:"Supabase 設定"},
              {id:"schema",  icon:"🗄️", lbl:"資料庫結構"},
              {id:"rawdata", icon:"📦",lbl:"原始資料"},
              {id:"export",  icon:"⬇️", lbl:"資料匯出"},
              {id:"autobot", icon:"⚡",lbl:"自動推播說明"},
              {id:"devguide",icon:"📖",lbl:"開發手冊"},
            ].map(t=>(
              <button key={t.id} onClick={()=>setAdminTab(t.id)} style={{display:"flex",alignItems:"center",gap:8,width:"100%",textAlign:"left",background:adminTab===t.id?"#eff6ff":"transparent",border:"none",borderLeft:adminTab===t.id?"3px solid #2563eb":"3px solid transparent",color:adminTab===t.id?"#2563eb":"#4b5563",padding:"9px 14px",cursor:"pointer",fontSize:12,fontFamily:"inherit",fontWeight:adminTab===t.id?600:400}}>
                <span>{t.icon}</span><span>{t.lbl}</span>
              </button>
            ))}
          </div>

          <div style={{flex:1,overflow:"auto",padding:20,background:"#f9fafb"}} className="fade">

            {adminTab==="overview"&&(
              <div style={{maxWidth:680}}>
                <div style={aH1}>📋 系統總覽</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12,marginBottom:16}}>
                  {[{lbl:"資料來源",val:isOnline?"Supabase 雲端":"IndexedDB 本機",col:isOnline?"#16a34a":"#ea580c"},{lbl:"批號總數",val:items.length,col:"#2563eb"},{lbl:"風險批號",val:stats.expired+stats.red,col:"#dc2626"},{lbl:"庫存總值",val:fmtMoney(stats.totalVal),col:"#16a34a"}].map((s,i)=>(
                    <div key={i} style={aBox}><div style={{color:s.col,fontSize:20,fontWeight:700}}>{s.val}</div><div style={{color:"#9ca3af",fontSize:11,marginTop:3}}>{s.lbl}</div></div>
                  ))}
                </div>
                <div style={aBox}><div style={aH2}>架構說明</div><div style={{color:"#4b5563",fontSize:13,lineHeight:1.8}}><strong style={{color:"#16a34a"}}>Layer 1 — IndexedDB（瀏覽器本機）</strong><br/>離線可用，清除瀏覽器資料會遺失。<br/><br/><strong style={{color:"#2563eb"}}>Layer 2 — Supabase（雲端 PostgreSQL）</strong><br/>設定 URL 和 Key 後啟用，多設備同步永久儲存。<br/><br/>儲存順序：先寫 IndexedDB → 再同步 Supabase。</div></div>
              </div>
            )}

            {adminTab==="supabase"&&(
              <div style={{maxWidth:640}}>
                <div style={aH1}>🔌 Supabase 設定</div>
                <div style={{...aBox,borderColor:isOnline?"#86efac":"#fdba74",background:isOnline?"#f0fdf4":"#fff7ed"}}>
                  <div style={{color:isOnline?"#16a34a":"#ea580c",fontSize:14,fontWeight:700}}>{isOnline?"✅ 已連線，雲端模式啟用":"⚠️ 未連線，目前為本機模式"}</div>
                </div>
                <div style={aBox}><div style={aH2}>設定步驟</div>
                  {[{n:"1",t:"建立專案",d:"前往 supabase.com → New Project → 名稱 + 密碼 → Singapore → Create"},{n:"2",t:"取得金鑰",d:"Settings → API → 複製 Project URL 和 anon/public key"},{n:"3",t:"修改程式碼（只改兩行）",d:"SUPABASE_URL = \"你的 Project URL\"\nSUPABASE_KEY = \"你的 anon key\""},{n:"4",t:"建立資料表",d:"到「資料庫結構」頁複製 SQL → Supabase SQL Editor 貼上執行"},{n:"5",t:"完成",d:"重新整理頁面，頂部顯示綠色「雲端連線中」"}].map(s=>(
                    <div key={s.n} style={{display:"flex",gap:12,marginBottom:14}}>
                      <div style={{width:24,height:24,borderRadius:12,background:"#2563eb",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,flexShrink:0}}>{s.n}</div>
                      <div><div style={{color:"#111827",fontSize:13,fontWeight:600,marginBottom:2}}>{s.t}</div><div style={{color:"#6b7280",fontSize:12,lineHeight:1.6,whiteSpace:"pre-wrap"}}>{s.d}</div></div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {adminTab==="schema"&&(
              <div style={{maxWidth:700}}>
                <div style={aH1}>🗄️ 資料庫結構</div>
                <div style={aBox}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                    <div style={aH2}>建立資料表 SQL</div>
                    <button onClick={()=>{navigator.clipboard?.writeText(SCHEMA);showToast("已複製 SQL");}} style={cpyBtn}>📋 複製</button>
                  </div>
                  <pre style={{...codeS,maxHeight:340,overflow:"auto"}}>{SCHEMA}</pre>
                </div>
                <div style={aBox}><div style={aH2}>欄位說明</div>
                  {[["id","UUID","主鍵，自動產生"],["barcode","TEXT","國際條碼"],
                    ["product_no","TEXT","產品編號"],["name","TEXT","商品名稱（必填）"],["batch_no","TEXT","批號，FIFO追蹤用"],["expiry_date","DATE","有效日期"],["qty","INTEGER","庫存數量"],["cost","NUMERIC","成本價"],["price","NUMERIC","售價"],["location","TEXT","儲位"],["created_at","TIMESTAMPTZ","建立時間（自動）"]].map(([c,t,d])=>(
                    <div key={c} style={{display:"flex",gap:12,padding:"6px 0",borderBottom:"1px solid #f3f4f6",fontSize:12}}>
                      <span style={{color:"#2563eb",width:110,flexShrink:0,fontWeight:600}}>{c}</span>
                      <span style={{color:"#16a34a",width:80,flexShrink:0}}>{t}</span>
                      <span style={{color:"#6b7280"}}>{d}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {adminTab==="rawdata"&&(
              <div style={{maxWidth:860}}>
                <div style={aH1}>📦 原始資料瀏覽</div>
                <div style={{color:"#6b7280",fontSize:12,marginBottom:12}}>共 {items.length} 筆 · 來源：{isOnline?"Supabase 雲端":"IndexedDB 本機"}</div>
                <div style={{...aBox,padding:0,overflow:"hidden"}}>
                  <div style={{overflowX:"auto"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                      <thead><tr style={{background:"#f8fafc",borderBottom:"2px solid #e2e8f0"}}>{["id","barcode","商品名稱","批號","有效日期","庫存","成本","狀態"].map(h=><th key={h} style={{padding:"10px 12px",textAlign:"left",color:"#6b7280",fontWeight:600,fontSize:11,letterSpacing:0.5,textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                      <tbody>
                        {items.slice(0,60).map((b,i)=>{ const t=tierOf(daysLeft(b.expiry_date)); return (
                          <tr key={b.id} style={{borderBottom:"1px solid #f3f4f6",background:i%2===0?"#fff":"#fafafa"}}>
                            <td style={{padding:"8px 12px",color:"#9ca3af",fontFamily:"monospace",fontSize:11}}>{b.id.slice(0,8)}…</td>
                            <td style={{padding:"8px 12px",color:"#2563eb",fontWeight:500}}>{b.barcode||"—"}</td>
                            <td style={{padding:"8px 12px",color:"#111827",fontWeight:500,maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{b.name}</td>
                            <td style={{padding:"8px 12px",color:"#6b7280"}}>{b.batch_no||"—"}</td>
                            <td style={{padding:"8px 12px",color:"#6b7280",whiteSpace:"nowrap"}}>{fmtDate(b.expiry_date)||"—"}</td>
                            <td style={{padding:"8px 12px",color:"#16a34a",fontWeight:500}}>{b.qty}{b.unit}</td>
                            <td style={{padding:"8px 12px",color:"#16a34a",fontWeight:500}}>{fmtMoney(b.cost)}</td>
                            <td style={{padding:"8px 12px"}}><span style={{background:TIER[t].bg,color:TIER[t].color,border:`1px solid ${TIER[t].border}`,padding:"2px 7px",borderRadius:4,fontSize:10,fontWeight:600}}>{TIER[t].label}</span></td>
                          </tr>
                        );})}
                      </tbody>
                    </table>
                    {items.length>60&&<div style={{padding:"8px 14px",color:"#9ca3af",fontSize:11}}>顯示前 60 筆，共 {items.length} 筆</div>}
                    {items.length===0&&<div style={{padding:40,textAlign:"center",color:"#9ca3af"}}>尚無資料</div>}
                  </div>
                </div>
              </div>
            )}

            {adminTab==="export"&&(
              <div style={{maxWidth:580}}>
                <div style={aH1}>⬇️ 資料匯出</div>
                {[
                  {lbl:"全部庫存 Excel",desc:"所有批號完整資料",act:()=>{ const r=items.map(b=>({條碼:b.barcode,商品名稱:b.name,批號:b.batch_no,有效日期:fmtDate(b.expiry_date),類別:b.category,庫存量:b.qty,單位:b.unit,成本:b.cost,售價:b.price,儲位:b.location,供應商:b.supplier,備註:b.note,狀態:TIER[tierOf(daysLeft(b.expiry_date))].label}));const ws=XLSX.utils.json_to_sheet(r);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"庫存");XLSX.writeFile(wb,`inventory_${new Date().toISOString().slice(0,10)}.xlsx`);showToast("匯出完成 ✅"); }},
                  {lbl:"即期警示清單",desc:"只匯出 90 天內到期批號",act:()=>{ const r=items.filter(b=>{const d=daysLeft(b.expiry_date);return d!==null&&d<90;}).map(b=>({條碼:b.barcode,商品名稱:b.name,批號:b.batch_no,有效日期:fmtDate(b.expiry_date),剩餘天數:daysLeft(b.expiry_date),狀態:TIER[tierOf(daysLeft(b.expiry_date))].label,庫存量:b.qty,風險金額:(b.qty||0)*(b.cost||0)}));const ws=XLSX.utils.json_to_sheet(r);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"即期清單");XLSX.writeFile(wb,`expiry_${new Date().toISOString().slice(0,10)}.xlsx`);showToast("匯出完成 ✅"); }},
                  {lbl:"JSON 備份",desc:"完整原始資料，可用於還原",act:()=>{ const blob=new Blob([JSON.stringify(items,null,2)],{type:"application/json"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`backup_${new Date().toISOString().slice(0,10)}.json`;a.click();showToast("備份完成 ✅"); }},
                ].map((btn,i)=>(
                  <div key={i} style={{...aBox,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                    <div><div style={{color:"#111827",fontSize:13,fontWeight:600}}>{btn.lbl}</div><div style={{color:"#9ca3af",fontSize:12,marginTop:2}}>{btn.desc}</div></div>
                    <button onClick={btn.act} style={{background:"#e5e7eb",color:"#374151",border:"none",padding:"8px 18px",borderRadius:999,cursor:"pointer",fontSize:12,fontFamily:"inherit",fontWeight:600,flexShrink:0,marginLeft:14,transition:"all 0.18s"}} onMouseEnter={e=>{e.currentTarget.style.background="#1d4ed8";e.currentTarget.style.color="#fff";e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 4px 12px rgba(37,99,235,0.35)";}} onMouseLeave={e=>{e.currentTarget.style.background="#e5e7eb";e.currentTarget.style.color="#374151";e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="none";}}>匯出</button>
                  </div>
                ))}
              </div>
            )}

            {adminTab==="autobot"&&(
              <div style={{maxWidth:640}}>
                <div style={aH1}>⚡ 自動推播設定說明</div>
                <div style={aBox}><div style={aH2}>設定步驟概覽</div><pre style={codeS}>{AUTO_GUIDE}</pre></div>
                <div style={aBox}><div style={aH2}>需要的環境變數</div>
                  {[{k:"SUPABASE_URL",v:"你的 Supabase Project URL"},{k:"SUPABASE_SERVICE_ROLE_KEY",v:"Settings → API → service_role key"},{k:"LINE_NOTIFY_TOKEN",v:"LINE Notify 個人頁面發行的 Token"}].map(r=>(
                    <div key={r.k} style={{padding:"8px 0",borderBottom:"1px solid #f3f4f6",fontSize:12}}>
                      <div style={{color:"#2563eb",fontWeight:600,fontFamily:"monospace"}}>{r.k}</div>
                      <div style={{color:"#6b7280",marginTop:2}}>{r.v}</div>
                    </div>
                  ))}
                </div>
                <div style={{...aBox,background:"#fefce8",borderColor:"#fde047"}}><div style={{color:"#854d0e",fontSize:13}}>💡 不想自己架伺服器？可以用 <strong>Make.com</strong> 或 <strong>Zapier</strong> 無程式碼串接 LINE Notify，不需要安裝任何東西。</div></div>
              </div>
            )}

            {adminTab==="devguide"&&(
              <div style={{maxWidth:680}}>
                <div style={aH1}>📖 初級工程師開發手冊</div>
                {[
                  {q:"Q: 如何新增欄位？",a:"1. Supabase SQL Editor 執行：\n   ALTER TABLE inventory_batches ADD COLUMN new_field TEXT;\n\n2. 在程式 COLS 陣列新增：\n   {k:\"new_field\", lbl:\"新欄位\", w:100, ed:true}\n\n3. 在 newBatch() 加入預設值：\n   new_field: \"\""},
                  {q:"Q: 如何修改三色分級天數？",a:"找到 tierOf() 函數：\n\nfunction tierOf(days) {\n  if (days < 0)    return \"expired\";\n  if (days < 30)   return \"red\";    ← 改這裡\n  if (days < 90)   return \"yellow\"; ← 改這裡\n  if (days <= 180) return \"green\";  ← 改這裡\n  return \"safe\";\n}"},
                  {q:"Q: 如何讓其他人使用同一份庫存？",a:"設定 Supabase 後資料在雲端，\n其他人打開同一個網址就能看到相同庫存。\n\n注意：目前無登入機制。\n需要權限管理請在 Supabase 設定 RLS。"},
                  {q:"Q: 如何查看錯誤？",a:"按 F12 → Console 分頁\n紅色文字就是錯誤訊息\n\n401 → API Key 錯誤\n404 → 資料表不存在，請執行 Schema SQL\n400 → 資料格式錯誤"},
                ].map((item,i)=>(
                  <div key={i} style={{...aBox,marginBottom:12}}>
                    <div style={{color:"#2563eb",fontSize:13,fontWeight:700,marginBottom:10}}>{item.q}</div>
                    <pre style={{...codeS,fontSize:11}}>{item.a}</pre>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* ADD FORM */}
      {showForm&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={e=>e.target===e.currentTarget&&setShowForm(false)}>
          <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:24,width:"100%",maxWidth:540,maxHeight:"90vh",overflow:"auto",boxShadow:"0 20px 60px rgba(0,0,0,0.15)"}}>
            <div style={{fontSize:16,fontWeight:700,color:"#111827",marginBottom:16}}>＋ 新增批號入庫</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              {[{k:"barcode",l:"國際條碼"},{k:"product_no",l:"產品編號"},{k:"name",l:"商品名稱 *"},{k:"expiry_date",l:"有效日期",t:"date"},{k:"category",l:"類別",t:"sel"},{k:"qty",l:"庫存量",t:"number"},{k:"unit",l:"單位"},{k:"cost",l:"成本價",t:"number"},{k:"price",l:"售價",t:"number"},{k:"location",l:"儲位"},{k:"supplier",l:"供應商"}].map(f=>(
                <div key={f.k}>
                  <label style={{fontSize:11,color:"#6b7280",display:"block",marginBottom:4,fontWeight:500}}>{f.l}</label>
                  {f.t==="sel"?<select value={form[f.k]||""} onChange={e=>setForm(p=>({...p,[f.k]:e.target.value}))} style={fldSt}>{dynCats.map(c=><option key={c}>{c}</option>)}</select>:<input type={f.t||"text"} value={form[f.k]||""} onChange={e=>setForm(p=>({...p,[f.k]:e.target.value}))} style={fldSt}/>}
                </div>
              ))}
            </div>
            <div style={{marginTop:10}}>
              <label style={{fontSize:11,color:"#6b7280",display:"block",marginBottom:4,fontWeight:500}}>備註</label>
              <input value={form.note||""} onChange={e=>setForm(p=>({...p,note:e.target.value}))} style={{...fldSt,width:"100%"}} placeholder="選填"/>
            </div>
            <div style={{display:"flex",gap:10,marginTop:18}}>
              <button onClick={()=>setShowForm(false)} style={{flex:1,padding:"11px",background:"#e5e7eb",color:"#374151",border:"none",borderRadius:999,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:500,transition:"all 0.18s"}} onMouseEnter={e=>{e.currentTarget.style.background="#1d4ed8";e.currentTarget.style.color="#fff";e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 4px 16px rgba(37,99,235,0.35)";}} onMouseLeave={e=>{e.currentTarget.style.background="#e5e7eb";e.currentTarget.style.color="#374151";e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="none";}}>取消</button>
              <button onClick={async()=>{ if(!form.name?.trim())return showToast("請填商品名稱","error");await saveItem({...form,id:form.id||genId()});setShowForm(false);showToast("✅ 入庫完成"); }} style={{flex:2,padding:"11px",background:"#e5e7eb",color:"#374151",border:"none",borderRadius:999,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:700,transition:"all 0.18s"}} onMouseEnter={e=>{e.currentTarget.style.background="#1d4ed8";e.currentTarget.style.color="#fff";e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 4px 16px rgba(37,99,235,0.35)";}} onMouseLeave={e=>{e.currentTarget.style.background="#e5e7eb";e.currentTarget.style.color="#374151";e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="none";}}>確認入庫</button>
            </div>
          </div>
        </div>
      )}

      {/* IMPORT LOADING OVERLAY */}
      {importing&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:500,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"#fff",borderRadius:16,padding:"36px 48px",textAlign:"center",boxShadow:"0 20px 60px rgba(0,0,0,0.2)",minWidth:300}}>
            <div style={{width:56,height:56,border:"5px solid #e5e7eb",borderTop:"5px solid #2563eb",borderRadius:"50%",margin:"0 auto 20px",animation:"spin 0.9s linear infinite"}}/>
            <div style={{fontSize:16,fontWeight:700,color:"#111827",marginBottom:8}}>資料匯入中，請稍候…</div>
            <div style={{fontSize:13,color:"#6b7280",marginBottom:16}}>{importProgress.current} / {importProgress.total} 筆</div>
            <div style={{background:"#f3f4f6",borderRadius:999,height:10,overflow:"hidden"}}>
              <div style={{height:"100%",background:"linear-gradient(90deg,#2563eb,#7c3aed)",borderRadius:999,width:`${importProgress.total?Math.round(importProgress.current/importProgress.total*100):0}%`,transition:"width 0.2s"}}/>
            </div>
            <div style={{fontSize:12,color:"#9ca3af",marginTop:8}}>{importProgress.total?Math.round(importProgress.current/importProgress.total*100):0}%</div>
          </div>
        </div>
      )}

      {/* TOAST */}
      {toast&&<div style={{position:"fixed",bottom:20,right:20,zIndex:9999,background:toast.type==="ok"?"#f0fdf4":toast.type==="warning"?"#fff7ed":"#fef2f2",border:`1px solid ${toast.type==="ok"?"#86efac":toast.type==="warning"?"#fdba74":"#fca5a5"}`,color:toast.type==="ok"?"#16a34a":toast.type==="warning"?"#ea580c":"#dc2626",padding:"12px 18px",borderRadius:8,fontSize:13,fontWeight:600,boxShadow:"0 4px 20px rgba(0,0,0,0.12)",fontFamily:"inherit",maxWidth:360}}>{toast.msg}</div>}
    </div>
  );
}
