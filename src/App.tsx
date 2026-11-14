import {memo,useCallback,useEffect,useMemo,useRef,useState} from 'react';
import {useLiveQuery} from 'dexie-react-hooks';
import toast,{Toaster} from 'react-hot-toast';
import {motion,AnimatePresence,animate} from 'framer-motion';
import type {WorkBook} from 'xlsx';
import {Shield,Upload,Copy,Check,Search,X,Trash2,Download,Settings as Cog,LayoutDashboard,List as ListI,History as Hist,LogOut,Menu,Sun,Moon,ChevronLeft,Inbox,CheckCircle2} from 'lucide-react';
import {db,Account,runImport,parseBackup,restoreAccounts} from './db';
import * as U from './utils';
import {card,btn,primary,ghost,inp} from './ui';
import {Login,Logo} from './auth';

type Page='dash'|'accounts'|'import'|'history'|'settings';
type Ask=(t:string,m:string,y:()=>void)=>void;
const mut='text-[var(--mut)]';

function Dlg({c,close}:{c:{t:string;m:string;y:()=>void};close:()=>void}){
  const r=useRef<HTMLButtonElement>(null);
  useEffect(()=>{r.current?.focus();const k=(e:KeyboardEvent)=>e.key==='Escape'&&close();addEventListener('keydown',k);return()=>removeEventListener('keydown',k);},[close]);
  return(<div className="fixed inset-0 z-50 grid place-items-center bg-black/60 p-4" onClick={close}>
    <motion.div role="dialog" aria-modal="true" aria-label={c.t} initial={{scale:.95,opacity:0}} animate={{scale:1,opacity:1}} className={`${card} w-full max-w-sm space-y-4`} onClick={e=>e.stopPropagation()}>
      <h2 className="font-semibold">{c.t}</h2><p className={`text-sm ${mut}`}>{c.m}</p>
      <div className="flex justify-end gap-2"><button className={ghost} onClick={close}>Cancel</button><button ref={r} className={`${btn} bg-[#EF4444] text-white`} onClick={()=>{c.y();close();}}>Confirm</button></div></motion.div></div>);
}
function Num({v}:{v:number}){const [n,setN]=useState(0);useEffect(()=>{const a=animate(0,v,{duration:.8,onUpdate:x=>setN(Math.round(x))});return()=>a.stop();},[v]);return <>{n}</>;}
function Donut({used,total}:{used:number;total:number}){
  const r=52,c=2*Math.PI*r,p=total?used/total:0;
  return(<svg viewBox="0 0 128 128" className="w-36 h-36 -rotate-90" role="img" aria-label={`${used} of ${total} accounts used`}>
    <circle cx="64" cy="64" r={r} fill="none" stroke="var(--bd)" strokeWidth="14"/>
    <circle cx="64" cy="64" r={r} fill="none" stroke="#22C55E" strokeWidth="14" strokeLinecap="round" strokeDasharray={`${c*p} ${c}`} style={{transition:'stroke-dasharray .6s'}}/></svg>);
}

function Dashboard({all,go}:{all:Account[];go:(p:Page)=>void}){
  const used=all.filter(a=>a.status==='used').length,unused=all.length-used,pct=all.length?Math.round(used/all.length*100):0;
  const logs=useLiveQuery(()=>db.imports.orderBy('date').reverse().limit(4).toArray(),[])||[];
  const act=useMemo(()=>[...all].sort((a,b)=>+new Date(b.updatedAt)-+new Date(a.updatedAt)).slice(0,5),[all]);
  const stats:[string,number,string,string][]=[['TOTAL ACCOUNTS',all.length,'',''],['USED ACCOUNTS',used,'text-[#22C55E]',''],['UNUSED ACCOUNTS',unused,'text-[#F59E0B]',''],['USAGE',pct,'text-[#8B5CF6]','%']];
  return(<div className="space-y-4">
    <div className="grid grid-cols-2 xl:grid-cols-4 gap-4">{stats.map(([t,v,c,s])=><div key={t} className={card}><div className={`text-xs ${mut}`}>{t}</div><div className={`text-3xl font-bold mt-1 ${c}`}><Num v={v}/>{s}</div></div>)}</div>
    <div className="grid lg:grid-cols-3 gap-4">
      <div className={`${card} flex items-center gap-5`}><div className="relative"><Donut used={used} total={all.length}/><div className="absolute inset-0 grid place-items-center font-bold">{pct}%</div></div>
        <div className="text-sm space-y-1"><div><span className="inline-block size-2 rounded-full bg-[#22C55E] mr-2"/>Used {used}</div><div><span className="inline-block size-2 rounded-full bg-[var(--bd)] mr-2"/>Unused {unused}</div></div></div>
      <div className={card}><div className="font-semibold mb-2">Recent imports</div>{!logs.length&&<div className={`text-sm ${mut}`}>No imports yet.</div>}
        {logs.map(l=><div key={l.id} className="text-sm flex justify-between gap-2 py-1"><span className="truncate">{l.fileName}</span><span className={`${mut} shrink-0`}>+{l.imported}</span></div>)}</div>
      <div className={card}><div className="font-semibold mb-2">Account activity</div>{!act.length&&<div className={`text-sm ${mut}`}>No accounts yet.</div>}
        {act.map(a=><div key={a.id} className="text-sm flex justify-between gap-2 py-1"><span className="truncate">{a.email}</span><span className={a.status==='used'?'text-[#22C55E]':mut}>{a.status}</span></div>)}</div></div>
    <div className="flex flex-wrap gap-2"><button className={primary} onClick={()=>go('import')}><Upload className="inline mr-1" size={14}/>Quick upload</button><button className={ghost} onClick={()=>go('accounts')}>Browse accounts</button></div></div>);
}

function ImportPage({all}:{all:Account[]}){
  const ref=useRef<HTMLInputElement>(null),wbRef=useRef<WorkBook|null>(null);
  const [file,setFile]=useState<File|null>(null),[rows,setRows]=useState<any[][]>([]),[sheets,setSheets]=useState<string[]>([]),[sheet,setSheet]=useState(''),[col,setCol]=useState(0),[busy,setBusy]=useState(false),[prog,setProg]=useState(0),[done,setDone]=useState<{added:number;dup:number;inv:number}|null>(null),[drag,setDrag]=useState(false);
  const have=useMemo(()=>new Set(all.map(a=>a.normalizedEmail)),[all]);
  const clear=()=>{setFile(null);setRows([]);setSheets([]);setSheet('');setCol(0);setProg(0);wbRef.current=null;if(ref.current)ref.current.value='';};
  const load=(name:string)=>{const r=U.sheetRows(wbRef.current!,name);setSheet(name);setRows(r);setCol(U.detectCol(r));if(!r.length)toast.error('This sheet is empty');};
  const pick=async(f?:File)=>{
    if(!f)return;setDone(null);setBusy(true);
    try{const wb=await U.readWB(f);wbRef.current=wb;setSheets(wb.SheetNames);setFile(f);load(wb.SheetNames[0]);}
    catch(e){toast.error((e as Error).message==='type'?'Use an XLSX, XLS or CSV file':'This file is empty or could not be read');clear();}
    setBusy(false);if(ref.current)ref.current.value='';
  };
  const prev=useMemo(()=>U.extract(rows,col,have),[rows,col,have]);
  const ncols=Math.max(1,...rows.slice(0,50).map(r=>r.length));
  const commit=async()=>{
    if(busy)return;setBusy(true);setProg(15);
    try{
      const r=await runImport(file!.name,prev.valid,prev.total,prev.inv);
      setProg(100);setDone({added:r.added,dup:r.dup,inv:prev.inv});clear();
      r.added?toast.success('Import successful. Temporary Excel file cleared.'):toast('No new emails — every email was already saved. Temporary file cleared.',{icon:'ℹ️'});
    }catch{toast.error('Import failed. Nothing was changed — you can retry.');}
    setBusy(false);
  };
  return(<div className="space-y-4">
    <div className={card} onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);pick(e.dataTransfer.files[0]);}}>
      <div className={`border-2 border-dashed rounded-2xl p-8 text-center transition ${drag?'border-[#6366F1] bg-[#6366F1]/10':'border-[var(--bd)]'}`}>
        <Upload className="mx-auto mb-2 text-[#8B5CF6]" size={32}/><div className="font-medium">Drag & drop your spreadsheet</div>
        <div className={`text-xs ${mut} mb-4`}>XLSX, XLS or CSV · read locally in your browser, never uploaded</div>
        <input ref={ref} type="file" accept=".xlsx,.xls,.csv" hidden aria-label="Choose spreadsheet" onChange={e=>pick(e.target.files?.[0])}/>
        <button className={primary} disabled={busy} onClick={()=>ref.current?.click()}>Browse Files</button></div>
      {busy&&<div className="mt-4" role="status"><div className="h-2 rounded-full bg-[var(--bd)] overflow-hidden"><div className="h-full bg-[#6366F1] transition-all" style={{width:`${prog||30}%`}}/></div><div className={`text-xs mt-1 ${mut}`}>Working…</div></div>}
      <AnimatePresence>{done&&<motion.div initial={{opacity:0,y:8}} animate={{opacity:1,y:0}} className="mt-4 text-center space-y-1"><motion.div initial={{scale:0}} animate={{scale:1}} className="inline-block"><CheckCircle2 size={44} className="text-[#22C55E]"/></motion.div>
        <div className="font-semibold">Import successful. Temporary Excel file cleared.</div><div className={`text-sm ${mut}`}>{done.added} added · {done.dup} duplicates · {done.inv} invalid</div>
        <div className={`text-xs ${mut}`}>Your original file on your device is untouched.</div><button className={ghost} onClick={()=>{setDone(null);ref.current?.click();}}>Upload another file</button></motion.div>}</AnimatePresence>
    </div>
    {file&&<div className={`${card} space-y-4`}><div className="flex flex-wrap items-center gap-3 text-sm"><b className="break-all">{file.name}</b><span className={mut}>{(file.size/1024).toFixed(1)} KB</span>
      {sheets.length>1&&<label className="flex items-center gap-2">Sheet<select className={`${inp} !w-auto`} value={sheet} onChange={e=>load(e.target.value)}>{sheets.map(s=><option key={s}>{s}</option>)}</select></label>}
      <label className="flex items-center gap-2">Email column<select className={`${inp} !w-auto`} value={col} onChange={e=>setCol(+e.target.value)}>{Array.from({length:ncols},(_,i)=><option key={i} value={i}>Column {i+1}{rows[0]?.[i]?` · ${String(rows[0][i]).slice(0,18)}`:''}</option>)}</select></label>
      <button className={ghost} onClick={clear}><X className="inline" size={14}/> Remove file</button></div>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-center text-sm">{([['Emails found',prev.total,''],['New emails',prev.valid.length,'text-[#22C55E]'],['Duplicates',prev.dup,'text-[#F59E0B]'],['Invalid',prev.inv,'text-[#EF4444]']] as const).map(([t,v,c])=><div key={t} className="rounded-xl bg-[var(--bd)] p-3"><b className={`text-xl ${c}`}>{v}</b><div className={mut}>{t}</div></div>)}</div>
      {prev.header&&<div className={`text-xs ${mut}`}>First row detected as a header and skipped.</div>}
      <ul className="text-sm space-y-1">{prev.valid.slice(0,10).map(e=><li key={e} className="break-all">{e}</li>)}{prev.valid.length>10&&<li className={mut}>…and {prev.valid.length-10} more</li>}</ul>
      {prev.total>0&&!prev.valid.length&&<div className="text-sm text-[#F59E0B]">No new emails to add: all {prev.total-prev.inv} valid emails already exist. You can still process the file to log it in history.</div>}
      {prev.total===0&&<div className="text-sm text-[#F59E0B]">No emails found in this column or sheet.</div>}
      <button className={primary} disabled={busy||prev.total===0} onClick={commit}>{prev.valid.length?`Import ${prev.valid.length} Accounts`:'Process file (0 new)'}</button></div>}
  </div>);
}

const Row=memo(function Row({a,n,sel,copied,onSel,onUse,onCopy}:{a:Account;n:number;sel:boolean;copied:boolean;onSel:(id:number,v:boolean)=>void;onUse:(a:Account)=>void;onCopy:(a:Account)=>void}){
  const used=a.status==='used';
  return(<div className={`flex flex-wrap items-center gap-x-3 gap-y-2 px-3 py-3 border-b border-[var(--bd)] border-l-4 transition-colors ${used?'bg-[#22C55E]/10 border-l-[#22C55E]':'border-l-transparent hover:bg-[var(--bd)]'}`}>
    <input type="checkbox" className="size-5" checked={sel} aria-label={`Select ${a.email}`} onChange={e=>onSel(a.id!,e.target.checked)}/>
    <span className={`w-8 text-xs ${mut}`}>{String(n).padStart(2,'0')}</span>
    <span className="flex-1 min-w-[180px] break-all text-sm">{used&&<Check size={14} className="inline mr-1 text-[#22C55E]"/>}{a.email}</span>
    <button className={`${ghost} !py-2`} aria-label={`Copy ${a.email}`} onClick={()=>onCopy(a)}>{copied?<><Check className="inline" size={12}/> COPIED</>:<><Copy className="inline" size={12}/> COPY</>}</button>
    <label className="flex items-center gap-2 text-xs min-h-[40px]"><input type="checkbox" className="size-5 accent-[#22C55E]" checked={used} onChange={()=>onUse(a)}/>Used</label>
    <span className={`text-xs px-2 py-1 rounded-full ${used?'bg-[#22C55E]/20 text-[#22C55E]':'bg-[var(--bd)]'}`}>{used?'USED':'UNUSED'}</span>
    <span className={`text-xs ${mut} w-24 hidden lg:block`}>{new Date(a.importedAt).toLocaleDateString()}</span></div>);
});

function Accounts({all,ask}:{all:Account[];ask:Ask}){
  const [q,setQ]=useState(''),[tab,setTab]=useState<'all'|'used'|'unused'>('all'),[sort,setSort]=useState(localStorage.getItem('ov_sort')||'az'),[page,setPage]=useState(0),[sel,setSel]=useState<Set<number>>(new Set()),[copied,setCopied]=useState<number|null>(null);
  const PS=50,tm=useRef<number>();useEffect(()=>()=>clearTimeout(tm.current),[]);
  const serial=useMemo(()=>new Map(all.map((a,i)=>[a.id!,i+1])),[all]);
  const cnt={all:all.length,used:all.filter(a=>a.status==='used').length,unused:all.filter(a=>a.status==='unused').length};
  useEffect(()=>setSel(s=>{const ids=new Set(all.map(a=>a.id!));const k=[...s].filter(i=>ids.has(i));return k.length===s.size?s:new Set(k);}),[all]);
  const list=useMemo(()=>{
    const ql=q.trim().toLowerCase(),u=(a:Account)=>a.status==='used'?1:0;
    const f:Record<string,(a:Account,b:Account)=>number>={az:(a,b)=>a.email.localeCompare(b.email),za:(a,b)=>b.email.localeCompare(a.email),new:(a,b)=>U.importTs(b)-U.importTs(a)||b.id!-a.id!,old:(a,b)=>U.importTs(a)-U.importTs(b)||a.id!-b.id!,usedf:(a,b)=>u(b)-u(a),unusedf:(a,b)=>u(a)-u(b)};
    return all.filter(a=>(tab==='all'||a.status===tab)&&(!ql||a.email.toLowerCase().includes(ql)||String(serial.get(a.id!)).includes(ql))).sort(f[sort]||f.az);
  },[all,q,tab,sort,serial]);
  const pages=Math.max(1,Math.ceil(list.length/PS));
  useEffect(()=>setPage(0),[q,tab,sort]);useEffect(()=>{if(page>=pages)setPage(pages-1);},[page,pages]);
  const vis=list.slice(page*PS,page*PS+PS);
  const setStatus=(ids:number[],status:'used'|'unused')=>db.accounts.where('id').anyOf(ids).modify({status,updatedAt:new Date()});
  const onSel=useCallback((id:number,v:boolean)=>setSel(s=>{const n=new Set(s);v?n.add(id):n.delete(id);return n;}),[]);
  const onUse=useCallback((a:Account)=>{db.accounts.update(a.id!,{status:a.status==='used'?'unused':'used',updatedAt:new Date()});},[]);
  const copy=useCallback(async(text:string,msg:string,id?:number)=>{
    if(await U.copyText(text)){toast.success(msg);if(id){setCopied(id);clearTimeout(tm.current);tm.current=window.setTimeout(()=>setCopied(null),1500);}}else toast.error('Copy blocked by the browser. Select the text and copy manually.');
  },[]);
  const onCopy=useCallback((a:Account)=>copy(a.email,'Copied Successfully',a.id),[copy]);
  const picked=all.filter(a=>sel.has(a.id!)),ids=[...sel];
  return(<div className="space-y-3">
    <div className={`${card} space-y-3`}>
      <div className="flex flex-wrap gap-2"><div className="relative flex-1 min-w-[200px]"><Search size={14} className={`absolute left-3 top-3.5 ${mut}`}/>
        <input value={q} onChange={e=>setQ(e.target.value)} aria-label="Search accounts" placeholder="Search email or serial number" className={`${inp} pl-9 pr-9`}/>
        {q&&<button aria-label="Clear search" className="absolute right-3 top-3" onClick={()=>setQ('')}><X size={16}/></button>}</div>
        <select aria-label="Sort" value={sort} onChange={e=>{setSort(e.target.value);localStorage.setItem('ov_sort',e.target.value);}} className={`${inp} !w-auto`}>
          <option value="az">A to Z</option><option value="za">Z to A</option><option value="new">Recently Imported</option><option value="old">Oldest Imported</option><option value="usedf">Used First</option><option value="unusedf">Unused First</option></select></div>
      <div className="flex flex-wrap gap-2">{(['all','used','unused'] as const).map(t=><button key={t} aria-pressed={tab===t} onClick={()=>setTab(t)} className={`${btn} ${tab===t?'bg-[#6366F1] text-white':'bg-[var(--bd)]'}`}>{t.toUpperCase()} ({cnt[t]})</button>)}</div>
      <div className="flex flex-wrap gap-2 text-xs items-center"><span className={mut}>{sel.size} selected</span>
        <button className={ghost} onClick={()=>setSel(s=>{const n=new Set(s);vis.forEach(a=>n.add(a.id!));return n;})}>Select all visible</button><button className={ghost} onClick={()=>setSel(new Set())}>Deselect all</button>
        <button className={ghost} disabled={!sel.size} onClick={()=>setStatus(ids,'used')}>Mark used</button><button className={ghost} disabled={!sel.size} onClick={()=>setStatus(ids,'unused')}>Mark unused</button>
        <button className={ghost} disabled={!sel.size} onClick={()=>copy(picked.map(a=>a.email).join('\n'),`Copied ${picked.length} emails`)}>Copy selected</button>
        <button className={`${ghost} text-[#EF4444]`} disabled={!sel.size} onClick={()=>ask('Delete accounts',`Permanently delete ${sel.size} selected account(s)?`,async()=>{await db.accounts.bulkDelete(ids);setSel(new Set());toast.success('Deleted');})}><Trash2 className="inline" size={12}/> Delete</button></div></div>
    <div className={`${card} !p-0 overflow-hidden`}>{vis.map(a=><Row key={a.id} a={a} n={serial.get(a.id!)!} sel={sel.has(a.id!)} copied={copied===a.id} onSel={onSel} onUse={onUse} onCopy={onCopy}/>)}
      {!vis.length&&<div className={`p-10 text-center ${mut}`}><Inbox className="mx-auto mb-2" size={32}/>{all.length?'No accounts match your search.':'No accounts yet. Import a spreadsheet to begin.'}</div>}</div>
    <div className="flex justify-center items-center gap-3 text-sm"><button className={ghost} disabled={page===0} onClick={()=>setPage(page-1)}>Prev</button>Page {page+1} / {pages} · {list.length} results<button className={ghost} disabled={page+1>=pages} onClick={()=>setPage(page+1)}>Next</button></div></div>);
}

function History({ask}:{ask:Ask}){
  const logs=useLiveQuery(()=>db.imports.orderBy('date').reverse().toArray(),[])||[];
  return(<div className="space-y-3"><div className="flex justify-end"><button className={`${ghost} text-[#EF4444]`} disabled={!logs.length} onClick={()=>ask('Clear import history','Only the history log is removed. Your accounts are kept.',()=>db.imports.clear())}>Clear history</button></div>
    <div className={`${card} !p-0`}>{logs.map(l=><div key={l.id} className="p-4 border-b border-[var(--bd)] grid grid-cols-2 md:grid-cols-6 gap-2 text-sm"><span className="col-span-2 md:col-span-1">{new Date(l.date).toLocaleString()}</span><b className="break-all md:col-span-1">{l.fileName}</b>
      <span>Rows {l.total}</span><span className="text-[#22C55E]">Imported {l.imported}</span><span className="text-[#F59E0B]">Dupes {l.duplicates}</span><span className="text-[#EF4444]">Invalid {l.invalid}</span></div>)}
      {!logs.length&&<div className={`p-10 text-center ${mut}`}><Inbox className="mx-auto mb-2" size={32}/>No imports yet.</div>}</div></div>);
}

function Section({t,children}:{t:string;children:React.ReactNode}){return <div className={`${card} space-y-3`}><div className="font-semibold">{t}</div>{children}</div>;}
function SettingsPage({all,theme,setTheme,ask}:{all:Account[];theme:string;setTheme:(t:string)=>void;ask:Ask}){
  const [lock,setLock]=useState(localStorage.getItem('ov_lock')||'15'),rf=useRef<HTMLInputElement>(null);
  const restore=async(f?:File)=>{
    if(!f)return;if(rf.current)rf.current.value='';
    try{
      const p=parseBackup(JSON.parse(await f.text()),new Set(all.map(a=>a.normalizedEmail)));
      ask('Restore backup',`Add ${p.add.length} new accounts? Skipping ${p.dD} that already exist, ${p.dF} duplicates inside the file and ${p.inv} invalid entries. Existing records are not changed.`,async()=>{
        try{const r=await restoreAccounts(p.add);toast.success(`Restore done: ${r.added} added, ${r.skipped+p.dD+p.dF} skipped, ${p.inv} invalid`);}catch{toast.error('Restore failed. No changes were made.');}});
    }catch{toast.error('Not a valid backup file');}
  };
  const used=all.filter(a=>a.status==='used'),unused=all.filter(a=>a.status==='unused');
  return(<div className="grid lg:grid-cols-2 gap-4">
    <Section t="Appearance"><div className="flex gap-2">{['dark','light'].map(t=><button key={t} aria-pressed={theme===t} className={`${btn} ${theme===t?'bg-[#6366F1] text-white':'bg-[var(--bd)]'}`} onClick={()=>setTheme(t)}>{t==='dark'?'Dark':'Light'} theme</button>)}</div></Section>
    <Section t="Security"><label className="text-sm flex items-center gap-2">Auto-lock after<select className={`${inp} !w-auto`} value={lock} onChange={e=>{setLock(e.target.value);localStorage.setItem('ov_lock',e.target.value);}}>{['5','15','30','60'].map(v=><option key={v} value={v}>{v} min</option>)}</select></label>
      <p className={`text-xs ${mut}`}>The in-app password is a convenience lock only, not server authentication. Protect the site with Cloudflare Access (see README). Data lives in this browser profile on this device only.</p></Section>
    <Section t="Export Data"><div className="flex flex-wrap gap-2 text-xs">{([['All',all,'all'],['Used',used,'used'],['Unused',unused,'unused']] as const).map(([n,l,k])=>(['xlsx','csv'] as const).map(f=><button key={n+f} className={ghost} disabled={!l.length} onClick={()=>U.exportFile(l,`outlook-vault-${k}`,f)}><Download className="inline" size={12}/> {n} .{f}</button>))}</div></Section>
    <Section t="Backup & Restore"><div className="flex gap-2 flex-wrap"><button className={ghost} disabled={!all.length} onClick={()=>U.downloadJSON({app:'outlook-vault',version:2,exportedAt:new Date(),accounts:all},'outlook-vault-backup.json')}>Backup JSON</button>
      <input ref={rf} type="file" accept=".json" hidden aria-label="Choose backup file" onChange={e=>restore(e.target.files?.[0])}/><button className={ghost} onClick={()=>rf.current?.click()}>Restore JSON</button></div></Section>
    <Section t="Account Management"><button className={ghost} disabled={!used.length} onClick={()=>ask('Reset all statuses','Mark every account as UNUSED?',async()=>{await db.accounts.toCollection().modify({status:'unused',updatedAt:new Date()});toast.success('Statuses reset');})}>Reset all statuses</button></Section>
    <Section t="Danger Zone"><button className={`${btn} bg-[#EF4444] text-white`} disabled={!all.length} onClick={()=>ask('Clear all accounts','Delete ALL accounts permanently? Export a backup first.',async()=>{await db.accounts.clear();toast.success('All accounts cleared');})}>Clear all accounts</button></Section>
    <div className={`${card} text-sm ${mut} lg:col-span-2`}>OUTLOOK VAULT v1.1.0 · IndexedDB “OutlookVaultDB” · spreadsheets are never stored or uploaded · no Outlook passwords or tokens are handled.</div></div>);
}

export default function App(){
  const [authed,setAuthed]=useState(!!sessionStorage.getItem('ov_in')),[page,setPage]=useState<Page>('dash'),[open,setOpen]=useState(false),[col,setCol]=useState(false),[theme,setTheme]=useState(localStorage.getItem('ov_theme')||'dark');
  const [c,setC]=useState<{t:string;m:string;y:()=>void}|null>(null);
  const all=useLiveQuery(()=>db.accounts.orderBy('id').toArray(),[])||[];
  const ask:Ask=useCallback((t,m,y)=>setC({t,m,y}),[]);const closeC=useCallback(()=>setC(null),[]);
  useEffect(()=>{document.documentElement.classList.toggle('light',theme==='light');localStorage.setItem('ov_theme',theme);},[theme]);
  useEffect(()=>{document.documentElement.classList.toggle('light',theme==='light');},[authed,theme]);
  useEffect(()=>{
    if(!authed)return;let t=0;const reset=()=>{clearTimeout(t);t=window.setTimeout(()=>{sessionStorage.removeItem('ov_in');setAuthed(false);},+(localStorage.getItem('ov_lock')||15)*60000);};
    const ev=['click','keydown','touchstart'];ev.forEach(e=>addEventListener(e,reset));reset();
    const k=(e:KeyboardEvent)=>{if((e.ctrlKey||e.metaKey)&&e.key==='k'){e.preventDefault();setPage('accounts');}};addEventListener('keydown',k);
    return()=>{clearTimeout(t);ev.forEach(e=>removeEventListener(e,reset));removeEventListener('keydown',k);};
  },[authed]);
  const nav:[Page,string,typeof Cog][]=[['dash','Dashboard',LayoutDashboard],['accounts','All Accounts',ListI],['import','Import Accounts',Upload],['history','Import History',Hist],['settings','Settings',Cog]];
  const go=(p:Page)=>{setPage(p);setOpen(false);};
  if(!authed)return <><Toaster/><Login onOk={()=>setAuthed(true)}/></>;
  return(<><Toaster toastOptions={{style:{background:'var(--card)',color:'var(--fg)',border:'1px solid var(--bd)'}}}/>
    <AnimatePresence>{open&&<motion.div initial={{opacity:0}} animate={{opacity:1}} exit={{opacity:0}} className="fixed inset-0 z-30 bg-black/60 md:hidden" onClick={()=>setOpen(false)}/>}</AnimatePresence>
    <aside aria-label="Sidebar" className={`fixed inset-y-0 left-0 z-40 w-64 ${col?'md:w-20':'md:w-60'} bg-[var(--side)] border-r border-[var(--bd)] p-3 flex flex-col gap-4 transition-all duration-300 ${open?'translate-x-0':'-translate-x-full'} md:translate-x-0`}>
      <div className="flex items-center justify-between">{col?<Shield className="mx-auto text-[#6366F1]"/>:<Logo/>}<button className="md:hidden p-2" aria-label="Close menu" onClick={()=>setOpen(false)}><X/></button></div>
      <nav className="flex-1 space-y-1">{nav.map(([k,n,Ic])=><button key={k} aria-current={page===k?'page':undefined} title={n} onClick={()=>go(k)} className={`relative w-full flex items-center gap-3 rounded-xl px-3 py-3 text-sm ${page===k?'text-white':mut} hover:bg-[var(--bd)]`}>
        {page===k&&<motion.span layoutId="act" className="absolute inset-0 rounded-xl bg-gradient-to-r from-[#6366F1] to-[#8B5CF6]"/>}<Ic size={18} className="relative shrink-0"/><span className={`relative ${col?'md:hidden':''}`}>{n}</span></button>)}
        <button className={`w-full flex items-center gap-3 rounded-xl px-3 py-3 text-sm ${mut} hover:bg-[var(--bd)]`} onClick={()=>{sessionStorage.removeItem('ov_in');setAuthed(false);}}><LogOut size={18}/><span className={col?'md:hidden':''}>Lock & logout</span></button></nav>
      <div className={`text-xs ${mut} text-center`}>v1.1.0</div></aside>
    <div className={`transition-all duration-300 ${col?'md:pl-20':'md:pl-60'}`}>
      <header className="sticky top-0 z-20 flex items-center gap-2 px-4 py-3 border-b border-[var(--bd)] bg-[var(--bg)]/80 backdrop-blur">
        <button className="md:hidden p-2" aria-label="Open menu" onClick={()=>setOpen(true)}><Menu/></button>
        <button className="hidden md:block p-2" aria-label="Collapse sidebar" onClick={()=>setCol(!col)}><ChevronLeft className={col?'rotate-180':''}/></button>
        <div className="flex-1 min-w-0"><h1 className="font-semibold truncate">{page==='dash'?'Welcome back':page==='accounts'?'Manage Your Accounts':nav.find(n=>n[0]===page)![1]}</h1>
          <p className={`text-xs ${mut} truncate`}>{page==='accounts'?`Organize, copy and track your email accounts. ${all.length} total.`:new Date().toLocaleDateString(undefined,{weekday:'long',year:'numeric',month:'long',day:'numeric'})}</p></div>
        <button className={`${ghost} hidden sm:block !py-2`} onClick={()=>setPage('accounts')}><Search className="inline" size={14}/> Search <kbd className="text-xs">Ctrl K</kbd></button>
        <button className="p-2" aria-label="Toggle theme" onClick={()=>setTheme(theme==='dark'?'light':'dark')}>{theme==='dark'?<Sun size={18}/>:<Moon size={18}/>}</button>
        <span title="Convenience lock active" className="p-2 text-[#22C55E]"><Shield size={18}/></span></header>
      <main className="p-4 md:p-6 max-w-6xl mx-auto"><AnimatePresence mode="wait"><motion.div key={page} initial={{opacity:0,y:8}} animate={{opacity:1,y:0}} exit={{opacity:0}} transition={{duration:.15}}>
        {page==='dash'&&<Dashboard all={all} go={go}/>}{page==='accounts'&&<Accounts all={all} ask={ask}/>}{page==='import'&&<ImportPage all={all}/>}{page==='history'&&<History ask={ask}/>}{page==='settings'&&<SettingsPage all={all} theme={theme} setTheme={setTheme} ask={ask}/>}
      </motion.div></AnimatePresence></main></div>
    {c&&<Dlg c={c} close={closeC}/>}</>);
}
