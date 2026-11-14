import {useState} from 'react';import toast from 'react-hot-toast';import {Lock} from 'lucide-react';import {card,primary} from './ui';
const sha=async(s:string)=>[...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(s)))].map(b=>b.toString(16).padStart(2,'0')).join('');
export const Logo=()=>(<div className="flex items-center gap-2"><svg width="32" height="32" viewBox="0 0 32 32"><path d="M16 2l12 4v9c0 8-5 13-12 15C9 28 4 23 4 15V6z" fill="#6366F1"/><rect x="9" y="11" width="14" height="10" rx="2" fill="none" stroke="#fff" strokeWidth="1.8"/><path d="M9 12l7 5 7-5" fill="none" stroke="#fff" strokeWidth="1.8"/></svg><div><div className="font-bold tracking-wide leading-none">OUTLOOK VAULT</div><div className="text-[10px] text-slate-400">Smart Personal Account Management</div></div></div>);

export function Login({onOk}:{onOk:()=>void}){
  const first=!localStorage.getItem('ov_hash');const [p,setP]=useState('');const [p2,setP2]=useState('');
  const go=async()=>{
    if(first){if(p.length<8)return toast.error('Use at least 8 characters');if(p!==p2)return toast.error('Passwords do not match');localStorage.setItem('ov_hash',await sha(p));sessionStorage.setItem('ov_in','1');onOk();}
    else if(await sha(p)===localStorage.getItem('ov_hash')){sessionStorage.setItem('ov_in','1');onOk();}else toast.error('Wrong password');
  };
  return(<div className="min-h-screen grid place-items-center p-4"><div className={`${card} w-full max-w-sm space-y-4`}><Logo/>
    <h2 className="font-semibold">{first?'Set your access password':'Unlock'}</h2>
    <input type="password" value={p} onChange={e=>setP(e.target.value)} onKeyDown={e=>e.key==='Enter'&&go()} placeholder="Password" className="w-full rounded-xl bg-[var(--bg)] border border-white/10 px-3 py-2"/>
    {first&&<input type="password" value={p2} onChange={e=>setP2(e.target.value)} placeholder="Confirm password" className="w-full rounded-xl bg-[var(--bg)] border border-white/10 px-3 py-2"/>}
    <button className={`${primary} w-full`} onClick={go}><Lock className="inline mr-1" size={14}/>{first?'Create & enter':'Unlock'}</button>
    <p className="text-xs text-slate-400">This screen is a convenience lock only. Real protection comes from Cloudflare Access (see README).</p></div></div>);
}

