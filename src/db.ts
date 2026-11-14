import Dexie,{Table} from 'dexie';
import {isEmail} from './utils';
export interface Account{id?:number;email:string;normalizedEmail:string;status:'used'|'unused';importedAt:Date;updatedAt:Date}
export interface ImportLog{id?:number;date:Date;fileName:string;total:number;imported:number;duplicates:number;invalid:number}
class DB extends Dexie{accounts!:Table<Account,number>;imports!:Table<ImportLog,number>;
constructor(){super('OutlookVaultDB');this.version(1).stores({accounts:'++id,&normalizedEmail,status,importedAt',imports:'++id,date'});}}
export const db=new DB();
export async function addMany(list:Account[]){let added=0;
  for(let i=0;i<list.length;i+=500){const ch=list.slice(i,i+500);
    try{await db.accounts.bulkAdd(ch);added+=ch.length;}catch(e){if(e instanceof Dexie.BulkError)added+=ch.length-e.failures.length;else throw e;}}
  return added;}

export const safeDate=(v:unknown,fb:Date=new Date()):Date=>{if(v==null||v==='')return fb;const x=v instanceof Date?v:new Date(v as string|number);return isNaN(+x)?fb:x;};
export function runImport(fileName:string,valid:string[],total:number,inv:number){
  return db.transaction('rw',db.accounts,db.imports,async()=>{
    const now=new Date(),uniq=[...new Map(valid.map(e=>[e.toLowerCase(),e])).values()],exist=new Set<string>();
    for(let i=0;i<uniq.length;i+=500)(await db.accounts.where('normalizedEmail').anyOf(uniq.slice(i,i+500).map(e=>e.toLowerCase())).keys()).forEach(k=>exist.add(String(k)));
    const added=await addMany(uniq.filter(e=>!exist.has(e.toLowerCase())).map(e=>({email:e,normalizedEmail:e.toLowerCase(),status:'unused' as const,importedAt:now,updatedAt:now})));
    const dup=total-inv-added;
    await db.imports.add({date:now,fileName,total,imported:added,duplicates:dup,invalid:inv});
    return{added,dup};
  });
}
export function parseBackup(j:any,have:Set<string>){
  const arr=Array.isArray(j)?j:j?.accounts;if(!Array.isArray(arr))throw new Error('format');
  const seen=new Set<string>(),add:Account[]=[],now=new Date();let inv=0,dF=0,dD=0;
  for(const d of arr){const e=String(d?.email??'').trim();if(!isEmail(e)){inv++;continue;}const n=e.toLowerCase();
    if(seen.has(n)){dF++;continue;}seen.add(n);if(have.has(n)){dD++;continue;}
    const imp=safeDate(d.importedAt,now);add.push({email:e,normalizedEmail:n,status:d.status==='used'?'used':'unused',importedAt:imp,updatedAt:safeDate(d.updatedAt,imp)});}
  return{add,dD,dF,inv};
}
export function restoreAccounts(add:Account[]){
  return db.transaction('rw',db.accounts,async()=>{
    const exist=new Set<string>();
    for(let i=0;i<add.length;i+=500)(await db.accounts.where('normalizedEmail').anyOf(add.slice(i,i+500).map(a=>a.normalizedEmail)).keys()).forEach(k=>exist.add(String(k)));
    const n=await addMany(add.filter(a=>!exist.has(a.normalizedEmail)));return{added:n,skipped:add.length-n};
  });
}
