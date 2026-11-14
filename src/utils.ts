import * as XLSX from 'xlsx';
export const isEmail=(s:string)=>/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(s);
export async function readWB(f:File){
  if(!/\.(xlsx|xls|csv)$/i.test(f.name))throw new Error('type');
  const wb=XLSX.read(await f.arrayBuffer(),{type:'array'});if(!wb.SheetNames.length)throw new Error('empty');return wb;
}
export const sheetRows=(wb:XLSX.WorkBook,n:string)=>XLSX.utils.sheet_to_json<any[]>(wb.Sheets[n],{header:1,blankrows:false,defval:''}).filter(r=>r.some(c=>String(c).trim()));
export function detectCol(rows:any[][]){
  const s=rows.slice(0,200),n=Math.max(0,...s.map(r=>r.length));let best=-1,bc=0;
  for(let c=0;c<n;c++){const k=s.filter(r=>isEmail(String(r[c]??'').trim())).length;if(k>bc){bc=k;best=c;}}
  if(best>=0)return best;
  const h=(rows[0]||[]).findIndex(x=>/e-?mail|address|account|login|user/i.test(String(x)));return h>=0?h:0;
}
export const hasHeader=(rows:any[][],col:number)=>{const v=String(rows[0]?.[col]??'').trim();return !!v&&!v.includes('@');};
export function extract(rows:any[][],col:number,have:Set<string>){
  const start=hasHeader(rows,col)?1:0,seen=new Set<string>(),valid:string[]=[];let dup=0,inv=0,total=0;
  for(let i=start;i<rows.length;i++){const v=String(rows[i][col]??'').trim();if(!v)continue;total++;
    if(!isEmail(v)){inv++;continue;}const n=v.toLowerCase();if(have.has(n)||seen.has(n)){dup++;continue;}seen.add(n);valid.push(v);}
  return{total,valid,dup,inv,header:start===1};
}
export async function copyText(t:string){
  try{if(navigator.clipboard&&window.isSecureContext){await navigator.clipboard.writeText(t);return true;}}catch{}
  try{const ta=document.createElement('textarea');ta.value=t;ta.style.cssText='position:fixed;opacity:0';document.body.appendChild(ta);ta.select();const ok=document.execCommand('copy');ta.remove();return ok;}catch{return false;}
}
export function exportFile(list:{email:string;status:string;importedAt:Date}[],name:string,fmt:'xlsx'|'csv'){
  const ws=XLSX.utils.json_to_sheet(list.map(a=>({Email:a.email,Status:a.status,ImportDate:new Date(a.importedAt).toISOString()})));
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Accounts');XLSX.writeFile(wb,`${name}.${fmt}`,{bookType:fmt});
}
export function downloadJSON(data:unknown,name:string){
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([JSON.stringify(data)],{type:'application/json'}));a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}
export const importTs=(a:{importedAt:unknown})=>{const n=+new Date(a.importedAt as string);return isNaN(n)?0:n;};
