import 'fake-indexeddb/auto';
import * as XLSX from 'xlsx';
import {db,runImport,parseBackup,restoreAccounts} from '../src/db';
import * as U from '../src/utils';
let fails=0;const ok=(c:boolean,m:string)=>{console.log((c?'PASS ':'FAIL ')+m);if(!c)fails++;};
const mk=(rows:any[][],fmt:'xlsx'|'biff8'|'csv',extra:{n:string;rows:any[][]}[]=[])=>{
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'S1');
  extra.forEach(x=>XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(x.rows),x.n));
  return new File([XLSX.write(wb,{type:'buffer',bookType:fmt})],`t.${fmt==='biff8'?'xls':fmt}`);};
async function imp(f:File,col?:number){
  const wb=await U.readWB(f),rows=U.sheetRows(wb,wb.SheetNames[0]),c=col??U.detectCol(rows);
  const have=new Set((await db.accounts.toArray()).map(a=>a.normalizedEmail)),p=U.extract(rows,c,have);
  return{p,wb,r:await runImport(f.name,p.valid,p.total,p.inv)};}
(async()=>{
  let o=await imp(mk([['Email','Name'],['a@x.com','A'],['b@x.com','B'],['B@X.com','B2'],['bad','x'],['c@x.com','C']],'xlsx'));
  ok(o.p.header&&o.p.total===5&&o.p.valid.length===3&&o.p.dup===1&&o.p.inv===1,'XLSX header, dup(case), invalid counts');
  ok(o.r.added===3&&await db.accounts.count()===3,'XLSX import saved 3');
  o=await imp(mk([['d@x.com'],['e@x.com']],'biff8'));ok(!o.p.header&&o.r.added===2,'XLS, email in first row kept (no header)');
  o=await imp(mk([['g@x.com'],['h@x.com']],'csv'));ok(o.r.added===2&&await db.accounts.count()===7,'CSV import, repeated uploads append (7 total)');
  o=await imp(mk([['Name','Account Email'],['n','f@x.com']],'xlsx'));ok(U.detectCol(U.sheetRows(o.wb,'S1'))===1&&o.r.added===1,'different column name detected');
  const ms=await U.readWB(mk([['i@x.com']],'xlsx',[{n:'S2',rows:[['j@x.com']]}]));ok(ms.SheetNames.length===2&&U.sheetRows(ms,'S2')[0][0]==='j@x.com','multiple sheets readable');
  let empty=false;try{const w=await U.readWB(new File([''],'e.csv'));empty=U.sheetRows(w,w.SheetNames[0]).length===0;}catch{empty=true;}ok(empty,'empty file handled');
  let bad=false;try{await U.readWB(new File(['x'],'a.txt'));}catch{bad=true;}ok(bad,'wrong file type rejected');
  const acc=(await db.accounts.where('normalizedEmail').equals('a@x.com').first())!;await db.accounts.update(acc.id!,{status:'used'});
  const hist=await db.imports.count();
  o=await imp(mk([['Email'],['A@x.com'],['b@x.com']],'xlsx'));
  ok(o.p.valid.length===0&&o.p.total===2&&o.r.added===0&&o.r.dup===2,'duplicate-only stats');
  ok(await db.accounts.count()===8&&(await db.accounts.get(acc.id!))!.status==='used','duplicate-only: no inserts, status kept');
  ok(await db.imports.count()===hist+1,'duplicate-only creates history entry');
  const json=JSON.parse(JSON.stringify({accounts:await db.accounts.toArray()}));
  ok(json.accounts[0].id&&json.accounts[0].normalizedEmail&&json.accounts[0].updatedAt,'backup has id, normalizedEmail, dates');
  const dump=[{email:'new1@x.com',status:'used',importedAt:'2020-01-01T00:00:00Z'},{email:'NEW1@x.com'},{email:'a@x.com',status:'unused'},{email:'zz'},{email:'new2@x.com',status:'weird',importedAt:'garbage'},5];
  const p=parseBackup(dump,new Set((await db.accounts.toArray()).map(a=>a.normalizedEmail)));
  ok(p.add.length===2&&p.dF===1&&p.dD===1&&p.inv===2,'restore parse: dup in file, dup in db, invalid');
  ok(p.add[1].status==='unused'&&!isNaN(+p.add[1].importedAt)&&p.add[0].status==='used','invalid status/date handled');
  const r=await restoreAccounts(p.add);const r2=await restoreAccounts(p.add);
  ok(r.added===2&&r2.added===0&&await db.accounts.count()===10&&(await db.accounts.get(acc.id!))!.status==='used','restore adds 2, re-restore adds 0, existing status kept');
  const l=await db.accounts.toArray();
  ok(l.sort((a,b)=>U.importTs(b)-U.importTs(a)||b.id!-a.id!)[l.length-1].email==='new1@x.com','Recently Imported sorts by importedAt not id');
  ok(l.sort((a,b)=>U.importTs(a)-U.importTs(b)||a.id!-b.id!)[0].email==='new1@x.com','Oldest sorts by importedAt');
  db.close();await db.open();ok((await db.accounts.get(acc.id!))!.status==='used'&&await db.accounts.count()===10,'data/status persist after reopen');
  console.log(fails?`${fails} FAILED`:'ALL PASSED');process.exit(fails?1:0);
})();
