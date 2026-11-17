import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync,mkdirSync,writeFileSync} from 'node:fs';
import {DatabaseSync} from 'node:sqlite';
import {calculateBudget,validateClient,emptyClient,sections} from '../lib/budget.mjs';
import {writeFile,readFile,listFiles,deleteFile} from '../lib/storage.mjs';
import {excelBytes,pdfBytes} from '../lib/exports.ts';
import ExcelJS from 'exceljs';
import {PDFDocument} from 'pdf-lib';
const cases=JSON.parse(readFileSync(new URL('./parity.json',import.meta.url)));
for(const [i,{data,expected}] of cases.entries())test(`Calculs identiques à Python, scénario ${i+1}`,()=>assert.deepEqual(calculateBudget(data),expected));
test('Validation des montants, profils et dates',()=>{
 const d=cases[0].data;assert.deepEqual(validateClient(d),d);
 for(const patch of [{dec1:NaN},{dec1:-1},{dec1:'1000'},{date_naissance:'2100-01-01'},{experience:{}}])assert.throws(()=>validateClient({...d,...patch}));
 assert.equal(calculateBudget(emptyClient()).revenus_retenus,0);
});
function adapter(raw){return {prepare(sql){return {bind(...args){return {async run(){const result=raw.prepare(sql).run(...args);return {meta:{changes:result.changes}};},async first(){return raw.prepare(sql).get(...args);},async all(){return {results:raw.prepare(sql).all(...args)};}};}};}};}
test('Stockage SQL : reprise, isolement des comptes et conflits de version',async()=>{
 const sqlite=new DatabaseSync(':memory:');
 const migrations=readFileSync(new URL('../drizzle/0000_lush_ezekiel_stane.sql',import.meta.url),'utf8');sqlite.exec(migrations);
 const db=adapter(sqlite),data=cases[0].data,id='00000000-0000-4000-8000-000000000001';
 assert.equal((await writeFile(db,'owner1',id,data,0)).version,1);
 assert.deepEqual((await readFile(db,'owner1',id)).data,data);
 assert.equal(await readFile(db,'owner2',id),null);
 assert.equal((await listFiles(db,'owner2')).length,0);
 assert.equal(await writeFile(db,'owner2',id,{...data,nom:'INTRUS'},0),null);
 assert.equal(await writeFile(db,'owner2',id,{...data,nom:'INTRUS'},1),null);
 assert.equal(await deleteFile(db,'owner2',id,1),false);
 assert.equal((await writeFile(db,'owner1',id,{...data,nom:'MISE À JOUR'},1)).version,2);
 assert.equal(await writeFile(db,'owner1',id,{...data,nom:'ANCIENNE VERSION'},1),null);
 assert.equal((await readFile(db,'owner1',id)).data.nom,'MISE À JOUR');
 assert.equal(await deleteFile(db,'owner1',id,1),false);
 assert.equal(await deleteFile(db,'owner1',id,2),true);
 assert.equal(await readFile(db,'owner1',id),null);sqlite.close();
});
test('Exports des 41 champs et résultats du client fictif',async()=>{
 const data=cases[0].data,excel=await excelBytes(data),pdf=await pdfBytes(data);
 const book=new ExcelJS.Workbook();await book.xlsx.load(excel);const rows=new Map();book.worksheets[0].eachRow(row=>rows.set(row.getCell(1).value,row.getCell(2).value));
 let count=0;for(const[,fields]of sections(data))for(const[label,value]of fields){assert.equal(rows.get(label),value);count++;}assert.equal(count,41);
 assert.equal(rows.get('Téléphone'),'0612345678');assert.equal((await PDFDocument.load(pdf)).getPageCount()>0,true);
 mkdirSync('outputs',{recursive:true});writeFileSync('outputs/client_fictif_web.xlsx',new Uint8Array(excel));writeFileSync('outputs/client_fictif_web.pdf',pdf);
 const literal=await excelBytes({...data,nom:'=1+2'});const b=new ExcelJS.Workbook();await b.xlsx.load(literal);assert.equal(b.worksheets[0].getCell('B4').value,'=1+2');
});
