import { authorize,json,body } from '../../../lib/api';
import { validateClient } from '../../../lib/budget.mjs';
import { excelBytes,pdfBytes } from '../../../lib/exports';
export async function POST(request:Request){
 const auth=await authorize(request);if(auth.error)return auth.error;
 try{
  const input=await body(request),data=validateClient(input.data),pdf=input.format==='pdf';
  if(!pdf&&input.format!=='xlsx')return json({error:'Format invalide.'},400);
  const bytes=pdf?await pdfBytes(data):await excelBytes(data);
  return new Response(bytes as BodyInit,{headers:{'Content-Type':pdf?'application/pdf':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','Content-Disposition':`attachment; filename="analyse_patrimoniale.${pdf?'pdf':'xlsx'}"`,'Cache-Control':'no-store','X-Content-Type-Options':'nosniff'}});
 }catch(e:any){return json({error:e.message||'Export impossible.'},400);}
}
