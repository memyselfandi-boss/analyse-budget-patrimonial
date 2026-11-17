import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import ExcelJS from 'exceljs';
import { sections } from './budget.mjs';
export async function excelBytes(data: any) {
 const rows: any[][]=[['Cabinet','CL Conseils']];
 for(const [title,fields] of sections(data))rows.push([], [title,'Valeur'],...fields);
 const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('Synthèse client');
 sheet.columns=[{width:62},{width:55}];
 sheet.addRows(rows);sheet.eachRow(row=>{row.alignment={vertical:'top',wrapText:true};row.height=32;});
 return book.xlsx.writeBuffer();
}
export async function pdfBytes(data: any) {
 const doc=await PDFDocument.create();const font=await doc.embedFont(StandardFonts.Helvetica),bold=await doc.embedFont(StandardFonts.HelveticaBold);
 let page=doc.addPage([595.28,841.89]),y=790;
 const text=(v:any)=>String(v??'').replace(/[\r\n]+/g,' ').split('').map(c=>{try{font.encodeText(c);return c;}catch{return '?';}}).join('');
 const wrap=(v:any,width:number,size:number)=>{const lines:string[]=[];let line='';for(const char of text(v)){if(font.widthOfTextAtSize(line+char,size)>width){lines.push(line);line='';}line+=char;}lines.push(line);return lines;};
 const newPage=()=>{page=doc.addPage([595.28,841.89]);y=790;};
 page.drawText('CL Conseils',{x:45,y,size:22,font:bold,color:rgb(.10,.16,.25)});y-=32;
 page.drawText('Analyse patrimoniale - Synthèse client',{x:45,y,size:14,font:bold});y-=35;
 for(const [title,fields] of sections(data)){
  if(y<115)newPage();
  page.drawText(title,{x:45,y,size:13,font:bold,color:rgb(.1,.3,.5)});y-=25;
  for(const [label,value]of fields){
   const left=wrap(label,270,10),right=wrap(value,205,10),height=Math.max(left.length,right.length)*14+10;
   if(y-height<45)newPage();
   page.drawRectangle({x:42,y:y-height+10,width:510,height,borderWidth:.3,borderColor:rgb(.8,.83,.87)});
   left.forEach((line,i)=>page.drawText(line,{x:49,y:y-i*14,size:10,font}));
   right.forEach((line,i)=>page.drawText(line,{x:330,y:y-i*14,size:10,font}));y-=height;
  }y-=22;
 }
 doc.setTitle('Analyse patrimoniale - CL Conseils');return doc.save();
}
