import { authorize, database, json, body, validId } from '../../../../lib/api';
import { validateClient } from '../../../../lib/budget.mjs';
import {readFile,writeFile,deleteFile} from '../../../../lib/storage.mjs';
type Context={params:Promise<{id:string}>};
export async function GET(request:Request,ctx:Context){
 const auth=await authorize(request);if(auth.error)return auth.error;const{id}=await ctx.params;
 if(!validId(id))return json({error:'Identifiant invalide.'},400);
 const row=await readFile(database(),auth.owner,id);
 return row?json(row):json({error:'Dossier introuvable.'},404);
}
export async function PUT(request:Request,ctx:Context){
 const auth=await authorize(request);if(auth.error)return auth.error;const{id}=await ctx.params;
 if(!validId(id))return json({error:'Identifiant invalide.'},400);
 let input:any,data:any;try{input=await body(request);data=validateClient(input.data);if(!Number.isInteger(input.version)||input.version<0)throw Error('Version invalide.');}catch(e:any){return json({error:e.message},400);}
 const result=await writeFile(database(),auth.owner,id,data,input.version);
 return result?json(result):json({error:'Ce dossier a été modifié ailleurs. Réouvre sa version enregistrée ou sauvegarde une copie.'},409);
}
export async function DELETE(request:Request,ctx:Context){
 const auth=await authorize(request);if(auth.error)return auth.error;const{id}=await ctx.params;
 if(!validId(id))return json({error:'Identifiant invalide.'},400);
 let version:number;try{version=(await body(request)).version;if(!Number.isInteger(version)||version<1)throw Error();}catch{return json({error:'Version requise.'},400);}
 const deleted=await deleteFile(database(),auth.owner,id,version);
 return deleted?json({deleted:true}):json({error:'Le dossier a changé. Recharge-le avant suppression.'},409);
}
