import { getChatGPTUser } from '../app/chatgpt-auth';
import { env } from 'cloudflare:workers';
export function json(value: unknown, status=200){return Response.json(value,{status,headers:{'Cache-Control':'no-store','X-Content-Type-Options':'nosniff'}});}
export async function authorize(request: Request){
 const user=await getChatGPTUser();
 if(!user) return {error:json({error:'Connecte-toi pour accéder aux dossiers.'},401)};
 if(request.method!=='GET'){
  const origin=request.headers.get('origin');
  if(!origin || origin!==new URL(request.url).origin)return {error:json({error:'Origine refusée.'},403)};
  if(!request.headers.get('content-type')?.startsWith('application/json'))return {error:json({error:'Format JSON requis.'},415)};
 }
 return {owner:user.userId};
}
export function database(){if(!env.DB)throw Error('Stockage indisponible.');return env.DB;}
export async function body(request:Request){const text=await request.text();if(text.length>40000)throw Error('Dossier trop volumineux.');return JSON.parse(text);}
export function validId(id:string){return /^[a-f0-9-]{36}$/.test(id);}
