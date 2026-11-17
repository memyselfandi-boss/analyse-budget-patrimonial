export async function listFiles(db,owner){const result=await db.prepare('SELECT id,name,version,updated_at FROM client_files WHERE owner=? ORDER BY updated_at DESC LIMIT 500').bind(owner).all();return result.results;}
export async function readFile(db,owner,id){const row=await db.prepare('SELECT id,payload,version,updated_at FROM client_files WHERE id=? AND owner=?').bind(id,owner).first();return row?{id:row.id,data:JSON.parse(row.payload),version:row.version,updated_at:row.updated_at}:null;}
export async function writeFile(db,owner,id,data,version){
 const now=new Date().toISOString(),name=([data.nom,data.prenom].filter(Boolean).join(' ')||'Sans nom').slice(0,200);
 const sql=version===0?'INSERT INTO client_files (id,owner,name,payload,version,created_at,updated_at) VALUES (?,?,?,?,1,?,?) ON CONFLICT(id) DO NOTHING':'UPDATE client_files SET name=?,payload=?,version=version+1,updated_at=? WHERE id=? AND owner=? AND version=?';
 const args=version===0?[id,owner,name,JSON.stringify(data),now,now]:[name,JSON.stringify(data),now,id,owner,version];
 const result=await db.prepare(sql).bind(...args).run();
 return result.meta.changes?{id,version:version+1,updated_at:now}:null;
}
export async function deleteFile(db,owner,id,version){const result=await db.prepare('DELETE FROM client_files WHERE id=? AND owner=? AND version=?').bind(id,owner,version).run();return !!result.meta.changes;}
