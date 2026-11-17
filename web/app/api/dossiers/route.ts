import { authorize, database, json } from '../../../lib/api';
import { listFiles } from '../../../lib/storage.mjs';
export async function GET(request:Request){const auth=await authorize(request);if(auth.error)return auth.error;return json({dossiers:await listFiles(database(),auth.owner)});}
