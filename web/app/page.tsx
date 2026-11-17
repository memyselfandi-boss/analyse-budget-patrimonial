import { requireChatGPTUser } from './chatgpt-auth';
import PatrimoineApp from './patrimoine-app';
export const dynamic='force-dynamic';
export default async function Home(){await requireChatGPTUser('/');return <PatrimoineApp/>;}
