import schema from './schema.json' with { type: 'json' };
export { schema };
export const emptyClient = () => ({...Object.fromEntries(schema.identity.map(f=>[f.key,f.type==='number'?0:''])), ...Object.fromEntries(schema.budget.map(f=>[f.key,0])), credit_flag:0, experience:Object.fromEntries(schema.categories.map(c=>[c,false]))});
const cents = n => Math.round((n + Number.EPSILON) * 100) / 100;
export function calculateBudget(d) {
 const salaire=d.dec1+d.dec2, rev_loc=(d.bien1+d.bien2)*.8;
 const revenus=salaire+rev_loc, credits=d.emprunt_rp+d.autres_credits;
 const charges=d.charges+d.essence+d.impots, reste=revenus-credits-charges;
 return Object.fromEntries(Object.entries({salaire,rev_loc,revenus_retenus:revenus,charges_et_credits:charges+credits,reste_a_vivre:reste,reste_apres_loisirs:reste-d.loisirs_budget,capacite_endettement_max:revenus*.35,capacite_mensualite_restante:Math.max(0,revenus*.35-credits)}).map(([k,v])=>[k,cents(v)]));
}
export function validateClient(input) {
 if(!input || typeof input!=='object' || Array.isArray(input)) throw Error('Dossier invalide.');
 const result=emptyClient();
 for(const f of [...schema.identity,...schema.budget]){
  const v=input[f.key];
  if(f.type==='number'||schema.budget.some(b=>b.key===f.key)){
   if(typeof v!=='number'||!Number.isFinite(v)||v<0||v>1e10) throw Error(`Montant invalide : ${f.label}`);
  }else if(typeof v!=='string'||v.length>500){throw Error(`Champ invalide : ${f.label}`);}
  result[f.key]=v;
 }
 if(result.date_naissance && (!/^\d{4}-\d{2}-\d{2}$/.test(result.date_naissance)||(Number.isNaN(Date.parse(result.date_naissance))||new Date(result.date_naissance).toISOString().slice(0,10)!==result.date_naissance)||result.date_naissance<'1900-01-01'||result.date_naissance>new Date().toISOString().slice(0,10))) throw Error('Date de naissance invalide.');
 if(![0,1].includes(input.credit_flag))throw Error('Réponse crédits invalide.');
 result.credit_flag=input.credit_flag;
 if(!input.experience || typeof input.experience!=='object') throw Error('Profil invalide.');
 for(const cat of schema.categories){if(typeof input.experience[cat]!=='boolean')throw Error('Profil invalide.'); result.experience[cat]=input.experience[cat];}
 return result;
}
/** @returns {Array<[string, Array<[string, any]>]>} */
export function sections(d){const b=calculateBudget(d);return [
 ['Identité',schema.identity.map(f=>[f.label,d[f.key]])],
 ['Budget mensuel',[["Total des salaires mensuels (avant PAS)",b.salaire],...schema.budget.slice(0,2).map(f=>[f.label,d[f.key]]),["Total des loyers retenus à 80 %",b.rev_loc],...schema.budget.slice(2).map(f=>[f.label,d[f.key]]),['Crédits - souscripteur ?',d.credit_flag?'Oui':'Non']]],
 ['Indicateurs mensuels (euros)',Object.entries(schema.metrics).map(([key,label])=>[label,b[key]])],
 ['Profil financier',schema.categories.map(c=>[c,d.experience[c]?'Oui':'Non'])]
 ];}
