// Compose existing SELF readers. Their shared projection coalesces the server
// read; no new authority, endpoint or institutional role is created here.
export async function readMyActivity(members){
 const [status,functions,directors]=await Promise.all([members.status(),members.currentFunctions(),members.currentDirectors()]);
 const rows=[...new Map([...functions.data,...directors.data].map(f=>[f.idVigencia,f])).values()];
 return {data:{membership:status.data,board:rows.filter(f=>f.group==='DIRETORIA'),advisory:rows.filter(f=>f.group==='ASSESSORIA'),council:rows.filter(f=>f.group==='CONSELHO'),other:rows.filter(f=>!['DIRETORIA','ASSESSORIA','CONSELHO'].includes(f.group))}};
}
export function renderMyActivity(data,escape,statusLabels){
 const sections=[['Diretoria atual',data.board],['Assessorias',data.advisory],['Conselho',data.council],['Outras funções',data.other]];
 return '<section><h3>Vínculo atual</h3><p>'+escape(statusLabels[data.membership.state]||data.membership.state)+'</p></section>'+sections.filter(([,rows])=>rows.length).map(([title,rows])=>'<section class="situation-section"><h3>'+title+'</h3>'+rows.map(f=>'<article><h4>'+escape(f.roleKey.replaceAll('_',' '))+'</h4><p>Início: '+escape(f.startDate)+'</p>'+(f.endDateReal||f.endDatePlanned?'<p>Fim: '+escape(f.endDateReal||f.endDatePlanned)+'</p>':'')+'<p>Gestão: '+escape(f.idDiretoria)+'</p></article>').join('')+'</section>').join('');
}
