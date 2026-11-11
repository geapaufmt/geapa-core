'use strict';
// Documentary reproduction only. Never substitutes normative enablement or
// executes a membership mutation. Kept separate from immutable archived minutes.
const {jsPDF}=require('jspdf'),crypto=require('node:crypto'),assets=require('./minute-letterhead.json');
const VERSION='GEAPA_BOARD_PROCEDURE_PDF_1',sha=b=>crypto.createHash('sha256').update(b).digest('hex');
const states={APPROVED:'aprovada',REJECTED:'rejeitada',INSUFFICIENT_MANIFESTATIONS:'não aprovada por insuficiência de manifestações',CANCELLED:'cancelada',SUSPENDED_COMPOSITION:'suspensa por mudança da composição',RECORDED:'registrada'};
const stamp=n=>Number.isSafeInteger(n)?new Intl.DateTimeFormat('pt-BR',{timeZone:'America/Cuiaba',dateStyle:'long',timeStyle:'short',hour12:false}).format(new Date(n))+' (Cuiabá)':'não registrado';
function source(record){const p=record?.procedure;if(record?.kind!=='BOARD_PROCEDURE'||!states[p?.state]||!Number.isSafeInteger(record.revision)||!Number.isSafeInteger(record.updatedAt))throw Error('PROCEDURE_DOCUMENT_NOT_FINAL');const {document,documentHistory,...content}=p;return {id:record.id,revision:record.revision,title:record.title,updatedAt:record.updatedAt,procedure:content};}
function renderBoardProcedurePdf(record){
 const s=source(record),p=s.procedure,collective=['UNANIMITY','ABSOLUTE_MAJORITY'].includes(p.rule),sourceHash=sha(Buffer.from(JSON.stringify(s))),at=record.updatedAt;
 const doc=new jsPDF({unit:'mm',format:'a4',compress:true,putOnlyUsedFonts:true,precision:4});
 doc.addFileToVFS('NotoSerif-Regular.ttf',assets.regular);doc.addFont('NotoSerif-Regular.ttf','NotoSerif','normal');doc.addFileToVFS('NotoSerif-Bold.ttf',assets.bold);doc.addFont('NotoSerif-Bold.ttf','NotoSerif','bold');
 doc.setCreationDate('D:'+new Date(at).toISOString().slice(0,19).replace(/[-:T]/g,'')+"-00'00'");doc.setFileId(sourceHash.slice(0,32));
 // Without a verified enabling norm this must NOT call itself an official ata.
 const title=collective?'REGISTRO DE DELIBERAÇÃO ELETRÔNICA':'TERMO ADMINISTRATIVO';
 doc.setProperties({title,author:'GEAPA',creator:VERSION,subject:s.title});let y=54;
 const font=(size=10,bold=false)=>doc.setFont('NotoSerif',bold?'bold':'normal').setFontSize(size).setTextColor(3,3,2);
 function header(){doc.addImage(assets.icaa,'JPEG',18,10,22,23.4,'icaa');doc.addImage(assets.geapa,'PNG',164,10,30,24.9,'geapa');font(7.6);['UNIVERSIDADE FEDERAL DE MATO GROSSO','CÂMPUS UNIVERSITÁRIO DE SINOP','INSTITUTO DE CIÊNCIAS AGRÁRIAS E AMBIENTAIS','GRUPO DE ESTUDOS E APOIO À PRODUÇÃO AGRÍCOLA'].forEach((t,i)=>doc.text(t,52,15+i*4.6));doc.setDrawColor(92,55,26);doc.line(18,40,194,40);}
 function paragraph(text,bold=false,size=10){font(size,bold);const clean=String(text||'').replace(/\r\n?/g,'\n');if(/[\uD800-\uDFFF]/.test(clean))throw Error('DOCUMENT_UNSUPPORTED_CHARACTER');for(const line of doc.splitTextToSize(clean,164)){if(y>255){doc.addPage();header();y=52;font(size,bold);}doc.text(line,23,y);y+=size*.52;}y+=4;}
 header();paragraph(title,true,12);paragraph(s.title,true);paragraph('Identificação: '+s.id+' · revisão '+s.revision, false,8);
 if(collective){paragraph('Reprodução documental de procedimento eletrônico. Não equivale a ata formal nem autoriza efeitos institucionais sem a habilitação normativa e as demais condições aplicáveis.',true,9);paragraph('No período compreendido entre '+stamp(p.openedAt)+' e '+stamp(p.closedAt)+', foi registrada no Portal GEAPA a apreciação assíncrona da proposta abaixo pela composição identificada. Este documento não afirma a realização de reunião simultânea.');}
 else paragraph('Este termo reproduz o registro administrativo individual realizado no Portal GEAPA em '+stamp(p.createdAt)+'. Não constitui ata de reunião.');
 paragraph('Gestão: '+p.boardId+'. Proponente: '+p.proposer+'.');paragraph('Objeto: '+p.object);paragraph('Justificativa: '+p.reason);paragraph('Fundamentação declarada: '+p.basis);
 paragraph('Operação estruturada',true);for(const [key,value]of Object.entries(p.target))paragraph(key+': '+String(value),false,9);
 paragraph('Resultado registrado: '+states[p.state]+'.',true);paragraph('Efetivação: não realizada por este documento. A apuração e a aplicação da decisão são etapas distintas.');
 if(collective){paragraph('Regra: '+(p.rule==='UNANIMITY'?'unanimidade de toda a composição habilitada':'maioria absoluta de toda a composição habilitada')+'.');paragraph('Manifestações autenticadas',true);for(const person of p.electorate?.people||[]){const v=p.votes.find(v=>v.personId===person.personId);paragraph(person.name+' ('+person.personId+'), '+person.functions.map(f=>f.roleKey+' / '+f.linkId).join('; ')+': '+(v?({YES:'favorável',NO:'contrário',ABSTAIN:'abstenção'}[v.choice]||'inválida')+' em '+stamp(v.at):'sem manifestação')+'.',false,9);if(v)paragraph('Versão '+v.contentVersion+' · operação '+v.operationId,false,7);}}
 if(p.consent)paragraph('Aceite individual: '+(p.consent.accepted?'aceitou':'recusou')+' em '+stamp(p.consent.at)+'.');
 if(p.defense)paragraph('Manifestação do interessado, em '+stamp(p.defense.at)+': '+p.defense.statement);
 if(p.clearance)paragraph('Quitação administrativa: '+(p.clearance.pendingDuties||'sem pendências declaradas')+'. Não constitui veto ao desligamento.');
 for(const d of p.documents||[])paragraph('Documento relacionado: '+d.recordId+' · revisão '+d.revision+' · SHA-256 '+d.sha256,false,8);
 paragraph('Autenticação eletrônica interna, sem certificação qualificada ICP-Brasil.',false,8);paragraph('Conteúdo proposto: '+p.contentHash,false,7);paragraph('Verificação da reprodução: '+sourceHash,false,7);
 const pages=doc.getNumberOfPages();for(let i=1;i<=pages;i++){doc.setPage(i);font(7);doc.setDrawColor(130);doc.line(23,271,187,271);assets.footer.forEach((t,j)=>doc.text(t,23,276+j*3.4));doc.text(i+' / '+pages,187,293,{align:'right'});}
 const bytes=Buffer.from(doc.output('arraybuffer'));if(bytes.length>4194304)throw Error('SCOPE_TOO_LARGE');return {bytes,sha256:sha(bytes),size:bytes.length,renderer:VERSION,sourceHash,pages};
}
module.exports={renderBoardProcedurePdf,source,VERSION};
