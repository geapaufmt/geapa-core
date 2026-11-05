'use strict';
// Pure, deterministic document rendering. No URL fetch, HTML execution or browser printing.
const {jsPDF}=require('jspdf'),crypto=require('node:crypto'),assets=require('./minute-letterhead.json');
const VERSION='GEAPA_MINUTE_PDF_1',sha=b=>crypto.createHash('sha256').update(b).digest('hex');
const roles={PRESIDENTE:'Presidência',VICE_PRESIDENTE:'Vice-Presidência',DIRETOR_ADMINISTRATIVO:'Diretoria Administrativa',DIRETOR_RECURSOS_HUMANOS:'Diretoria de Recursos Humanos',DIRETOR_FINANCEIRO:'Diretoria Financeira',TESOUREIRO:'Tesouraria',DIRETOR_COMUNICACAO:'Diretoria de Comunicação',DIRETOR_EXTENSAO:'Diretoria de Extensão'};
const stamp=n=>new Intl.DateTimeFormat('pt-BR',{timeZone:'America/Cuiaba',dateStyle:'short',timeStyle:'short',hour12:false}).format(new Date(n));
function renderMinutePdf(record,{draft=false}={}){
 const m=record?.minutes;if(!m||!draft&&m.state!=='ARCHIVED'||m.fullText&&m.documentLayoutVersion!==VERSION)throw Error('MINUTE_DOCUMENT_INVALID');
 for(const im of m.images||[]){const b=Buffer.from(im.base64,'base64');if(b.length>67500||sha(b)!==im.sha256)throw Error('MINUTE_IMAGE_INVALID');let w=0,h=0;
  if(im.mime==='image/png'&&b.subarray(0,8).equals(Buffer.from([137,80,78,71,13,10,26,10]))&&b.length>=24){w=b.readUInt32BE(16);h=b.readUInt32BE(20);}
  else if(im.mime==='image/jpeg'&&b.readUInt16BE(0)===0xffd8){let p=2;while(p+9<b.length){if(b[p++]!==255)break;const marker=b[p++],n=b.readUInt16BE(p);if([192,193,194].includes(marker)){h=b.readUInt16BE(p+3);w=b.readUInt16BE(p+5);break;}if(n<2)break;p+=n;}}
  if(!w||!h||w>2000||h>2000||w*h>2000000)throw Error('MINUTE_IMAGE_INVALID');
 }
 const hash=m.archive?.contentHash||m.submitted?.contentHash||sha(Buffer.from(JSON.stringify(record)));
 const at=m.archive?.at||record.updatedAt; if(!Number.isSafeInteger(at))throw Error('MINUTE_DOCUMENT_INVALID');
 const doc=new jsPDF({unit:'mm',format:'a4',orientation:'portrait',compress:true,putOnlyUsedFonts:true,precision:4});
 doc.addFileToVFS('NotoSerif-Regular.ttf',assets.regular);doc.addFont('NotoSerif-Regular.ttf','NotoSerif','normal');
 doc.addFileToVFS('NotoSerif-Bold.ttf',assets.bold);doc.addFont('NotoSerif-Bold.ttf','NotoSerif','bold');
 doc.setCreationDate(new Date(at));doc.setFileId(hash.slice(0,32));doc.setProperties({title:'Ata '+m.number,subject:'Documento institucional - texto versão '+m.contentRevision,author:'GEAPA',creator:VERSION,keywords:'GEAPA; '+hash});
 const left=23,width=164,bottom=256;let y=61;
 const font=(size=10.5,bold=false)=>doc.setFont('NotoSerif',bold?'bold':'normal').setFontSize(size).setTextColor(3,3,2);
 function header(){
  doc.addImage(assets.icaa,'JPEG',18,10,22,23.4,'icaa');doc.addImage(assets.geapa,'PNG',164,10,30,24.9,'geapa');
  doc.setDrawColor(80);doc.setLineWidth(.2);doc.line(46,11,46,34);font(7.6);
  ['UNIVERSIDADE FEDERAL DE MATO GROSSO','CÂMPUS UNIVERSITÁRIO DE SINOP','INSTITUTO DE CIÊNCIAS AGRÁRIAS E AMBIENTAIS','GRUPO DE ESTUDOS E APOIO À PRODUÇÃO AGRÍCOLA'].forEach((s,i)=>doc.text(s,52,15+i*4.6));
  doc.line(18,40,194,40);
 }
 function page(){doc.addPage();header();y=51;}
 function space(h){if(y+h>bottom)page();}
 function line(s,last,size=10.5){font(size);if(/[\uD800-\uDFFF]/.test(s))throw Error('DOCUMENT_UNSUPPORTED_CHARACTER');
  const words=s.split(' '),gap=words.length>1?(width-doc.getTextWidth(words.join('')))/(words.length-1):0;
  if(!last&&words.length>1&&gap<7){let x=left;for(const w of words){doc.text(w,x,y);x+=doc.getTextWidth(w)+gap;}}else doc.text(s,left,y);
 }
 function paragraph(text,size=10.5){for(const p of String(text||'').replace(/\r\n?/g,'\n').split(/\n\s*\n/).filter(s=>s.trim())){font(size);const lines=doc.splitTextToSize(p.replace(/\n/g,' '),width),height=size*.49;space(Math.min(lines.length,3)*height+3);for(let i=0;i<lines.length;i++){if(i===lines.length-2)space(height*2);else space(height);line(lines[i],i===lines.length-1,size);y+=height;}y+=3.5;}}
 function heading(s){space(22);font(10,true);doc.text(s,left,y);y+=7;}
 header();font(13,true);doc.text('ATA Nº '+m.number,105,y,{align:'center',maxWidth:width});y+=8;
 font(10,true);const title=doc.splitTextToSize('REUNIÃO '+m.meetingType.toLocaleUpperCase('pt-BR')+' DA DIRETORIA DO GEAPA',width);doc.text(title,105,y,{align:'center'});y+=title.length*5+7;
 if(draft){font(8);doc.setTextColor(110);doc.text('PRÉVIA - TEXTO AINDA NÃO ARQUIVADO',105,y,{align:'center'});y+=8;}
 if(m.fullText)paragraph(m.fullText);
 else {paragraph('Representação dos campos documentais preservados. Esta versão foi arquivada antes do cadastro de texto integral; nenhuma redação narrativa foi acrescentada à versão aprovada.',8.5);paragraph('Pauta registrada: '+m.agenda);paragraph(record.body);paragraph('Decisões e resultados registrados: '+m.decisions);paragraph('Encaminhamentos registrados: '+m.outcomes);}
 const people=m.composition.people;
 heading('Participação registrada');
 for(const [kind,label]of [['PRESENT','Presentes'],['ABSENT','Ausentes']]){const names=people.filter(p=>m.attendance.some(a=>a.personId===p.personId&&a.presence===kind)).map(p=>p.name);if(names.length)paragraph(label+': '+names.join('; ')+'.',9);}
 for(const [action,title]of [['APPROVE_SIGN','Assinaturas eletrônicas internas dos presentes'],['ACKNOWLEDGE','Ciências dos ausentes']]){
  const rows=m.manifestations.filter(x=>x.action===action&&x.contentRevision===m.contentRevision);if(!rows.length)continue;heading(title);
  for(const r of rows){const p=people.find(p=>p.personId===r.personId);space(23);paragraph(p.name+' - '+r.functions.map(f=>roles[f.roleKey]||f.roleKey.replaceAll('_',' ')).join(', '),9);paragraph(stamp(r.at)+' (Cuiabá) - texto versão '+r.contentRevision+'.',8);}
 }
 const observations=m.manifestations.filter(x=>['OBSERVE','REQUEST_CORRECTION'].includes(x.action)&&x.contentRevision===m.contentRevision);
 if(observations.length){heading('Observações documentais');for(const x of observations)paragraph(people.find(p=>p.personId===x.personId).name+': '+(x.excerpt?x.excerpt+' - ':'')+x.reason,9);}
 space(22);paragraph('Ata '+m.number+' | Texto v. '+m.contentRevision+(m.archive?' | Arquivada em '+stamp(m.archive.at)+' (Cuiabá)':''),8);
 paragraph('Manifestações eletrônicas internas, sem certificação qualificada ICP-Brasil. A aprovação do texto não equivale a voto favorável a todas as deliberações.',8);
 if(record.references?.length){heading('Documentos relacionados');record.references.forEach((r,i)=>paragraph((i+1)+'. '+r.label+'. '+r.evidence,9));}
 for(const [index,img]of (m.images||[]).entries()){
  page();heading('Anexo fotográfico '+(index+1));const data='data:'+img.mime+';base64,'+img.base64,info=doc.getImageProperties(data);const w=Math.min(width,145),h=Math.min(165,w*info.height/info.width),iw=h*info.width/info.height;
  doc.addImage(data,img.mime==='image/png'?'PNG':'JPEG',105-iw/2,y,iw,h);y+=h+8;paragraph(img.caption,9);
 }
 const count=doc.getNumberOfPages();for(let i=1;i<=count;i++){doc.setPage(i);font(7);doc.setDrawColor(130);doc.line(23,271,187,271);assets.footer.forEach((s,j)=>doc.text(s,23,276+j*3.4));doc.text(i+' / '+count,187,293,{align:'right'});}
 const bytes=Buffer.from(doc.output('arraybuffer'));if(bytes.length>4*1024*1024)throw Error('SCOPE_TOO_LARGE');return {bytes,sha256:sha(bytes),size:bytes.length,renderer:VERSION,sourceHash:hash,pages:count};
}
module.exports={renderMinutePdf,VERSION};
