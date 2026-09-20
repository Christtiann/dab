const pptxgen = require('pptxgenjs');
const { iconSet } = require('./icons');

const BG="0B0B0B", CARD="161616", CARD2="1F1F1F", YEL="FFC709", W="FFFFFF",
      G1="A8A8A8", G2="737373", G3="3A3A3A";
const F="Arial";
const PW=13.3, M=0.6, CW=PW-2*M;

const sh = () => ({ type:"outer", color:"000000", blur:14, offset:3, angle:90, opacity:0.55 });

function newSlide(p){ const s=p.addSlide(); s.background={color:BG}; return s; }

function header(s, eyebrow, title, sub){
  s.addText(eyebrow, {x:M,y:0.42,w:CW,h:0.26,fontSize:11,bold:true,color:YEL,fontFace:F,charSpacing:2.4,isTextBox:true,margin:0});
  s.addText(title, {x:M,y:0.72,w:CW,h:0.62,fontSize:title.length>46?29:33,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  if(sub) s.addText(sub,{x:M,y:1.34,w:CW,h:0.34,fontSize:13.5,color:G1,fontFace:F,isTextBox:true,margin:0});
}

function footer(s,n){
  s.addText("XP Investimentos  ·  Planejamento Patrimonial",{x:M,y:7.02,w:6,h:0.26,fontSize:9,color:G2,fontFace:F,isTextBox:true,margin:0});
  s.addText(String(n).padStart(2,"0")+" / 10",{x:PW-M-1.4,y:7.02,w:1.4,h:0.26,fontSize:9,color:G2,fontFace:F,align:"right",isTextBox:true,margin:0});
}

function card(s,x,y,w,h,fill){
  s.addShape("roundRect",{x,y,w,h,fill:{color:fill||CARD},rectRadius:0.09,line:{color:G3,width:0.75},shadow:sh()});
}

function badge(s,x,y,d,txt,dark){
  s.addShape("ellipse",{x,y,w:d,h:d,fill:{color:YEL},line:{color:YEL,width:0}});
  s.addText(txt,{x,y,w:d,h:d,fontSize:15,bold:true,color:dark||"0B0B0B",fontFace:F,align:"center",valign:"middle",isTextBox:true,margin:0});
}

(async () => {
const I = await iconSet({
  user:"FiUser", chart:"FiBarChart2", shield:"FiShield", clock:"FiClock",
  alert:"FiAlertTriangle", trend:"FiTrendingUp", lock:"FiLock", target:"FiTarget",
  layers:"FiLayers", eye:"FiEye", file:"FiFileText", check:"FiCheck", calendar:"FiCalendar"
}, "#FFC709");
const ID = await iconSet({ trend:"FiTrendingUp" }, "#0B0B0B");

const p = new pptxgen();
p.layout = "LAYOUT_WIDE";
p.author = "Assessoria XP"; p.title = "Planejamento Financeiro Integrado";

/* ---------- 1 · CAPA ---------- */
let s = newSlide(p);
s.addText("XP INVESTIMENTOS  ·  PLANEJAMENTO PATRIMONIAL",{x:0.7,y:1.42,w:8.4,h:0.3,fontSize:11.5,bold:true,color:YEL,fontFace:F,charSpacing:2.4,isTextBox:true,margin:0});
s.addText("Planejamento\nFinanceiro Integrado",{x:0.7,y:1.88,w:8.3,h:1.9,fontSize:46,bold:true,color:W,fontFace:F,lineSpacing:50,isTextBox:true,margin:0});
s.addText("Estruturação consultiva para a proteção da sua família e o futuro da sua filha.",{x:0.7,y:3.92,w:7.6,h:0.6,fontSize:15.5,color:G1,fontFace:F,isTextBox:true,margin:0});

s.addShape("ellipse",{x:9.42,y:1.75,w:3.3,h:3.3,fill:{color:BG},line:{color:YEL,width:2}});
s.addShape("ellipse",{x:10.42,y:2.75,w:1.3,h:1.3,fill:{color:YEL},line:{color:YEL,width:0}});
s.addImage({data:ID.trend,x:10.72,y:3.05,w:0.7,h:0.7});

const chips=[["R$ 400.000","patrimônio total mapeado"],["3 caixas","alocação por objetivo"],["10%","em renda variável, na base total"]];
chips.forEach(([v,l],i)=>{
  const x=0.7+i*2.85;
  s.addText(v,{x,y:5.35,w:2.7,h:0.42,fontSize:21,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
  s.addText(l,{x,y:5.78,w:2.7,h:0.5,fontSize:10.5,color:G1,fontFace:F,isTextBox:true,margin:0});
});
s.addText("Reunião de diagnóstico  ·  30 minutos",{x:PW-M-4.5,y:7.02,w:4.5,h:0.26,fontSize:9.5,color:G2,fontFace:F,align:"right",isTextBox:true,margin:0});
s.addNotes("Olá, é um prazer receber você. Entendemos que sua rotina como engenheiro é intensa e seu tempo é escasso. O objetivo de hoje é mostrar como a curadoria da XP pode organizar seus R$ 400 mil com segurança, eficiência fiscal e zero trabalho no seu dia a dia. São 30 minutos: diagnóstico, estrutura proposta e próximos passos.");

/* ---------- 2 · DIAGNÓSTICO ---------- */
s = newSlide(p);
header(s,"DIAGNÓSTICO","A fotografia do seu momento financeiro","Patrimônio total de R$ 400.000 — e dois gargalos que resolvemos hoje.");

card(s,M,1.82,3.6,3.95);
s.addText("PERFIL",{x:M+0.28,y:2.06,w:3,h:0.26,fontSize:10,bold:true,color:YEL,fontFace:F,charSpacing:1.8,isTextBox:true,margin:0});
[["user","40 anos","Engenheiro"],["shield","Casado","1 filha dependente"],["chart","Moderado","perfil autodeclarado"],["trend","R$ 12.000","renda mensal"]].forEach(([ic,t,d],i)=>{
  const y=2.48+i*0.83;
  s.addImage({data:I[ic],x:M+0.28,y:y+0.06,w:0.3,h:0.3});
  s.addText(t,{x:M+0.72,y,w:2.6,h:0.3,fontSize:14.5,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:M+0.72,y:y+0.3,w:2.6,h:0.28,fontSize:11,color:G1,fontFace:F,isTextBox:true,margin:0});
});

const gar=[
 ["01","R$ 200.000 parados em conta corrente","Rendimento zero. Metade do seu patrimônio perde poder de compra todos os dias, enquanto a outra metade assume todo o risco."],
 ["02","R$ 100.000 em ações — 25% do patrimônio total","Exposição direta à bolsa acima do que um perfil moderado sustenta sem desconforto, ainda mais sem tempo para acompanhar o mercado."]
];
gar.forEach(([n,t,d],i)=>{
  const y=1.82+i*2.05;
  card(s,4.5,y,8.2,1.85,CARD2);
  badge(s,4.78,y+0.3,0.42,n);
  s.addText(t,{x:5.36,y:y+0.28,w:7.1,h:0.42,fontSize:16.5,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:5.36,y:y+0.78,w:7.1,h:0.85,fontSize:12,color:G1,fontFace:F,lineSpacing:17,isTextBox:true,margin:0});
});

s.addText("SUAS DORES, NAS SUAS PALAVRAS",{x:M,y:5.92,w:CW,h:0.26,fontSize:10,bold:true,color:YEL,fontFace:F,charSpacing:1.8,isTextBox:true,margin:0});
["Carteira sem plano definido","A oscilação incomoda","Sempre reclamando das taxas"].forEach((t,i)=>{
  const x=M+i*4.07;
  card(s,x,6.24,3.87,0.6,CARD2);
  s.addText('"'+t+'"',{x:x+0.2,y:6.24,w:3.5,h:0.6,fontSize:12.5,italic:true,color:W,fontFace:F,valign:"middle",isTextBox:true,margin:0});
});
footer(s,2);
s.addNotes("Analisando seu momento: você construiu um patrimônio excelente de R$ 400 mil. Mas ele está partido ao meio de forma ineficiente. Metade parada na conta corrente rendendo zero, e metade concentrada em ações — R$ 100 mil, que representam 25% de tudo o que você tem. É por isso que a oscilação incomoda: não é fragilidade sua, é dimensionamento. Vamos resolver os dois gargalos hoje.");

/* ---------- 3 · ATRITO ---------- */
s = newSlide(p);
header(s,"RECONQUISTA","Você já foi assessorado pela XP","O que muda desta vez — antes de falar de produto, precisamos falar do que não funcionou.");

s.addText("O QUE PROVAVELMENTE FALTOU",{x:M,y:1.86,w:5.9,h:0.28,fontSize:10.5,bold:true,color:G2,fontFace:F,charSpacing:1.8,isTextBox:true,margin:0});
s.addText("O NOSSO COMPROMISSO",{x:6.8,y:1.86,w:5.9,h:0.28,fontSize:10.5,bold:true,color:YEL,fontFace:F,charSpacing:1.8,isTextBox:true,margin:0});

const pares=[
 ["Uma carteira montada sem tese e sem objetivo nomeado","Cada real com uma finalidade e um prazo declarados"],
 ["Produto oferecido sem o custo total na mesa","Custo total apresentado antes de cada alocação"],
 ["Contato só quando havia algo para vender","Agenda fixa: relatório trimestral e 2 reuniões por ano"],
 ["Volatilidade sem contexto, gerando ansiedade","Risco dimensionado por caixa, não pela carteira inteira"]
];
pares.forEach(([a,b],i)=>{
  const y=2.24+i*0.83;
  card(s,M,y,5.9,0.7,CARD);
  s.addText(a,{x:M+0.24,y,w:5.45,h:0.7,fontSize:12,color:G1,fontFace:F,valign:"middle",isTextBox:true,margin:0});
  card(s,6.8,y,5.9,0.7,CARD2);
  s.addImage({data:I.check,x:7.02,y:y+0.235,w:0.23,h:0.23});
  s.addText(b,{x:7.38,y,w:5.1,h:0.7,fontSize:12,bold:true,color:W,fontFace:F,valign:"middle",isTextBox:true,margin:0});
});

card(s,M,5.78,CW,0.92,CARD2);
s.addText("Carteira sem plano definido não é um problema de produto. É um problema de método — e é exatamente isso que estamos corrigindo hoje.",
  {x:M+0.34,y:5.78,w:CW-0.68,h:0.92,fontSize:14.5,bold:true,italic:true,color:YEL,fontFace:F,valign:"middle",isTextBox:true,margin:0});
footer(s,3);
s.addNotes("Faço questão de começar por aqui. Você já foi nosso cliente assessorado e saiu — e isso é uma informação, não uma crítica. A dor que você descreve hoje, carteira sem plano definido, é o sintoma de um atendimento centrado em produto e não em objetivo. O que eu trago de diferente não é uma prateleira melhor: é método, custo na mesa antes da decisão e uma agenda de acompanhamento que existe independentemente de haver algo para vender.");

/* ---------- 4 · MÉTODO ---------- */
s = newSlide(p);
header(s,"METODOLOGIA","Seu patrimônio dividido por objetivo, não por produto","Três caixas independentes. Cada uma com um prazo, um risco e uma função.");

const caixas=[
 ["01","LIQUIDEZ","R$ 60.000","15% do patrimônio","Curto prazo","Reserva de emergência com resgate no mesmo dia e volatilidade zero.","lock"],
 ["02","PROTEÇÃO & FUTURO","R$ 140.000","35% do patrimônio","Médio e longo prazo","Educação da filha, eficiência fiscal e sucessão sem inventário.","shield"],
 ["03","CRESCIMENTO","R$ 200.000","50% do patrimônio","Longo prazo","Rentabilidade real acima da inflação, com risco dimensionado.","trend"]
];
caixas.forEach(([n,t,v,pc,pz,d,ic],i)=>{
  const x=M+i*4.07;
  card(s,x,1.92,3.87,3.86,i===1?CARD2:CARD);
  badge(s,x+0.3,2.2,0.46,n);
  s.addImage({data:I[ic],x:x+3.06,y:2.24,w:0.38,h:0.38});
  s.addText(t,{x:x+0.3,y:2.82,w:3.3,h:0.3,fontSize:11,bold:true,color:YEL,fontFace:F,charSpacing:1.4,isTextBox:true,margin:0});
  s.addText(v,{x:x+0.3,y:3.14,w:3.3,h:0.52,fontSize:26,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(pc,{x:x+0.3,y:3.66,w:3.3,h:0.28,fontSize:11.5,color:G1,fontFace:F,isTextBox:true,margin:0});
  s.addText(pz,{x:x+0.3,y:4.08,w:3.3,h:0.28,fontSize:11,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:x+0.3,y:4.42,w:3.3,h:1.1,fontSize:11.5,color:G1,fontFace:F,lineSpacing:16,isTextBox:true,margin:0});
});

card(s,M,5.94,CW,0.86,CARD2);
s.addText([
 {text:"Origem dos recursos:  ",options:{bold:true,color:YEL}},
 {text:"R$ 200.000 da conta corrente financiam as Caixas 1 e 2.  ·  Os R$ 200.000 já investidos são reestruturados na Caixa 3.  ·  ",options:{color:G1}},
 {text:"Premissa declarada: ",options:{bold:true,color:W}},
 {text:"o briefing não informa despesas; assumimos custo de vida de R$ 8.400/mês (70% da renda).",options:{color:G1}}
],{x:M+0.34,y:5.94,w:CW-0.68,h:0.86,fontSize:11,fontFace:F,valign:"middle",lineSpacing:15,isTextBox:true,margin:0});
footer(s,4);
s.addNotes("Para eliminar a ansiedade com oscilação, dividimos o patrimônio em três caixas independentes. O dinheiro de emergência fica imune a mercado. O capital da sua filha fica protegido da inflação e blindado na sucessão. E só o dinheiro de longo prazo assume risco. Repare numa coisa importante: eu declaro a premissa. O briefing não me deu suas despesas, então assumi 70% da renda. Se esse número for outro, a Caixa 1 muda e o resto se ajusta.");

/* ---------- 5 · CAIXA 1 ---------- */
s = newSlide(p);
header(s,"CAIXA 01  ·  CURTO PRAZO","Liquidez imediata e reserva de emergência","O primeiro dinheiro a sair da conta corrente — e o único que nunca oscila.");

card(s,M,1.92,4.0,4.35,CARD2);
s.addText("R$ 60.000",{x:M+0.34,y:2.3,w:3.4,h:0.75,fontSize:40,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
s.addText("15% do patrimônio total",{x:M+0.34,y:3.06,w:3.4,h:0.3,fontSize:12.5,color:G1,fontFace:F,isTextBox:true,margin:0});
s.addText("COMO CHEGAMOS AO VALOR",{x:M+0.34,y:3.72,w:3.4,h:0.26,fontSize:9.5,bold:true,color:YEL,fontFace:F,charSpacing:1.6,isTextBox:true,margin:0});
s.addText([
 {text:"Custo de vida estimado",options:{bold:true,color:W,breakLine:true}},
 {text:"R$ 8.400 / mês",options:{color:G1,breakLine:true}},
 {text:"Cobertura alvo (6 meses)",options:{bold:true,color:W,breakLine:true}},
 {text:"R$ 50.400",options:{color:G1,breakLine:true}},
 {text:"Arredondado para",options:{bold:true,color:W,breakLine:true}},
 {text:"R$ 60.000  =  7,1 meses",options:{color:G1}}
],{x:M+0.34,y:4.04,w:3.4,h:1.9,fontSize:11.5,fontFace:F,lineSpacing:17,isTextBox:true,margin:0});

const at=[
 ["CDB de liquidez diária D+0","Bancos de primeira linha na plataforma XP, faixa de 100% a 103% do CDI, resgate no mesmo dia pelo app.","lock"],
 ["Fundo DI com taxa zero de administração","Alternativa sem taxa para a parcela de giro, com liquidez no mesmo dia útil.","file"],
 ["Cobertura do FGC","Até R$ 250 mil por CPF por instituição, com teto global de R$ 1 milhão a cada 4 anos. Pulverizando em dois emissores, sua exposição fica muito abaixo do limite.","shield"]
];
at.forEach(([t,d,ic],i)=>{
  const y=1.92+i*1.5;
  card(s,4.9,y,7.8,1.35,CARD);
  s.addImage({data:I[ic],x:5.16,y:y+0.24,w:0.32,h:0.32});
  s.addText(t,{x:5.64,y:y+0.2,w:6.8,h:0.36,fontSize:14.5,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:5.64,y:y+0.6,w:6.8,h:0.62,fontSize:11.5,color:G1,fontFace:F,lineSpacing:16,isTextBox:true,margin:0});
});
s.addText("Volatilidade da caixa: zero.  Este dinheiro não participa de mercado — ele existe para você nunca precisar resgatar as outras duas caixas no pior momento.",
  {x:M,y:6.46,w:CW,h:0.42,fontSize:12,italic:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
footer(s,5);
s.addNotes("Sua reserva sai da conta corrente e vai para CDB de liquidez diária com garantia do FGC. Rende todos os dias e resgata no mesmo dia pelo aplicativo. Note que eu não inventei o número: parti de uma estimativa de custo de vida e mostrei a conta. Se o seu gasto real for maior, a gente recalcula aqui mesmo. E essa caixa tem uma função silenciosa que é a mais importante de todas: ela impede que você precise vender um investimento de longo prazo num momento ruim.");

/* ---------- 6 · CAIXA 2 ---------- */
s = newSlide(p);
header(s,"CAIXA 02  ·  MÉDIO E LONGO PRAZO","Proteção da família e o futuro da sua filha","R$ 140.000 alocados, mais a camada de proteção que faltava no seu planejamento.");

const c2=[
 ["VGBL","R$ 80.000","Aporte único","Tabela regressiva de IR: 10% a partir de 10 anos, incidente apenas sobre o rendimento. Sucessão direta à sua filha, sem inventário e sem custo de partilha.","shield"],
 ["LCI e LCA","R$ 60.000","Isentos de IR","Isenção total de Imposto de Renda para pessoa física e cobertura do FGC. Vencimentos escalonados e casados com os marcos escolares da sua filha. Carência mínima de 9 meses.","lock"],
 ["SEGURO DE VIDA","R$ 720.000","Capital segurado","Cinco vezes sua renda anual. Custeado pelo fluxo mensal, não pelo principal — estimativa a partir de R$ 250/mês, sujeita à análise de risco da seguradora.","user"]
];
c2.forEach(([t,v,sub,d,ic],i)=>{
  const x=M+i*4.07;
  card(s,x,1.92,3.87,3.36,i===2?CARD2:CARD);
  s.addImage({data:I[ic],x:x+0.3,y:2.18,w:0.34,h:0.34});
  s.addText(t,{x:x+0.3,y:2.66,w:3.3,h:0.3,fontSize:11,bold:true,color:YEL,fontFace:F,charSpacing:1.4,isTextBox:true,margin:0});
  s.addText(v,{x:x+0.3,y:2.96,w:3.3,h:0.5,fontSize:25,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(sub,{x:x+0.3,y:3.46,w:3.3,h:0.26,fontSize:11,color:G1,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:x+0.3,y:3.84,w:3.3,h:1.4,fontSize:11,color:G1,fontFace:F,lineSpacing:15.5,isTextBox:true,margin:0});
});

card(s,M,5.46,CW,1.34,CARD2);
s.addImage({data:I.alert,x:M+0.34,y:5.74,w:0.32,h:0.32});
s.addText("POR QUE VGBL E NÃO PGBL NO APORTE ÚNICO",{x:M+0.8,y:5.70,w:11,h:0.28,fontSize:10.5,bold:true,color:YEL,fontFace:F,charSpacing:1.4,isTextBox:true,margin:0});
s.addText("A dedução do PGBL é limitada a 12% da renda bruta tributável anual — no seu caso, R$ 17.280 por ano. Um aporte único de R$ 80 mil em PGBL desperdiçaria o benefício e ainda seria tributado sobre o valor total no resgate. Por isso: VGBL no aporte único, e PGBL como aporte mensal de R$ 1.440, exatamente dentro do teto, se você declara no modelo completo.",
  {x:M+0.8,y:6.00,w:11.1,h:0.72,fontSize:11,color:G1,fontFace:F,lineSpacing:15.5,isTextBox:true,margin:0});
footer(s,6);
s.addNotes("Para o futuro da sua filha, três camadas. VGBL para o aporte único, com alíquota caindo a 10% no longo prazo e transmissão direta para ela sem passar por inventário. LCI e LCA isentas de Imposto de Renda, com vencimentos casados com a escola e a faculdade. E aqui eu chamo sua atenção: eu incluí seguro de vida. Você tem 40 anos, renda única e uma filha dependente. Todo o plano que montamos pressupõe que você continue gerando R$ 12 mil por mês. O seguro é o que protege o plano contra a única hipótese que ele não sobrevive.");

/* ---------- 7 · CAIXA 3 ---------- */
s = newSlide(p);
header(s,"CAIXA 03  ·  LONGO PRAZO","Crescimento com risco dimensionado","R$ 200.000 reestruturados a partir da carteira que você já tem hoje.");

const c3=[
 ["55%","R$ 110.000","Renda fixa de carrego","Tesouro IPCA+ e debêntures incentivadas de infraestrutura (Lei 12.431), estas últimas isentas de IR para pessoa física. Levadas ao vencimento: com o prazo casado ao objetivo, a marcação a mercado deixa de ser um problema."],
 ["25%","R$ 50.000","Multimercados de baixa volatilidade","Gestão ativa da prateleira XP, com mandato de baixa volatilidade e alvo de retorno acima do CDI. Diversificação de estratégia sem exigir acompanhamento seu."],
 ["20%","R$ 40.000","Renda variável via ETFs","BOVA11 e IVVB11: bolsa brasileira e mercado americano em duas ordens. Diversificação automática, sem necessidade de escolher ação nem acompanhar notícia."]
];
c3.forEach(([pc,v,t,d],i)=>{
  const y=1.92+i*1.33;
  card(s,M,y,CW,1.2,i===0?CARD2:CARD);
  s.addText(pc,{x:M+0.32,y:y+0.2,w:1.2,h:0.5,fontSize:27,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
  s.addText(v,{x:M+0.32,y:y+0.72,w:1.5,h:0.3,fontSize:13,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(t,{x:M+2.0,y:y+0.2,w:9.9,h:0.32,fontSize:15,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:M+2.0,y:y+0.56,w:9.9,h:0.6,fontSize:11.5,color:G1,fontFace:F,lineSpacing:16,isTextBox:true,margin:0});
});

card(s,M,5.9,CW,0.9,CARD2);
s.addImage({data:I.eye,x:M+0.34,y:6.19,w:0.3,h:0.3});
s.addText([
 {text:"Por que não indicamos COE nesta carteira:  ",options:{bold:true,color:YEL}},
 {text:"o custo de estruturação de um COE não aparece destacado no seu extrato. Como transparência de taxas é justamente o seu ponto sensível, seria incoerente. Entregamos proteção de capital pela renda fixa de carrego, onde cada custo é declarado.",options:{color:G1}}
],{x:M+0.78,y:5.9,w:11.3,h:0.9,fontSize:11,fontFace:F,valign:"middle",lineSpacing:15.5,isTextBox:true,margin:0});
footer(s,7);
s.addNotes("Aqui está a mudança mais importante. Sua exposição a ações cai de R$ 100 mil para R$ 40 mil, e sai de ação individual para ETF — diversificação automática. A maior fatia vai para renda fixa levada ao vencimento, boa parte dela isenta de IR. E eu quero ser explícito sobre uma ausência: não tem COE aqui. COE é um produto legítimo, mas o custo de estruturação não fica visível no extrato. Como taxa é o seu ponto sensível, eu não vou te oferecer justamente o produto mais opaco da prateleira.");

/* ---------- 8 · ANTES x DEPOIS ---------- */
s = newSlide(p);
header(s,"CONSOLIDAÇÃO","Antes e depois, na mesma base de comparação","Todos os percentuais abaixo se referem ao patrimônio total de R$ 400.000.");

s.addChart("bar",[
 {name:"Hoje", labels:["Ações / ETF","Multimercado","RF tributada","Previdência (VGBL)","Isentos de IR","Liquidez / pós-fixado","Parado em conta"], values:[100,0,100,0,0,0,200]},
 {name:"Proposto", labels:["Ações / ETF","Multimercado","RF tributada","Previdência (VGBL)","Isentos de IR","Liquidez / pós-fixado","Parado em conta"], values:[40,50,50,80,120,60,0]}
],{
 x:M, y:2.0, w:7.7, h:4.6, barDir:"bar", barGapWidthPct:45,
 chartColors:["4F4F4F",YEL], showValue:true, dataLabelPosition:"outEnd",
 dataLabelColor:W, dataLabelFontSize:9.5, dataLabelFontFace:F, dataLabelFormatCode:'#,##0"k"',
 showLegend:true, legendPos:"t", legendColor:G1, legendFontSize:11, legendFontFace:F,
 catAxisLabelColor:W, catAxisLabelFontSize:10.5, catAxisLabelFontFace:F, catAxisLineShow:false,
 valAxisLabelColor:G2, valAxisLabelFontSize:9.5, valAxisLabelFontFace:F, valAxisLineShow:false,
 valAxisMinVal:0, valAxisMaxVal:210, valAxisMajorUnit:50, valGridLine:{color:"262626",size:1}, catGridLine:{style:"none"},
 plotArea:{fill:{color:BG}}, chartArea:{fill:{color:BG}}
});
s.addText("Valores em R$ mil",{x:M,y:6.62,w:7.7,h:0.26,fontSize:9.5,color:G2,fontFace:F,isTextBox:true,margin:0});

const kpis=[
 ["Exposição a ações","25%  →  10%","R$ 100.000 para R$ 40.000, agora via ETF"],
 ["Capital rendendo","50%  →  100%","nenhum real parado em conta corrente"],
 ["Ativos isentos de IR","R$ 0  →  R$ 120.000","LCI, LCA e debêntures incentivadas"],
 ["Taxa de custódia","R$ 0","renda fixa, fundos e previdência"]
];
kpis.forEach(([t,v,d],i)=>{
  const y=2.0+i*1.17;
  card(s,8.55,y,4.15,1.02,i%2===0?CARD2:CARD);
  s.addText(t,{x:8.81,y:y+0.13,w:3.7,h:0.24,fontSize:9.5,bold:true,color:YEL,fontFace:F,charSpacing:1.2,isTextBox:true,margin:0});
  s.addText(v,{x:8.81,y:y+0.36,w:3.7,h:0.36,fontSize:18,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:8.81,y:y+0.72,w:3.7,h:0.24,fontSize:9.5,color:G1,fontFace:F,isTextBox:true,margin:0});
});
footer(s,8);
s.addNotes("Esta é a fotografia final, e repare que eu comparo tudo na mesma base: o patrimônio total de R$ 400 mil. Sua exposição a ações cai de 25% para 10%. O capital que está de fato rendendo sai de 50% para 100%. E surgem R$ 120 mil em ativos isentos de Imposto de Renda, que hoje são zero. Nenhum real fica parado.");

/* ---------- 9 · CUSTO E TEMPO ---------- */
s = newSlide(p);
header(s,"O VALOR DO SERVIÇO","Custo transparente e tempo preservado","As duas dores que sobraram: reclamação com taxas e falta de tempo para acompanhar.");

card(s,M,1.92,5.9,4.6,CARD);
s.addImage({data:I.eye,x:M+0.34,y:2.2,w:0.34,h:0.34});
s.addText("TRANSPARÊNCIA DE CUSTOS",{x:M+0.34,y:2.68,w:5.2,h:0.3,fontSize:11,bold:true,color:YEL,fontFace:F,charSpacing:1.4,isTextBox:true,margin:0});
s.addText([
 {text:"Taxa zero de custódia ",options:{bold:true,color:W}},
 {text:"em renda fixa, fundos e previdência na plataforma XP.",options:{color:G1,breakLine:true}},
 {text:"Renda fixa ",options:{bold:true,color:W}},
 {text:"remunera pelo spread do emissor — apresentado a você antes da alocação, e não depois.",options:{color:G1,breakLine:true}},
 {text:"Fundos e previdência ",options:{bold:true,color:W}},
 {text:"com taxa de administração na lâmina. Nenhum produto com custo embutido e não destacado.",options:{color:G1,breakLine:true}},
 {text:"Research XP ",options:{bold:true,color:W}},
 {text:"analisa rating e saúde financeira de cada emissor antes de ele entrar na sua carteira.",options:{color:G1}}
],{x:M+0.34,y:3.06,w:5.25,h:3.2,fontSize:11.5,fontFace:F,lineSpacing:16,paraSpaceAfter:9,valign:"top",isTextBox:true,margin:0});

card(s,6.8,1.92,5.9,4.6,CARD2);
s.addImage({data:I.clock,x:7.14,y:2.2,w:0.34,h:0.34});
s.addText("O SEU TEMPO, PRESERVADO",{x:7.14,y:2.68,w:5.2,h:0.3,fontSize:11,bold:true,color:YEL,fontFace:F,charSpacing:1.4,isTextBox:true,margin:0});
s.addText([
 {text:"Monitoramento diário ",options:{bold:true,color:W}},
 {text:"feito pela nossa equipe. Você não precisa acompanhar notícia de mercado.",options:{color:G1,breakLine:true}},
 {text:"Relatório trimestral ",options:{bold:true,color:W}},
 {text:"consolidado e resumido no app da XP.",options:{color:G1,breakLine:true}},
 {text:"Duas reuniões por ano ",options:{bold:true,color:W}},
 {text:"de 30 minutos, para revisar metas familiares e reenquadrar o perfil.",options:{color:G1,breakLine:true}},
 {text:"Rebalanceamento por aporte ",options:{bold:true,color:W}},
 {text:"e não por venda — assim não antecipamos Imposto de Renda à toa.",options:{color:G1}}
],{x:7.14,y:3.06,w:5.25,h:3.2,fontSize:11.5,fontFace:F,lineSpacing:16,paraSpaceAfter:9,valign:"top",isTextBox:true,margin:0});

s.addText([
 {text:"Capacidade de aporte mensal:  ",options:{bold:true,color:YEL}},
 {text:"R$ 3.600  =  R$ 1.440 no PGBL (teto de 12%)  ·  R$ 250 no seguro de vida  ·  R$ 1.910 para rebalanceamento",options:{color:G1}}
],{x:M,y:6.64,w:CW,h:0.32,fontSize:11.5,fontFace:F,isTextBox:true,margin:0});
footer(s,9);
s.addNotes("As duas dores que faltavam. Taxa: custódia zero, e o custo de cada produto na mesa antes da decisão, não depois. Tempo: o acompanhamento diário é nosso. Você recebe um resumo trimestral no app e me dá duas reuniões de 30 minutos por ano. E o rebalanceamento acontece direcionando seu aporte mensal, o que evita venda e evita antecipar imposto. Sua sobra de R$ 3.600 por mês já tem destino definido.");

/* ---------- 10 · PLANO DE AÇÃO ---------- */
s = newSlide(p);
header(s,"PRÓXIMOS PASSOS","Plano de ação imediato","Quatro movimentos. O primeiro começa ainda nesta reunião.");

const steps=[
 ["01","HOJE","Atualizar o seu suitability na conta XP e validar a proposta de alocação das três caixas.","calendar"],
 ["02","ESTA SEMANA","Aplicar os R$ 200.000 da conta corrente: R$ 60 mil na reserva em CDB D+0 e R$ 140 mil em VGBL, LCI e LCA.","target"],
 ["03","PRÓXIMOS 5 MESES","Desmontar a posição em ações em parcelas de até R$ 20 mil por mês, aproveitando a isenção de IR sobre vendas de ações até esse limite mensal.","layers"],
 ["04","CONTÍNUO","Aporte mensal de R$ 3.600, relatório trimestral no app e revisão semestral de 30 minutos.","trend"]
];
steps.forEach(([n,q,d,ic],i)=>{
  const y=1.98+i*1.15;
  card(s,M,y,CW,1.02,i===0?CARD2:CARD);
  badge(s,M+0.3,y+0.28,0.46,n);
  s.addText(q,{x:M+0.95,y:y+0.18,w:3.0,h:0.3,fontSize:12,bold:true,color:YEL,fontFace:F,charSpacing:1.3,isTextBox:true,margin:0});
  s.addText(d,{x:M+0.95,y:y+0.5,w:10.1,h:0.42,fontSize:12,color:G1,fontFace:F,lineSpacing:16,isTextBox:true,margin:0});
  s.addImage({data:I[ic],x:11.9,y:y+0.33,w:0.36,h:0.36});
});

s.addText("Atualizar o suitability leva 10 minutos. Seus R$ 200 mil param de perder para a inflação já no próximo dia útil.",
  {x:M,y:6.6,w:CW,h:0.34,fontSize:12.5,bold:true,italic:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
footer(s,10);
s.addNotes("Para colocar o plano em prática, o primeiro passo é atualizar seu perfil de investidor — leva 10 minutos e dá para fazer agora. Chamo atenção para a etapa 3: a saída das ações é gradual, em parcelas de até R$ 20 mil por mês, porque vendas de ações até esse valor mensal são isentas de Imposto de Renda para pessoa física. Fazendo assim, a reestruturação inteira sai sem custo tributário. Posso abrir o aplicativo e começarmos?");

await p.writeFile({ fileName: "Case_XP_Planejamento_Financeiro.pptx" });
console.log("deck gerado");
})();
