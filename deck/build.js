const pptxgen = require('pptxgenjs');

const BG="000000", TINT="121212", YEL="FFC709", W="FFFFFF",
      G1="A5A5A5", G2="6E6E6E", RULE="2B2B2B";
const F="Arial";
const PW=13.3, M=0.7, CW=PW-2*M;

function newSlide(p){ const s=p.addSlide(); s.background={color:BG}; return s; }

function rule(s,x,y,w,col){
  s.addShape("line",{x,y,w,h:0,line:{color:col||RULE,width:0.75}});
}
function block(s,x,y,w,h){
  s.addShape("rect",{x,y,w,h,fill:{color:TINT},line:{width:0}});
}
function header(s, sec, title, sub){
  s.addText(sec,{x:M,y:0.44,w:CW,h:0.26,fontSize:12,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
  s.addText(title,{x:M,y:0.74,w:CW,h:0.58,fontSize:title.length>44?30:34,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  if(sub) s.addText(sub,{x:M,y:1.34,w:CW,h:0.32,fontSize:13.5,color:G1,fontFace:F,isTextBox:true,margin:0});
}
function footer(s,n){
  rule(s,M,7.0,CW);
  s.addText("XP Investimentos  ·  Planejamento Patrimonial",{x:M,y:7.08,w:6,h:0.24,fontSize:9,color:G2,fontFace:F,isTextBox:true,margin:0});
  s.addText(String(n).padStart(2,"0"),{x:PW-M-1,y:7.08,w:1,h:0.24,fontSize:9,color:G2,fontFace:F,align:"right",isTextBox:true,margin:0});
}

const p = new pptxgen();
p.layout = "LAYOUT_WIDE";
p.author = "Assessoria XP"; p.title = "Planejamento Financeiro Integrado";
let s;

/* ---------- 1 · CAPA ---------- */
s = newSlide(p);
s.addText("Planejamento Financeiro",{x:M,y:2.05,w:11.6,h:0.95,fontSize:52,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
s.addText("Integrado",{x:M,y:2.95,w:11.6,h:0.95,fontSize:52,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
s.addText("Estruturação consultiva para a proteção da sua família e o futuro da sua filha.",
  {x:M,y:4.12,w:9.5,h:0.4,fontSize:16,color:G1,fontFace:F,isTextBox:true,margin:0});

rule(s,M,5.48,CW);
const chips=[["R$ 400.000","patrimônio total mapeado"],["3 caixas","alocação por objetivo"],["10%","em renda variável, na base total"],["R$ 0","de taxa de custódia"]];
chips.forEach(([v,l],i)=>{
  const x=M+i*3.0;
  s.addText(v,{x,y:5.66,w:2.8,h:0.4,fontSize:20,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(l,{x,y:6.06,w:2.8,h:0.48,fontSize:10.5,color:G2,fontFace:F,isTextBox:true,margin:0});
});
s.addText("XP Investimentos",{x:M,y:1.5,w:6,h:0.28,fontSize:12.5,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
s.addText("Reunião de diagnóstico  ·  30 minutos",{x:PW-M-4.5,y:1.5,w:4.5,h:0.28,fontSize:11,color:G2,fontFace:F,align:"right",isTextBox:true,margin:0});
s.addNotes("Olá, é um prazer receber você. Entendemos que sua rotina como engenheiro é intensa e seu tempo é escasso. O objetivo de hoje é mostrar como a curadoria da XP pode organizar seus R$ 400 mil com segurança, eficiência fiscal e zero trabalho no seu dia a dia. São 30 minutos: diagnóstico, estrutura proposta e próximos passos.");

/* ---------- 2 · DIAGNÓSTICO ---------- */
s = newSlide(p);
header(s,"Diagnóstico","A fotografia do seu momento financeiro","Patrimônio total de R$ 400.000 e dois gargalos a resolver.");

s.addText("Perfil",{x:M,y:1.98,w:3.4,h:0.28,fontSize:12,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
const perfil=[["Idade","40 anos"],["Profissão","Engenheiro"],["Estado civil","Casado, 1 filha"],["Renda mensal","R$ 12.000"],["Perfil declarado","Moderado"]];
perfil.forEach(([k,v],i)=>{
  const y=2.42+i*0.62;
  s.addText(k,{x:M,y,w:1.7,h:0.3,fontSize:11.5,color:G2,fontFace:F,isTextBox:true,margin:0});
  s.addText(v,{x:M+1.7,y,w:1.9,h:0.3,fontSize:13,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  rule(s,M,y+0.42,3.6);
});

const gar=[
 ["R$ 200.000","parados em conta corrente","Rendimento zero. Metade do seu patrimônio perde poder de compra todos os dias, enquanto a outra metade assume sozinha todo o risco."],
 ["R$ 100.000","em ações, 25% do patrimônio total","Exposição direta à bolsa acima do que um perfil moderado sustenta com conforto, ainda mais sem tempo para acompanhar o mercado."]
];
gar.forEach(([v,t,d],i)=>{
  const y=2.0+i*1.85;
  s.addText(v,{x:4.9,y,w:4.2,h:0.55,fontSize:34,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
  s.addText(t,{x:4.9,y:y+0.58,w:7.7,h:0.3,fontSize:14.5,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:4.9,y:y+0.94,w:7.7,h:0.7,fontSize:12,color:G1,fontFace:F,lineSpacing:17,isTextBox:true,margin:0});
  if(i===0) rule(s,4.9,y+1.68,7.7);
});

rule(s,M,5.9,CW);
s.addText("Suas dores, nas suas palavras",{x:M,y:6.06,w:4,h:0.28,fontSize:11.5,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
s.addText("“Carteira sem plano definido”      “A oscilação incomoda”      “Sempre reclamando das taxas”",
  {x:M+4.2,y:6.06,w:7.9,h:0.28,fontSize:13,italic:true,color:W,fontFace:F,isTextBox:true,margin:0});
footer(s,2);
s.addNotes("Analisando seu momento: você construiu um patrimônio excelente de R$ 400 mil. Mas ele está partido ao meio de forma ineficiente. Metade parada na conta corrente rendendo zero, e metade concentrada em ações, R$ 100 mil, que representam 25% de tudo o que você tem. É por isso que a oscilação incomoda. Não é fragilidade sua, é dimensionamento. Vamos resolver os dois gargalos hoje.");

/* ---------- 3 · ATRITO ---------- */
s = newSlide(p);
header(s,"Reconquista","Você já foi assessorado pela XP","O que muda desta vez. Antes de falar de produto, precisamos falar do que não funcionou.");

s.addText("O que provavelmente faltou",{x:M,y:2.06,w:5.6,h:0.28,fontSize:11.5,bold:true,color:G2,fontFace:F,isTextBox:true,margin:0});
s.addText("O nosso compromisso",{x:6.9,y:2.06,w:5.6,h:0.28,fontSize:11.5,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
rule(s,M,2.44,CW,"4A4A4A");

const pares=[
 ["Carteira montada sem tese e sem objetivo nomeado","Cada real com finalidade e prazo declarados"],
 ["Produto oferecido sem o custo total na mesa","Custo total apresentado antes de cada alocação"],
 ["Contato apenas quando havia algo para vender","Agenda fixa: relatório trimestral e duas reuniões por ano"],
 ["Volatilidade sem contexto, gerando ansiedade","Risco dimensionado por caixa, não pela carteira inteira"]
];
pares.forEach(([a,b],i)=>{
  const y=2.64+i*0.86;
  s.addText(a,{x:M,y,w:5.6,h:0.58,fontSize:12.5,color:G1,fontFace:F,valign:"middle",isTextBox:true,margin:0});
  s.addText(b,{x:6.9,y,w:5.7,h:0.58,fontSize:12.5,bold:true,color:W,fontFace:F,valign:"middle",isTextBox:true,margin:0});
  rule(s,M,y+0.66,CW);
});

s.addText("Corrigir isso exige método, não uma prateleira diferente. É o método que estamos trazendo hoje.",
  {x:M,y:6.26,w:CW,h:0.36,fontSize:14.5,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
footer(s,3);
s.addNotes("Faço questão de começar por aqui. Você já foi nosso cliente assessorado e saiu, e isso é uma informação, não uma crítica. A dor que você descreve hoje, carteira sem plano definido, é o sintoma de um atendimento centrado em produto e não em objetivo. O que eu trago de diferente não é uma prateleira melhor: é método, custo na mesa antes da decisão e uma agenda de acompanhamento que existe independentemente de haver algo para vender.");

/* ---------- 4 · MÉTODO ---------- */
s = newSlide(p);
header(s,"Metodologia","Seu patrimônio dividido por objetivo, não por produto","Três caixas independentes, cada uma com prazo, risco e função próprios.");

const caixas=[
 ["01","Liquidez","R$ 60.000","15%","Curto prazo","Reserva de emergência com resgate no mesmo dia e volatilidade zero."],
 ["02","Proteção e futuro","R$ 140.000","35%","Médio e longo prazo","Educação da filha, eficiência fiscal e sucessão sem inventário."],
 ["03","Crescimento","R$ 200.000","50%","Longo prazo","Rentabilidade real acima da inflação, com risco dimensionado."]
];
rule(s,M,2.1,CW,"4A4A4A");
caixas.forEach(([n,t,v,pc,pz,d],i)=>{
  const x=M+i*4.0;
  s.addText(n,{x,y:2.3,w:1,h:0.3,fontSize:12,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
  s.addText(t,{x,y:2.62,w:3.5,h:0.34,fontSize:17,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(v,{x,y:3.12,w:3.5,h:0.62,fontSize:33,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(pc+" do patrimônio",{x,y:3.76,w:3.5,h:0.28,fontSize:11.5,color:YEL,fontFace:F,isTextBox:true,margin:0});
  rule(s,x,4.16,3.5);
  s.addText(pz,{x,y:4.3,w:3.5,h:0.28,fontSize:11.5,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x,y:4.62,w:3.5,h:1.1,fontSize:12,color:G1,fontFace:F,lineSpacing:17,isTextBox:true,margin:0});
});

block(s,M,5.9,CW,0.92);
s.addText([
 {text:"Origem dos recursos.  ",options:{bold:true,color:YEL}},
 {text:"Os R$ 200.000 da conta corrente financiam as Caixas 1 e 2. Os R$ 200.000 já investidos são reestruturados na Caixa 3.\n",options:{color:G1}},
 {text:"Premissa declarada.  ",options:{bold:true,color:W}},
 {text:"O briefing não informa despesas, então assumimos custo de vida de R$ 8.400 por mês, 70% da renda.",options:{color:G1}}
],{x:M+0.3,y:5.9,w:CW-0.6,h:0.92,fontSize:11.5,fontFace:F,valign:"middle",lineSpacing:17,isTextBox:true,margin:0});
footer(s,4);
s.addNotes("Para eliminar a ansiedade com oscilação, dividimos o patrimônio em três caixas independentes. O dinheiro de emergência fica imune a mercado. O capital da sua filha fica protegido da inflação e blindado na sucessão. E só o dinheiro de longo prazo assume risco. Repare numa coisa importante: eu declaro a premissa. O briefing não me deu suas despesas, então assumi 70% da renda. Se esse número for outro, a Caixa 1 muda e o resto se ajusta.");

/* ---------- 5 · CAIXA 1 ---------- */
s = newSlide(p);
header(s,"Caixa 01  ·  Curto prazo","Liquidez imediata e reserva de emergência","O primeiro dinheiro a sair da conta corrente e o único que nunca oscila.");

s.addText("R$ 60.000",{x:M,y:2.1,w:4.1,h:0.85,fontSize:46,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
s.addText("15% do patrimônio total",{x:M,y:2.96,w:4.1,h:0.3,fontSize:13,color:G1,fontFace:F,isTextBox:true,margin:0});
rule(s,M,3.5,4.1);
s.addText("Como chegamos ao valor",{x:M,y:3.66,w:4.1,h:0.28,fontSize:11.5,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
const conta=[["Custo de vida estimado","R$ 8.400 / mês"],["Cobertura alvo","6 meses"],["Resultado","R$ 50.400"],["Arredondado para","R$ 60.000, ou 7,1 meses"]];
conta.forEach(([k,v],i)=>{
  const y=4.06+i*0.56;
  s.addText(k,{x:M,y,w:2.1,h:0.28,fontSize:11.5,color:G2,fontFace:F,isTextBox:true,margin:0});
  s.addText(v,{x:M+2.1,y,w:2.0,h:0.28,fontSize:11.5,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
});

rule(s,5.5,2.1,7.1,"4A4A4A");
const at=[
 ["CDB de liquidez diária D+0","Bancos de primeira linha na plataforma XP, faixa de 100% a 103% do CDI, com resgate no mesmo dia pelo aplicativo."],
 ["Fundo DI com taxa zero de administração","Alternativa sem taxa de administração para a parcela de giro, com liquidez no mesmo dia útil."],
 ["Cobertura do FGC","Até R$ 250 mil por CPF por instituição, com teto global de R$ 1 milhão a cada quatro anos. Pulverizando em dois emissores, sua exposição fica muito abaixo do limite."]
];
at.forEach(([t,d],i)=>{
  const y=2.26+i*1.42;
  s.addText(t,{x:5.5,y,w:7.1,h:0.34,fontSize:16,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:5.5,y:y+0.4,w:7.1,h:0.72,fontSize:12,color:G1,fontFace:F,lineSpacing:17,valign:"top",isTextBox:true,margin:0});
  if(i<2) rule(s,5.5,y+1.22,7.1);
});

block(s,M,6.36,CW,0.46);
s.addText("Volatilidade desta caixa: zero. Este dinheiro não participa de mercado. Ele existe para que você nunca precise resgatar as outras duas caixas num momento ruim.",
  {x:M+0.3,y:6.36,w:CW-0.6,h:0.46,fontSize:12,color:W,fontFace:F,valign:"middle",isTextBox:true,margin:0});
footer(s,5);
s.addNotes("Sua reserva sai da conta corrente e vai para CDB de liquidez diária com garantia do FGC. Rende todos os dias e resgata no mesmo dia pelo aplicativo. Note que eu não inventei o número: parti de uma estimativa de custo de vida e mostrei a conta. Se o seu gasto real for maior, a gente recalcula aqui mesmo. E essa caixa tem uma função silenciosa que é a mais importante de todas: ela impede que você precise vender um investimento de longo prazo num momento ruim.");

/* ---------- 6 · CAIXA 2 ---------- */
s = newSlide(p);
header(s,"Caixa 02  ·  Médio e longo prazo","Proteção da família e o futuro da sua filha","R$ 140.000 alocados, mais a camada de proteção que faltava no seu planejamento.");

rule(s,M,2.1,CW,"4A4A4A");
const c2=[
 ["VGBL","R$ 80.000","Aporte único","Tabela regressiva de IR, com alíquota de 10% a partir de dez anos e incidência apenas sobre o rendimento. Sucessão direta à sua filha, sem inventário e sem custo de partilha."],
 ["LCI e LCA","R$ 60.000","Isentos de IR","Isenção total de Imposto de Renda para pessoa física e cobertura do FGC. Vencimentos escalonados e casados com os marcos escolares da sua filha. Carência mínima de nove meses."],
 ["Seguro de vida","R$ 720.000","Capital segurado","Cinco vezes a sua renda anual, custeado pelo fluxo mensal e não pelo principal. Estimativa a partir de R$ 250 por mês, sujeita à análise de risco da seguradora."]
];
c2.forEach(([t,v,sub,d],i)=>{
  const x=M+i*4.0;
  s.addText(t,{x,y:2.3,w:3.5,h:0.32,fontSize:16.5,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
  s.addText(v,{x,y:2.76,w:3.5,h:0.58,fontSize:31,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(sub,{x,y:3.36,w:3.5,h:0.28,fontSize:11.5,color:G2,fontFace:F,isTextBox:true,margin:0});
  rule(s,x,3.76,3.5);
  s.addText(d,{x,y:3.92,w:3.5,h:1.5,fontSize:12,color:G1,fontFace:F,lineSpacing:17,isTextBox:true,margin:0});
});

block(s,M,5.62,CW,1.2);
s.addText("Por que VGBL e não PGBL no aporte único",{x:M+0.3,y:5.82,w:11,h:0.28,fontSize:12,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
s.addText("A dedução do PGBL é limitada a 12% da renda bruta tributável anual, o que dá R$ 17.280 no seu caso. Um aporte único de R$ 80 mil em PGBL desperdiçaria o benefício e ainda seria tributado sobre o valor total no resgate. Por isso: VGBL no aporte único, e PGBL como aporte mensal de R$ 1.440, exatamente dentro do teto, caso você declare no modelo completo.",
  {x:M+0.3,y:6.12,w:11.5,h:0.62,fontSize:11.5,color:G1,fontFace:F,lineSpacing:16,isTextBox:true,margin:0});
footer(s,6);
s.addNotes("Para o futuro da sua filha, três camadas. VGBL para o aporte único, com alíquota caindo a 10% no longo prazo e transmissão direta para ela sem passar por inventário. LCI e LCA isentas de Imposto de Renda, com vencimentos casados com a escola e a faculdade. E aqui eu chamo sua atenção: eu incluí seguro de vida. Você tem 40 anos, renda única e uma filha dependente. Todo o plano que montamos pressupõe que você continue gerando R$ 12 mil por mês. O seguro é o que protege o plano contra a única hipótese que ele não sobrevive.");

/* ---------- 7 · CAIXA 3 ---------- */
s = newSlide(p);
header(s,"Caixa 03  ·  Longo prazo","Crescimento com risco dimensionado","R$ 200.000 reestruturados a partir da carteira que você já tem hoje.");

rule(s,M,2.1,CW,"4A4A4A");
const c3=[
 ["55%","R$ 110.000","Renda fixa de carrego","Tesouro IPCA+ e debêntures incentivadas pela Lei 12.431, isentas de IR para pessoa física. Levadas ao vencimento, com o prazo casado ao objetivo, a marcação a mercado deixa de ser um problema."],
 ["25%","R$ 50.000","Multimercados de baixa volatilidade","Gestão ativa da prateleira XP, com mandato de baixa volatilidade e alvo de retorno acima do CDI. Diversificação de estratégia sem exigir acompanhamento seu."],
 ["20%","R$ 40.000","Renda variável via ETFs","BOVA11 e IVVB11 dão bolsa brasileira e mercado americano em duas ordens. Diversificação automática, sem escolher ação nem acompanhar notícia."]
];
c3.forEach(([pc,v,t,d],i)=>{
  const y=2.3+i*1.34;
  s.addText(pc,{x:M,y,w:1.3,h:0.5,fontSize:30,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
  s.addText(v,{x:M+1.4,y:y+0.08,w:2.0,h:0.34,fontSize:16,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(t,{x:M+3.5,y:y+0.04,w:8.4,h:0.32,fontSize:15.5,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:M+3.5,y:y+0.42,w:8.4,h:0.7,fontSize:12,color:G1,fontFace:F,lineSpacing:17,valign:"top",isTextBox:true,margin:0});
  rule(s,M,y+1.18,CW);
});

s.addText("Por que não indicamos COE nesta carteira",{x:M,y:6.26,w:CW,h:0.28,fontSize:12,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
s.addText("O custo de estruturação de um COE não aparece destacado no extrato. Como transparência de taxas é justamente o seu ponto sensível, indicá-lo seria incoerente. A proteção de capital vem da renda fixa de carrego, onde cada custo é declarado.",
  {x:M,y:6.54,w:CW,h:0.32,fontSize:11.5,color:G1,fontFace:F,isTextBox:true,margin:0});
footer(s,7);
s.addNotes("Aqui está a mudança mais importante. Sua exposição a ações cai de R$ 100 mil para R$ 40 mil, e sai de ação individual para ETF, com diversificação automática. A maior fatia vai para renda fixa levada ao vencimento, boa parte dela isenta de IR. E eu quero ser explícito sobre uma ausência: não tem COE aqui. COE é um produto legítimo, mas o custo de estruturação não fica visível no extrato. Como taxa é o seu ponto sensível, eu não vou te oferecer justamente o produto mais opaco da prateleira.");

/* ---------- 8 · ANTES x DEPOIS ---------- */
s = newSlide(p);
header(s,"Consolidação","Antes e depois, na mesma base de comparação","Todos os percentuais se referem ao patrimônio total de R$ 400.000.");

const LB=["Ações / ETF","Multimercado","RF tributada","Previdência (VGBL)","Isentos de IR","Liquidez / pós-fixado","Parado em conta"];
s.addChart("bar",[
 {name:"Hoje", labels:LB, values:[100,0,100,0,0,0,200]},
 {name:"Proposto", labels:LB, values:[40,50,50,80,120,60,0]}
],{
 x:M-0.1, y:2.05, w:7.7, h:4.55, barDir:"bar", barGapWidthPct:40,
 chartColors:["4A4A4A",YEL], showValue:true, dataLabelPosition:"outEnd",
 dataLabelColor:G1, dataLabelFontSize:9.5, dataLabelFontFace:F, dataLabelFormatCode:'#,##0"k"',
 showLegend:true, legendPos:"t", legendColor:G1, legendFontSize:11, legendFontFace:F,
 catAxisLabelColor:W, catAxisLabelFontSize:10.5, catAxisLabelFontFace:F, catAxisLineShow:false,
 valAxisLabelColor:G2, valAxisLabelFontSize:9.5, valAxisLabelFontFace:F, valAxisLineShow:false,
 valAxisMinVal:0, valAxisMaxVal:210, valAxisMajorUnit:50,
 valGridLine:{color:"1F1F1F",size:1}, catGridLine:{style:"none"},
 plotArea:{fill:{color:BG}}, chartArea:{fill:{color:BG}}
});
s.addText("Valores em R$ mil",{x:M,y:6.62,w:7.7,h:0.26,fontSize:9.5,color:G2,fontFace:F,isTextBox:true,margin:0});

rule(s,8.6,2.05,4.0,"4A4A4A");
const kpis=[
 ["Exposição a ações","25%  →  10%","R$ 100.000 para R$ 40.000, agora via ETF"],
 ["Capital rendendo","50%  →  100%","nenhum real parado em conta corrente"],
 ["Ativos isentos de IR","R$ 0  →  R$ 120.000","LCI, LCA e debêntures incentivadas"],
 ["Taxa de custódia","R$ 0","renda fixa, fundos e previdência"]
];
kpis.forEach(([t,v,d],i)=>{
  const y=2.22+i*1.15;
  s.addText(t,{x:8.6,y,w:4.0,h:0.26,fontSize:11,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
  s.addText(v,{x:8.6,y:y+0.28,w:4.0,h:0.38,fontSize:20,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:8.6,y:y+0.68,w:4.0,h:0.26,fontSize:10,color:G1,fontFace:F,isTextBox:true,margin:0});
  rule(s,8.6,y+1.0,4.0);
});
footer(s,8);
s.addNotes("Esta é a fotografia final, e repare que eu comparo tudo na mesma base: o patrimônio total de R$ 400 mil. Sua exposição a ações cai de 25% para 10%. O capital que está de fato rendendo sai de 50% para 100%. E surgem R$ 120 mil em ativos isentos de Imposto de Renda, que hoje são zero. Nenhum real fica parado.");

/* ---------- 9 · CUSTO E TEMPO ---------- */
s = newSlide(p);
header(s,"O valor do serviço","Custo transparente e tempo preservado","As duas dores que sobraram: a reclamação com taxas e a falta de tempo para acompanhar.");

rule(s,M,2.1,CW,"4A4A4A");
s.addText("Transparência de custos",{x:M,y:2.3,w:5.6,h:0.3,fontSize:16.5,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
const col1=[
 ["Taxa zero de custódia","em renda fixa, fundos e previdência na plataforma XP."],
 ["Renda fixa","remunera pelo spread do emissor, apresentado a você antes da alocação e não depois."],
 ["Fundos e previdência","com taxa de administração na lâmina. Nenhum produto com custo embutido e não destacado."],
 ["Research XP","analisa rating e saúde financeira de cada emissor antes de ele entrar na sua carteira."]
];
col1.forEach(([t,d],i)=>{
  const y=2.78+i*0.92;
  s.addText([{text:t+". ",options:{bold:true,color:W}},{text:d,options:{color:G1}}],
    {x:M,y,w:5.6,h:0.8,fontSize:12,fontFace:F,lineSpacing:17,valign:"top",isTextBox:true,margin:0});
  if(i<3) rule(s,M,y+0.84,5.6);
});

s.addText("O seu tempo, preservado",{x:6.9,y:2.3,w:5.7,h:0.3,fontSize:16.5,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
const col2=[
 ["Monitoramento diário","feito pela nossa equipe. Você não precisa acompanhar notícia de mercado."],
 ["Relatório trimestral","consolidado e resumido no aplicativo da XP."],
 ["Duas reuniões por ano","de 30 minutos, para revisar metas familiares e reenquadrar o perfil."],
 ["Rebalanceamento por aporte","e não por venda, evitando antecipar Imposto de Renda sem necessidade."]
];
col2.forEach(([t,d],i)=>{
  const y=2.78+i*0.92;
  s.addText([{text:t+". ",options:{bold:true,color:W}},{text:d,options:{color:G1}}],
    {x:6.9,y,w:5.7,h:0.8,fontSize:12,fontFace:F,lineSpacing:17,valign:"top",isTextBox:true,margin:0});
  if(i<3) rule(s,6.9,y+0.84,5.7);
});

block(s,M,6.3,CW,0.52);
s.addText([
 {text:"Capacidade de aporte mensal.  ",options:{bold:true,color:YEL}},
 {text:"R$ 3.600, sendo R$ 1.440 no PGBL dentro do teto de 12%, cerca de R$ 250 no seguro de vida e R$ 1.910 para rebalanceamento.",options:{color:W}}
],{x:M+0.3,y:6.3,w:CW-0.6,h:0.52,fontSize:11.5,fontFace:F,valign:"middle",isTextBox:true,margin:0});
footer(s,9);
s.addNotes("As duas dores que faltavam. Taxa: custódia zero, e o custo de cada produto na mesa antes da decisão, não depois. Tempo: o acompanhamento diário é nosso. Você recebe um resumo trimestral no app e me dá duas reuniões de 30 minutos por ano. E o rebalanceamento acontece direcionando seu aporte mensal, o que evita venda e evita antecipar imposto. Sua sobra de R$ 3.600 por mês já tem destino definido.");

/* ---------- 10 · PLANO DE AÇÃO ---------- */
s = newSlide(p);
header(s,"Próximos passos","Plano de ação imediato","Quatro movimentos. O primeiro começa ainda nesta reunião.");

rule(s,M,2.1,CW,"4A4A4A");
const steps=[
 ["01","Hoje","Atualizar o seu suitability na conta XP e validar a proposta de alocação das três caixas."],
 ["02","Esta semana","Aplicar os R$ 200.000 da conta corrente: R$ 60 mil na reserva em CDB D+0 e R$ 140 mil em VGBL, LCI e LCA."],
 ["03","Próximos 5 meses","Desmontar a posição em ações em parcelas de até R$ 20 mil por mês, aproveitando a isenção de IR sobre vendas de ações até esse limite mensal."],
 ["04","Contínuo","Aporte mensal de R$ 3.600, relatório trimestral no aplicativo e revisão semestral de 30 minutos."]
];
steps.forEach(([n,q,d],i)=>{
  const y=2.3+i*1.06;
  s.addText(n,{x:M,y:y+0.04,w:0.8,h:0.34,fontSize:14,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
  s.addText(q,{x:M+0.85,y:y+0.02,w:2.6,h:0.34,fontSize:15.5,bold:true,color:W,fontFace:F,isTextBox:true,margin:0});
  s.addText(d,{x:M+3.5,y:y+0.04,w:8.4,h:0.72,fontSize:12.5,color:G1,fontFace:F,lineSpacing:17,valign:"top",isTextBox:true,margin:0});
  rule(s,M,y+0.88,CW);
});

s.addText("Atualizar o suitability leva 10 minutos. Seus R$ 200 mil param de perder para a inflação já no próximo dia útil.",
  {x:M,y:6.44,w:CW,h:0.36,fontSize:15,bold:true,color:YEL,fontFace:F,isTextBox:true,margin:0});
footer(s,10);
s.addNotes("Para colocar o plano em prática, o primeiro passo é atualizar seu perfil de investidor, o que leva 10 minutos e dá para fazer agora. Chamo atenção para a etapa 3: a saída das ações é gradual, em parcelas de até R$ 20 mil por mês, porque vendas de ações até esse valor mensal são isentas de Imposto de Renda para pessoa física. Fazendo assim, a reestruturação inteira sai sem custo tributário. Posso abrir o aplicativo e começarmos?");

p.writeFile({ fileName: "Case_XP_Planejamento_Financeiro.pptx" }).then(()=>console.log("deck gerado"));
