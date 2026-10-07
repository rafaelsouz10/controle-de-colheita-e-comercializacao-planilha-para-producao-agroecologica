// Constantes, paleta e helpers de estilo usados pelas demais rotinas.

var ABA_GERAL = 'Geral';
var ABA_PRODUTOS = 'Produtos';
var ABA_LEIAME = 'Leia-me';
var ABA_CONFIG = 'Configuração';

// Campos da aba Configuracao, na ordem em que aparecem (a partir da linha 4, col B).
var CAMPOS_CONFIG = [
  ['nome', 'Agricultor(a)'],
  ['nucleo', 'Núcleo'],
  ['grupo', 'Grupo'],
  ['rede', 'Rede / associação'],
  ['ano', 'Ano da planilha']
];
var LINHA_CONFIG = 4;

// Referencia para usar DENTRO de formula (ex.: o subtitulo da Geral). Formula, e nao
// valor copiado, para que uma correcao feita pelo produtor offline apareca na hora.
function refConfig(chave) {
  for (var i = 0; i < CAMPOS_CONFIG.length; i++) {
    if (CAMPOS_CONFIG[i][0] === chave) return "'" + ABA_CONFIG + "'!$B$" + (LINHA_CONFIG + i);
  }
  throw new Error('campo de configuracao desconhecido: ' + chave);
}
var ABA_CONSOLIDADO = '_Consolidado'; // oculta: empilha as abas de mes

// Layout da aba de mes: dashboard em cima, registro embaixo.
var LINHA_CAB = 16;   // cabecalho da tabela de registro
var LINHA_DADOS = 17; // primeira linha digitavel
var COL_AUX = 16;     // coluna P: area auxiliar (oculta) que alimenta os graficos

var MESES = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun',
             'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];

var COR = {
  escuro: '#2E5E3A',   // cabecalhos
  medio: '#7A9E3F',    // destaques
  claro: '#EAF1E4',    // fundo de cards
  fundo: '#F7FAF5',    // fundo do dashboard
  texto: '#1F2D22',
  suave: '#6B7A6B',    // textos secundarios
  alerta: '#FBE9E7',   // fundo de celula suspeita
  borda: '#C9D8C0'
};

var MESES_LONGOS = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

// A planilha e de UM ano (o da aba Configuracao), entao a aba se chama so "Ago".
// O ano sai de um lugar so; o titulo de cada aba o mostra por formula.
function nomeAba(mes) {
  return MESES[mes - 1];
}

// Reconhece "Ago" e tambem o formato antigo "Ago 2026", que existe nas planilhas
// montadas antes desta mudanca, ate a migracao renomea-las (migrarNomesDasAbas).
var RE_ABA_MES = /^(Jan|Fev|Mar|Abr|Mai|Jun|Jul|Ago|Set|Out|Nov|Dez)( (\d{4}))?$/;

function ehAbaDeMes(nome) {
  return RE_ABA_MES.test(nome);
}

function anoDaPlanilha() {
  var cfg = lerConfig();
  var ano = cfg && Number(cfg.ano);
  return ano >= 2000 && ano <= 2100 ? ano : new Date().getFullYear();
}

// Abas de mes, em ordem de janeiro a dezembro. Base do consolidado e da Geral.
function abasDeMes() {
  var ano = anoDaPlanilha();
  return SpreadsheetApp.getActive().getSheets()
    .map(function (s) {
      var m = RE_ABA_MES.exec(s.getName());
      return m ? { aba: s, ano: m[3] ? Number(m[3]) : ano, mes: MESES.indexOf(m[1]) + 1 } : null;
    })
    .filter(function (x) { return x; })
    .sort(function (a, b) { return (a.ano - b.ano) || (a.mes - b.mes); });
}

// Renomeia "Ago 2026" para "Ago". So renomeia: nao apaga nem recria, entao os
// registros ficam onde estao. Sem isto, criarAbasDoAno nao acharia a aba antiga
// pelo nome novo e criaria uma "Ago" vazia ao lado dela. Rodar de novo nao faz mal.
function migrarNomesDasAbas() {
  var ss = SpreadsheetApp.getActive();
  var conflitos = [];
  abasDeMes().forEach(function (m) {
    var novo = nomeAba(m.mes);
    if (m.aba.getName() === novo) return;
    if (ss.getSheetByName(novo)) { conflitos.push(m.aba.getName()); return; }
    m.aba.setName(novo);
  });
  if (conflitos.length) {
    SpreadsheetApp.getUi().alert('Não renomeei ' + conflitos.join(', ')
      + ': já existe uma aba com o nome novo. Confira qual das duas tem os registros.');
  }
}

// Titulo de secao dentro de um dashboard.
function tituloSecao(aba, linha, col, largura, texto) {
  var r = aba.getRange(linha, col, 1, largura).merge();
  r.setValue(texto)
    .setFontSize(11).setFontWeight('bold').setFontColor(COR.escuro)
    .setVerticalAlignment('middle');
  aba.setRowHeight(linha, 26);
  return r;
}

// Card de indicador: rotulo em cima, numero grande embaixo.
function card(aba, linha, col, largura, rotulo, formula, formato) {
  aba.getRange(linha, col, 1, largura).merge()
    .setValue(rotulo)
    .setFontSize(9).setFontColor(COR.suave).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground(COR.claro);
  aba.getRange(linha + 1, col, 1, largura).merge()
    .setFormula(formula)
    .setFontSize(20).setFontWeight('bold').setFontColor(COR.escuro)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(COR.claro).setNumberFormat(formato || '0');
  aba.setRowHeight(linha + 1, 34);
  aba.getRange(linha, col, 2, largura)
    .setBorder(true, true, true, true, false, false, COR.borda, SpreadsheetApp.BorderStyle.SOLID);
}

// Cabecalho de uma tabelinha do dashboard (ex.: Produto | Pacotes).
function cabecalhoTabela(aba, linha, col, rotulos) {
  aba.getRange(linha, col, 1, rotulos.length).setValues([rotulos])
    .setFontSize(9).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground(COR.medio);
}

// Quantas linhas de uma coluna estao realmente preenchidas. Chamar SEMPRE depois de
// um flush, porque o valor vem de formula.
//
// Existe por causa de um bug que custou caro: os graficos nasciam com faixa fixa e
// larga (P3:Q40) para caber meses futuros. Com 36 das 38 linhas vazias, o Sheets
// erra a deducao de qual coluna e rotulo e qual e serie, e monta o grafico SEM SERIE
// NENHUMA - a tela em branco com "Adicione uma serie". O mesmo grafico feito a mao,
// com a faixa justa, funciona. Entao a faixa passou a ser medida, nao chutada.
function linhasComDado(aba, linha, col, maximo) {
  var v = aba.getRange(linha, col, maximo, 1).getValues();
  var n = 0;
  for (var i = 0; i < v.length; i++) {
    if (v[i][0] !== '' && v[i][0] !== null) n++;
  }
  return n;
}

function limparAba(aba) {
  aba.clear();
  aba.getCharts().forEach(function (c) { aba.removeChart(c); });
  var f = aba.getFilter();
  if (f) f.remove();
  aba.getRange(1, 1, aba.getMaxRows(), aba.getMaxColumns())
    .breakApart().clearDataValidations().clearNote();
  aba.setConditionalFormatRules([]);
  if (aba.getFrozenRows()) aba.setFrozenRows(0);
}


// O Sheets interpreta a formula no idioma da planilha: em pt_BR o separador de
// argumentos e ';', e tudo que o script escreve usa ','. Converter na mao seria
// arriscado (as QUERY tem virgula dentro do texto), entao escrevemos com a planilha
// em ingles e devolvemos o idioma no fim: formula ja aceita nao depende mais dele.
function comLocaleNeutro(fn) {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetLocale('en_US');
  SpreadsheetApp.flush();
  try {
    fn();
  } finally {
    SpreadsheetApp.flush();
    ss.setSpreadsheetLocale('pt_BR');
  }
}

// Erro de sintaxe vira #ERROR! e NAO e capturado por IFERROR, entao um painel
// quebrado passaria despercebido. Conferimos uma celula e avisamos na hora.
function conferirFormulas() {
  var v = SpreadsheetApp.getActive().getSheetByName(ABA_GERAL)
    .getRange(4, 1).getDisplayValue();
  return v.charAt(0) === '#'
    ? 'ATENÇÃO: o painel está mostrando ' + v
      + '. As fórmulas não foram aceitas.'
    : 'Painel conferido: ' + v + ' pacotes no total.';
}

// --- Modelo x caderno importado ----------------------------------------------
// O Dados.gs e o caderno de UMA produtora. O modelo tem que funcionar sem ele:
// para gerar a planilha de outro produtor, faz-se uma copia e apaga-se o Dados.gs.
// Por isso nada aqui fora referencia DADOS & cia. diretamente, so por caderno().

// Lista inicial para uma planilha em branco: hortalicas e frutas comuns na
// producao agroecologica, pelo nome comum. E ponto de partida, nao regra: o
// produtor desativa o que nao planta e cadastra o que falta.
var PRODUTOS_BASE = [
  ['Abóbora', 'pacote'], ['Aipo', 'pacote'], ['Alecrim', 'pacote'],
  ['Alface', 'pacote'], ['Alho-poró', 'unidade'], ['Almeirão', 'pacote'],
  ['Beterraba', 'pacote'], ['Caqui', 'unidade'], ['Cebolinha', 'pacote'],
  ['Cenoura', 'pacote'], ['Cheiro-verde', 'pacote'], ['Chicória', 'pacote'],
  ['Coentro', 'pacote'], ['Couve', 'pacote'], ['Espinafre', 'pacote'],
  ['Giló', 'pacote'], ['Hortelã', 'pacote'], ['Limão', 'pacote'],
  ['Mamão', 'unidade'], ['Manjericão', 'pacote'], ['Maxixe', 'pacote'],
  ['Ora-pro-nóbis', 'pacote'], ['Pepino', 'pacote'], ['Pimentão', 'pacote'],
  ['Rúcula', 'pacote'], ['Salsa', 'pacote'], ['Tomate', 'pacote'],
  ['Tomatinho', 'pacote']
];

// Planilha de bancada: a que tem um caderno real transcrito, usada para testar
// com dados. So nela aparece o submenu Teste (importar caderno, diagnosticos).
//
// Ela se identifica por uma PROPRIEDADE DO SCRIPT, nao por um ID escrito aqui: o
// repositorio e publico e nao leva nada que aponte para os dados de alguem. No
// editor da planilha de bancada: Configuracoes do projeto > Propriedades do script
// > ID_PLANILHA_TESTE = <ID dessa planilha>. No modelo e nas copias dos produtores
// a propriedade nao existe (ou aponta para outra planilha), e o submenu nao aparece.
function ehPlanilhaTeste() {
  try {
    var id = PropertiesService.getScriptProperties().getProperty('ID_PLANILHA_TESTE');
    return !!id && id === SpreadsheetApp.getActive().getId();
  } catch (e) {
    return false;   // na duvida, sem submenu de teste
  }
}

// O caderno importado, se o Dados.gs existir e tiver registros; senao, null.
function caderno() {
  if (typeof DADOS === 'undefined' || !DADOS.length) return null;
  return {
    registros: DADOS,
    produtos: PRODUTOS_INICIAIS.map(function (p) { return [p[0], p[1], 'sim', '']; }),
    destinos: DESTINOS_INICIAIS.map(function (d) { return [d, 'sim']; }),
    config: CONFIG_INICIAL
  };
}

function configVazia() {
  return { nome: '', nucleo: '', grupo: '', rede: '', ano: new Date().getFullYear() };
}
