// Menu da planilha e as rotinas que montam tudo.
//
// Dois jeitos de montar, os dois APAGAM a planilha e pedem confirmacao:
//   - novaPlanilha():    em branco, para um produtor novo.
//   - importarCaderno(): com os registros do Dados.gs. Fica no submenu Teste, que
//                        so existe na planilha de bancada (ver ehPlanilhaTeste),
//                        para nenhum produtor sobrescrever os proprios registros
//                        com o caderno de outra pessoa.

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Uso do dia a dia em cima; o que reconstroi ou apaga fica um nivel abaixo, em
  // Manutencao, longe do clique casual.
  var manutencao = ui.createMenu('Manutenção')
    .addItem('Recriar meses que faltam', 'criarAbasDoAno')
    .addItem('Aplicar estilo (mantém os dados)', 'estilizar')
    .addSeparator()
    .addItem('Nova planilha em branco (apaga tudo)', 'novaPlanilha');
  var menu = ui.createMenu('Colheita')
    .addItem('Atualizar resumo geral', 'atualizarGeral')
    .addSeparator()
    .addSubMenu(manutencao);
  if (ehPlanilhaTeste()) {
    var teste = ui.createMenu('Teste');
    if (caderno()) teste.addItem('Importar caderno transcrito (apaga tudo)', 'importarCaderno');
    teste.addSeparator()
      .addItem('Diagnosticar opções de gráfico', 'diagnosticarOpcoes')
      .addItem('Limpar gráficos de diagnóstico', 'limparTestes');
    menu.addSeparator().addSubMenu(teste);
  }
  menu.addToUi();
  irParaMesAtual();
}

// A planilha e de registro: quem abre, abre para anotar a colheita de hoje.
// So conveniencia - no app do celular e offline o onOpen nao roda, e tudo bem.
function irParaMesAtual() {
  var hoje = new Date();
  var aba = SpreadsheetApp.getActive()
    .getSheetByName(nomeAba(hoje.getMonth() + 1));
  if (aba) SpreadsheetApp.getActive().setActiveSheet(aba);
}

// Os doze meses de um ano civil. A producao agroecologica nao tem entressafra:
// ha colheita o ano todo, entao a planilha cobre janeiro a dezembro.
function mesesDoAno(ano) {
  var fora = [];
  for (var m = 1; m <= 12; m++) fora.push({ ano: ano, mes: m, registros: [] });
  return fora;
}

function novaPlanilha() {
  if (!confirmarApagar('Montar uma planilha EM BRANCO')) return;
  var cfg = configVazia();
  montarTudo({
    config: cfg,
    produtos: null,            // lista base do modelo
    destinos: [],
    meses: mesesDoAno(cfg.ano)
  });
  irParaMesAtual();
  SpreadsheetApp.getUi().alert('Planilha em branco pronta, de janeiro a dezembro de '
    + cfg.ano + '. Comece pela aba Configuração (nome, núcleo e grupo). '
    + conferirFormulas());
}

function importarCaderno() {
  if (!ehPlanilhaTeste()) {
    SpreadsheetApp.getUi().alert('Importar caderno só funciona na planilha de teste. '
      + 'Nesta planilha, use "Nova planilha em branco".');
    return;
  }
  var c = caderno();
  if (!c) {
    SpreadsheetApp.getUi().alert('Não há caderno para importar (o Dados.gs está vazio ou foi apagado).');
    return;
  }
  var anos = {};
  c.registros.forEach(function (r) { anos[r[0]] = true; });
  if (Object.keys(anos).length > 1) {
    SpreadsheetApp.getUi().alert('O caderno tem registros de mais de um ano ('
      + Object.keys(anos).join(', ') + '). Cada planilha é de um ano só.');
    return;
  }
  if (!confirmarApagar('Importar os ' + c.registros.length + ' registros do caderno transcrito')) return;

  // Cada ano que aparece no caderno vira janeiro a dezembro completos; os meses
  // com registro recebem os seus, os outros nascem vazios.
  var porMes = {};
  c.registros.forEach(function (r) {
    if (!porMes[r[0] * 100 + 1]) {
      mesesDoAno(r[0]).forEach(function (m) { porMes[m.ano * 100 + m.mes] = m; });
    }
    porMes[r[0] * 100 + r[1]].registros.push(r);
  });
  montarTudo({
    config: c.config,
    produtos: c.produtos,
    destinos: c.destinos,
    meses: Object.keys(porMes).sort().map(function (k) { return porMes[k]; })
  });
  irParaMesAtual();
  SpreadsheetApp.getUi().alert('Planilha montada com ' + c.registros.length
    + ' registros do caderno de campo. ' + conferirFormulas());
}

// Pergunta antes de apagar. Uma planilha de produtor tem meses de registros que
// nao existem em lugar nenhum mais: um clique distraido nao pode custar isso.
function confirmarApagar(oQue) {
  var ui = SpreadsheetApp.getUi();

  // Com registros, Sim/Nao e facil demais: um clique distraido apaga um ano de
  // colheita. Exige digitar APAGAR, que ninguem faz por engano.
  var n = abasDeMes().reduce(function (s, m) { return s + lerRegistros(m.aba).length; }, 0);
  if (n > 0) {
    var p = ui.prompt('Atenção: isto apaga ' + n + ' registros de colheita',
      oQue + ' vai APAGAR todas as abas e os ' + n + ' registros digitados. '
      + 'Pelo menu não há volta (só pelo Histórico de versões do Google).\n\n'
      + 'Para só arrumar a aparência sem perder nada, use "Aplicar estilo".\n\n'
      + 'Para continuar, digite APAGAR:', ui.ButtonSet.OK_CANCEL);
    if (p.getSelectedButton() === ui.Button.OK
        && p.getResponseText().trim().toUpperCase() === 'APAGAR') return true;
    ui.alert('Nada foi apagado.');
    return false;
  }

  var r = ui.alert('Atenção: isto apaga a planilha inteira',
    oQue + ' vai APAGAR todas as abas e todos os registros digitados.\n\n'
    + 'Para só arrumar a aparência sem perder nada, use "Aplicar estilo".\n\n'
    + 'Continuar mesmo assim?', ui.ButtonSet.YES_NO);
  return r === ui.Button.YES;
}

// Monta tudo do zero. Remove as abas de mes que existirem antes, para nao sobrar
// mes antigo de uma montagem anterior misturado com a nova.
function montarTudo(o) {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('America/Bahia');
  abasDeMes().forEach(function (m) { ss.deleteSheet(m.aba); });

  comLocaleNeutro(function () {
    // A aba de apoio nasce primeiro (ainda vazia): Produtos e Geral apontam para
    // ela, e formula que nasce apontando para aba inexistente vira #REF incuravel.
    montarConsolidado();
    montarConfig(o.config);
    montarProdutos(o.produtos, o.destinos);
    montarLeiaMe();
    o.meses.forEach(function (m) { montarMes(m.ano, m.mes, m.registros); });
    montarConsolidado();
    montarGeral();
  });
  ordenarAbas();

  var padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 1) ss.deleteSheet(padrao);
}

// Cria os meses que faltam no ano da Configuracao, sem tocar nos que ja existem
// (e nos registros deles). Rodar de novo nao faz mal: se nada falta, nada muda.
function criarAbasDoAno() {
  var ui = SpreadsheetApp.getUi();
  var cfg = lerConfig();
  var ano = cfg && Number(cfg.ano);
  if (!ano || ano < 2000 || ano > 2100) {
    ui.alert('Preencha o ano na aba Configuração (quatro dígitos, ex.: 2026) e tente de novo.');
    return;
  }
  migrarNomesDasAbas();
  var ss = SpreadsheetApp.getActive();
  var faltam = mesesDoAno(ano).filter(function (m) {
    return !ss.getSheetByName(nomeAba(m.mes));
  });
  if (!faltam.length) {
    ui.alert('As doze abas de ' + ano + ' já existem. Nada a fazer.');
    return;
  }
  // Dentro do comLocaleNeutro como o montarTudo: sem isso as formulas das abas
  // novas nascem com #ERROR!, porque em pt_BR o separador de argumentos e ';'.
  // A Geral entra junto porque a faixa dos graficos e justa (ver linhasComDado).
  comLocaleNeutro(function () {
    faltam.forEach(function (m) { montarMes(m.ano, m.mes, []); });
    montarConsolidado();
    montarGeral();
  });
  ordenarAbas();
  irParaMesAtual();
  ui.alert(faltam.length + ' aba(s) criada(s) para ' + ano + ': '
    + faltam.map(function (m) { return nomeAba(m.mes); }).join(', ')
    + '. ' + conferirFormulas());
}

function atualizarGeral() {
  migrarNomesDasAbas();
  comLocaleNeutro(function () {
    montarConsolidado();
    montarGeral();
  });
  ordenarAbas();
  SpreadsheetApp.getUi().alert('Resumo geral atualizado. ' + conferirFormulas());
}

// Geral primeiro, meses de janeiro a dezembro (como um calendario), apoio no fim.
function ordenarAbas() {
  var ss = SpreadsheetApp.getActive();
  var ordem = [ABA_GERAL]
    .concat(abasDeMes().map(function (m) { return m.aba.getName(); }))
    .concat([ABA_PRODUTOS, ABA_CONFIG, ABA_LEIAME]);
  ordem.forEach(function (nome, i) {
    var aba = ss.getSheetByName(nome);
    if (!aba) return;
    ss.setActiveSheet(aba);
    ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName(ABA_GERAL));
}
