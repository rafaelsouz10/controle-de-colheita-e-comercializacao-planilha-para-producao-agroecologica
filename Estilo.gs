// estilizar(): aplica (ou reaplica) toda a formatacao numa planilha que JA tem dados.
//
// Serve para quando a formatacao for bagunçada (coluna arrastada, cor trocada,
// grafico apagado) ou para uma planilha que chegou de outro lugar ja com registros.
// Diferenca para novaPlanilha()/importarCaderno(): aqueles apagam tudo, perdendo
// o que foi digitado depois da transcricao. Este devolve os mesmos registros no lugar.

function estilizar() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('America/Bahia');

  migrarNomesDasAbas();
  var meses = abasDeMes();
  if (!meses.length) {
    SpreadsheetApp.getUi().alert(
      'Não encontrei nenhuma aba de mês (nomes no padrão "Ago 2026"). '
      + 'Se a planilha está vazia, use "Reconstruir planilha inteira".');
    return;
  }

  // 1. Guarda o que e dado antes de qualquer limpeza.
  var guardados = meses.map(function (m) {
    return { ano: m.ano, mes: m.mes, registros: lerRegistros(m.aba) };
  });
  var cadastro = lerCadastro();
  var config = lerConfig() || configVazia();

  // 2. Reconstroi tudo devolvendo os mesmos dados.
  comLocaleNeutro(function () {
    montarConsolidado();
    montarConfig(config);
    montarProdutos(cadastro.produtos, cadastro.destinos);
    montarLeiaMe();
    guardados.forEach(function (g) { montarMes(g.ano, g.mes, g.registros); });
    montarConsolidado();
    montarGeral();
  });
  ordenarAbas();

  var total = guardados.reduce(function (s, g) { return s + g.registros.length; }, 0);
  SpreadsheetApp.getUi().alert(
    'Estilo aplicado: ' + guardados.length + ' aba(s) de mês, '
    + total + ' registros preservados. ' + conferirFormulas());
}

// Le a tabela de registro de uma aba de mes, descartando linhas vazias.
// Devolve no formato que montarMes espera: [ano, mes, dia, ...].
function lerRegistros(aba) {
  var ultima = aba.getLastRow();
  if (ultima < LINHA_DADOS) return [];
  return aba.getRange(LINHA_DADOS, 1, ultima - LINHA_DADOS + 1, N_COLS).getValues()
    .filter(function (l) {
      return l[0] instanceof Date || l[1] !== '' || l[2] !== '';
    })
    .map(function (l) {
      var d = l[0] instanceof Date ? l[0] : new Date(l[0]);
      var valida = !isNaN(d.getTime());
      return [
        valida ? d.getFullYear() : '', valida ? d.getMonth() + 1 : '',
        valida ? d.getDate() : '',
        l[1], l[2], l[3], l[4], l[5]
      ];
    });
}

// Preserva o cadastro de produtos e destinos, se ja existir.
function lerCadastro() {
  var aba = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS);
  if (!aba || aba.getLastRow() < 5) return { produtos: [], destinos: [] };
  var n = aba.getLastRow() - 4;
  var naoVazio = function (l) { return String(l[0]).trim() !== ''; };
  return {
    produtos: aba.getRange(5, 1, n, 4).getValues().filter(naoVazio),
    destinos: aba.getRange(5, 6, n, 2).getValues().filter(naoVazio)
  };
}
