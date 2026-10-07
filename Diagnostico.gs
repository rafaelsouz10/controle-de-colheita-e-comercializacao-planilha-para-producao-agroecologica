// Ferramenta de bancada (submenu Teste, so na planilha de teste). Nasceu para
// achar o bug do grafico vazio e ficou: a bissecao serve para qualquer opcao de
// grafico que venha a quebrar no futuro.
//
// O que ja foi MEDIDO (nao suposto):
//   - dados em P/Q corretos (conferidos contra as abas de mes)
//   - faixa do grafico de meses: P3:Q5, exatamente a pedida
//   - faixa da pizza: A24:B30, correta
//   - um grafico so com faixa + setNumHeaders(1) FUNCIONA
//
// Logo, o defeito esta em alguma das opcoes de estilo que os graficos de verdade
// passam e o de teste nao passava. diagnosticarOpcoes() monta oito graficos
// empilhados, cada um com uma opcao a mais que o anterior. O primeiro que sair
// vazio nomeia o culpado, sem precisar de oito ciclos de push.

function diagnosticarOpcoes() {
  var aba = SpreadsheetApp.getActive().getSheetByName(ABA_GERAL);
  limparTestes(aba);

  // Cada passo acumula os anteriores, na mesma ordem em que montarGeral aplica.
  var passos = [
    ['1 - so faixa', function (b) { return b; }],
    ['2 - + legenda', function (b) {
      return b.setOption('legend', { position: 'none' }); }],
    ['3 - + cores', function (b) {
      return b.setOption('colors', [COR.medio]); }],
    ['4 - + tamanho', function (b) {
      return b.setOption('width', 470).setOption('height', 250); }],
    ['5 - + fundo', function (b) {
      return b.setOption('backgroundColor', COR.fundo); }],
    ['6 - + eixo X', function (b) {
      return b.setOption('hAxis', { textStyle: { fontSize: 10, color: COR.suave } }); }],
    ['7 - + eixo Y', function (b) {
      return b.setOption('vAxis', { textStyle: { fontSize: 10, color: COR.suave },
                                    gridlines: { color: COR.borda } }); }],
    ['8 - + area', function (b) {
      return b.setOption('chartArea', { left: 45, top: 12,
                                        width: '85%', height: '78%' }); }]
  ];

  var linha = LINHA_TESTES;
  for (var i = 0; i < passos.length; i++) {
    var b = aba.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(aba.getRange(3, COL_AUX, 3, 2))
      .setNumHeaders(1);
    for (var j = 0; j <= i; j++) b = passos[j][1](b);
    aba.insertChart(b.setOption('title', passos[i][0]).setPosition(linha, 1, 0, 0).build());
    linha += 15;
  }

  SpreadsheetApp.getUi().alert('Oito gráficos criados a partir da linha '
    + LINHA_TESTES + '. Role para baixo e me diga o TÍTULO do primeiro que'
    + ' aparecer vazio.');
}

var LINHA_TESTES = 45;

// Tira os graficos de teste sem tocar nos dois de verdade, que ficam acima.
function limparTestes(aba) {
  aba = aba || SpreadsheetApp.getActive().getSheetByName(ABA_GERAL);
  aba.getCharts().forEach(function (c) {
    if (c.getContainerInfo().getAnchorRow() >= LINHA_TESTES) aba.removeChart(c);
  });
}
