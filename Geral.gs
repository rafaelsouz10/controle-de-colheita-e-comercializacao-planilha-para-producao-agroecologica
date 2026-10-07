// Aba oculta que empilha as abas de mes, e o dashboard consolidado que le dela.
// O empilhamento e reescrito toda vez que um mes novo e criado, para a Geral
// nunca depender de alguem lembrar de editar formula.
//
// Colunas herdadas da aba de mes: A Data | B Produto | C Quantidade | D Unidade |
// E Destino | F Observação. A coluna I guarda o 1o dia do mes de cada linha, e a J
// o numero do mes (1 a 12), que e o que o seletor de periodo da Geral compara.

function montarConsolidado() {
  var ss = SpreadsheetApp.getActive();
  var aba = ss.getSheetByName(ABA_CONSOLIDADO) || ss.insertSheet(ABA_CONSOLIDADO);
  limparAba(aba);
  aba.getRange(1, 1, 1, N_COLS).setValues([CAB_REGISTRO]).setFontWeight('bold');
  aba.getRange('I1:J1').setValues([['mes', 'n_mes']]);

  var partes = abasDeMes().map(function (m) {
    return "'" + m.aba.getName() + "'!A" + LINHA_DADOS + ':F';
  });
  if (partes.length) {
    aba.getRange('A2').setFormula('={' + partes.join(';') + '}');
    aba.getRange('I2').setFormula('=ARRAYFORMULA(IF(A2:A="","",EOMONTH(A2:A,-1)+1))');
    // Vazio de verdade (,,) e nao "": a QUERY decide o tipo da coluna pela maioria,
    // e milhares de "" nas linhas sem registro a fariam tratar J como texto.
    aba.getRange('J2').setFormula('=ARRAYFORMULA(IF(A2:A="",,MONTH(A2:A)))');
  }
  aba.getRange('A2:A').setNumberFormat('dd/mm/yyyy');
  aba.getRange('I2:I').setNumberFormat('mmm/yyyy');
  aba.hideSheet();
  return aba;
}

function montarGeral() {
  var ss = SpreadsheetApp.getActive();
  var aba = ss.getSheetByName(ABA_GERAL) || ss.insertSheet(ABA_GERAL, 0);
  // O periodo escolhido sobrevive a reconstrucao (Atualizar resumo, Aplicar estilo).
  var de = 'Jan', ate = 'Dez';
  var sel = aba.getRange('B6:D6').getValues()[0];
  if (MESES.indexOf(sel[0]) >= 0) de = sel[0];
  if (MESES.indexOf(sel[2]) >= 0) ate = sel[2];
  limparAba(aba);
  ss.setActiveSheet(aba);
  ss.moveActiveSheet(1);

  var C = "'" + ABA_CONSOLIDADO + "'!";
  aba.getRange(1, 1, aba.getMaxRows(), 8).setBackground(COR.fundo);

  // Periodo: o numero do mes "de" fica em T2 e o "ate" em T4 (area auxiliar).
  // Tudo que obedece ao periodo compara a coluna J da _Consolidado com essas duas
  // celulas, por um destes dois trechos. Filtrar por outra coisa no futuro (destino,
  // por exemplo) = mexer aqui, e mais nada.
  var NO_PERIODO = ',' + C + 'J2:J,">="&$T$2,' + C + 'J2:J,"<="&$T$4';     // SUMIFS
  var NO_PERIODO_Q = 'and J >= "&$T$2&" and J <= "&$T$4&" ';              // QUERY

  aba.getRange(1, 1, 1, 8).merge()
    .setFormula('="COLHEITA E COMERCIALIZAÇÃO - VISÃO GERAL "&' + refConfig('ano'))
    .setFontSize(16).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COR.escuro)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  aba.setRowHeight(1, 40);
  // Formula, nao texto: corrigir o nome na aba Configuracao atualiza aqui na hora,
  // inclusive offline, sem rodar script.
  aba.getRange(2, 1, 1, 8).merge()
    // Sem nome preenchido, a linha virava "|  Núcleo  |  Grupo  |" e parecia quebrada.
    .setFormula('=IF(' + refConfig('nome') + '="",'
      + '"Preencha a aba Configuração (nome, núcleo, grupo e rede)",'
      + refConfig('nome') + '&"   |   Núcleo "&' + refConfig('nucleo')
      + '&"   |   Grupo "&' + refConfig('grupo')
      + '&IF(' + refConfig('rede') + '="","","   |   "&' + refConfig('rede') + '))')
    .setFontSize(9).setFontColor(COR.suave).setHorizontalAlignment('center');

  card(aba, 3, 1, 2, 'PACOTES COLHIDOS',
       '=SUMIFS(' + C + 'C2:C,' + C + 'D2:D,"pacote"' + NO_PERIODO + ')');
  card(aba, 3, 3, 2, 'UNIDADES COLHIDAS',
       '=SUMIFS(' + C + 'C2:C,' + C + 'D2:D,"unidade"' + NO_PERIODO + ')');
  // So conta se houver registro no periodo; ver o comentario igual em Mes.gs
  // (sem registro, o COUNTA conta o erro do FILTER e mostra 1).
  card(aba, 3, 5, 2, 'DIAS DE COLHEITA',
       '=IF(COUNTIFS(' + C + 'A2:A,"<>"' + NO_PERIODO + ')=0,0,'
       + 'COUNTA(UNIQUE(FILTER(' + C + 'A2:A,' + C + 'A2:A<>"",'
       + C + 'J2:J>=$T$2,' + C + 'J2:J<=$T$4))))');
  card(aba, 3, 7, 2, 'PRODUTOS DIFERENTES',
       '=IF(COUNTIFS(' + C + 'B2:B,"<>"' + NO_PERIODO + ')=0,0,'
       + 'COUNTA(UNIQUE(FILTER(' + C + 'B2:B,' + C + 'B2:B<>"",'
       + C + 'J2:J>=$T$2,' + C + 'J2:J<=$T$4))))');

  // Linha de comparacao: o ultimo mes COM colheita contra o mes anterior a ele no
  // calendario, em valor absoluto. Com as doze abas, "o ultimo mes" era dezembro,
  // vazio, e a frase dizia "dezembro com 0, contra 0 em novembro". A posicao do
  // ultimo mes com colheita fica em T20 (area auxiliar).
  // Percentual foi descartado: um mes parcial (so o ultimo dia, por exemplo) contra
  // um mes cheio produzia algo como "+6000%", um numero que parece informacao e nao e.
  aba.getRange(5, 1, 1, 8).merge().setFormula(
    '=IF($T$20=0,"",'
    + '"Último mês com colheita: "&TEXT(INDEX(R4:R40,$T$20),"mmmm")&" com "'
    + '&INDEX(Q4:Q40,$T$20)&" pacotes"'
    + '&IF($T$20=1,"",", contra "&INDEX(Q4:Q40,$T$20-1)&" em "'
    + '&TEXT(INDEX(R4:R40,$T$20-1),"mmmm")))')
    .setFontSize(10).setFontColor(COR.suave).setHorizontalAlignment('center');

  // --- seletor de periodo (linha 6) -----------------------------------------
  aba.getRange('A6').setValue('PERÍODO:').setFontWeight('bold').setFontColor(COR.escuro)
    .setHorizontalAlignment('right');
  aba.getRange('C6').setValue('até').setHorizontalAlignment('center').setFontColor(COR.suave);
  var regraMes = SpreadsheetApp.newDataValidation()
    .requireValueInList(MESES, true).setAllowInvalid(false).build();
  aba.getRange('B6').setValue(de);
  aba.getRange('D6').setValue(ate);
  [aba.getRange('B6'), aba.getRange('D6')].forEach(function (r) {
    r.setDataValidation(regraMes).setBackground('#FFFDE7').setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, COR.borda,
                 SpreadsheetApp.BorderStyle.SOLID);
  });
  aba.getRange(6, 5, 1, 4).merge()
    .setFormula('=IF($T$2>$T$4,"Atenção: o mês inicial vem depois do final","")')
    .setFontColor('#C62828').setFontSize(9);
  aba.setRowHeight(6, 26);

  tituloSecao(aba, 7, 1, 5, 'Colheita por mês (pacotes, ano todo)');
  tituloSecao(aba, 7, 6, 3, 'Mais colhidos no período (pacotes)');
  cabecalhoTabela(aba, 8, 6, ['Produto', 'Pacotes', '% do total']);
  aba.getRange(9, 6).setFormula(
    '=IFERROR(QUERY(' + C + 'B2:J,"select B, sum(C) where B is not null and D=\'pacote\' '
    + NO_PERIODO_Q + 'group by B order by sum(C) desc limit 12 label sum(C) \'\'"),"")');
  aba.getRange(9, 8).setFormula(
    '=ARRAYFORMULA(IF(G9:G20="","",G9:G20/SUM($G$9:$G$20)))');
  aba.getRange(9, 8, 12, 1).setNumberFormat('0%');

  tituloSecao(aba, 23, 1, 5, 'Por destino no período (pacotes)');
  cabecalhoTabela(aba, 24, 1, ['Destino', 'Pacotes']);
  aba.getRange(25, 1).setFormula(
    '=IFERROR(QUERY(' + C + 'B2:J,"select E, sum(C) where E is not null and D=\'pacote\' '
    + NO_PERIODO_Q + 'group by E order by sum(C) desc label sum(C) \'\'"),"")');

  tituloSecao(aba, 23, 6, 3, 'Em unidades, no período');
  cabecalhoTabela(aba, 24, 6, ['Produto', 'Unidades']);
  aba.getRange(25, 6).setFormula(
    '=IFERROR(QUERY(' + C + 'B2:J,"select B, sum(C) where B is not null and D=\'unidade\' '
    + NO_PERIODO_Q + 'group by B order by sum(C) desc label sum(C) \'\'"),"")');

  tituloSecao(aba, 35, 1, 5, 'Distribuição por destino no período');

  // Auxiliar do periodo (coluna T): lista dos meses e o numero de cada escolha.
  var T = COL_AUX + 4;
  aba.getRange(1, T).setValue('de (nº)');
  aba.getRange(2, T).setFormula('=IFERROR(MATCH($B$6,$T$6:$T$17,0),1)');
  aba.getRange(3, T).setValue('até (nº)');
  aba.getRange(4, T).setFormula('=IFERROR(MATCH($D$6,$T$6:$T$17,0),12)');
  aba.getRange(6, T, 12, 1).setValues(MESES.map(function (m) { return [m]; }));
  aba.getRange(19, T).setValue('últ. mês c/ colheita');
  aba.getRange(20, T).setFormula('=IFERROR(ARRAYFORMULA(MAX(IF(Q4:Q40>0,ROW(Q4:Q40)-3,0))),0)');

  // --- area auxiliar que alimenta o grafico --------------------------------
  // R guarda a data de verdade, que e o que soma; P guarda so o rotulo em texto.
  // Com data no eixo, o grafico usa escala continua e dois meses viram duas barras
  // finissimas perdidas num ano inteiro. Com texto, o eixo e discreto e a barra
  // ocupa o espaco que lhe cabe.
  // Os meses vem das ABAS que existem, nao dos dados. A versao anterior derivava a
  // lista de _Consolidado, e por isso um mes recem-criado, ainda sem nenhuma
  // colheita, nao gerava linha nenhuma e simplesmente nao aparecia no grafico.
  // Escritos de uma vez, inclusive os zerados: a faixa nunca tem linha vazia (ver
  // linhasComDado em Config.gs) e um mes sem colheita fica visivel como tal.
  var meses = abasDeMes();
  aba.getRange(3, COL_AUX, 1, 3).setValues([['Mês', 'Pacotes', 'data (cálculo)']]);
  if (meses.length) {
    aba.getRange(4, COL_AUX + 2, meses.length, 1).setValues(meses.map(function (m) {
      return [new Date(m.ano, m.mes - 1, 1)];
    }));
    // Rotulo fixo e um SUMIFS por linha: so funcoes classicas, que recalculam
    // offline e em qualquer aparelho. (Antes era MAP/LAMBDA, de 2022.)
    aba.getRange(4, COL_AUX, meses.length, 1).setNumberFormat('@').setValues(meses.map(function (m) {
      return [MESES[m.mes - 1]];
    }));
    aba.getRange(4, COL_AUX + 1, meses.length, 1).setFormulas(meses.map(function (m, i) {
      return ['=SUMIFS(' + C + '$C$2:$C,' + C + '$D$2:$D,"pacote",' + C + '$I$2:$I,R' + (4 + i) + ')'];
    }));
    aba.getRange(4, COL_AUX + 2, meses.length, 1).setNumberFormat('mmm/yyyy');
  }
  // Nao ocultar: o grafico do Sheets ignora dados de linha ou coluna escondida,
  // e sairia em branco. Ficam longe o bastante para nao atrapalhar a leitura.
  aba.getRange(2, COL_AUX, 1, 3).merge()
    .setValue('área auxiliar - alimenta o gráfico ao lado')
    .setFontSize(8).setFontColor(COR.suave);

  // Espera o recalculo antes de criar o grafico: sem isso ele pode nascer
  // apontando para um intervalo ainda vazio e ficar sem serie nenhuma.
  SpreadsheetApp.flush();

  // Faixa justa: cabecalho + as linhas que existem. Ver linhasComDado em Config.gs.
  // Os meses ja sao conhecidos (vem das abas); os destinos saem de QUERY, entao
  // precisam ser contados depois do flush.
  // A pizza tem faixa FIXA de 10 destinos (a tabela A25:B34 inteira). A faixa e
  // gravada na montagem, entao medi-la pelos destinos existentes dava errado numa
  // planilha nova: zero destinos na montagem = pizza com espaco para um so, mesmo
  // depois de o produtor cadastrar os compradores dele. Linhas vazias no fim da
  // faixa nao atrapalham a pizza: testado num periodo de um mes so (1 destino, 9
  // linhas vazias). O que esvaziava grafico era o chartArea, nao linha vazia.
  var nDestinos = 10;

  aba.insertChart(aba.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(aba.getRange(3, COL_AUX, meses.length + 1, 2))
    .setNumHeaders(1)
    .setPosition(8, 1, 4, 2)
    .setOption('legend', { position: 'none' })
    .setOption('colors', [COR.medio])
    // 470 cabe nas colunas A-E (475px). Mais largo, invade a F e corta os nomes.
    .setOption('width', 470).setOption('height', 250)
    .setOption('backgroundColor', COR.fundo)
    .setOption('hAxis', { textStyle: { fontSize: 10, color: COR.suave } })
    .setOption('vAxis', { textStyle: { fontSize: 10, color: COR.suave },
                          gridlines: { color: COR.borda } })
    // Sem 'chartArea': medido que essa opcao esvazia o grafico (mistura px
    // em left/top com % em width/height). Ver o historico em Diagnostico.gs.
    .build());

  aba.insertChart(aba.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(aba.getRange(24, 1, nDestinos + 1, 2))
    .setNumHeaders(1)
    .setPosition(36, 1, 4, 2)
    .setOption('legend', { position: 'right', textStyle: { fontSize: 10, color: COR.suave } })
    .setOption('colors', ['#2E5E3A', '#7A9E3F', '#B8C96B', '#D9B24C', '#A8734A', '#6B8FA8'])
    .setOption('width', 380).setOption('height', 230)
    .setOption('pieSliceText', 'percentage')
    .setOption('backgroundColor', COR.fundo)
    // Sem 'chartArea': medido que essa opcao esvazia o grafico (mistura px
    // em left/top com % em width/height). Ver o historico em Diagnostico.gs.
    .build());

  [150, 95, 95, 95, 40, 175, 95, 95].forEach(function (w, i) {
    aba.setColumnWidth(i + 1, w);
  });
  aba.getRange(9, 6, 12, 3).setFontSize(10);
  aba.getRange(25, 1, 10, 2).setFontSize(10);
  aba.getRange(25, 6, 10, 2).setFontSize(10);
  return aba;
}
