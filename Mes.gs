// Aba de mes: dashboard nas linhas 1-14 e a tabela de registro a partir da linha 16.
// A tabela cresce para baixo, entao o dashboard fica sempre no mesmo lugar.
//
// Colunas: A Data | B Produto | C Quantidade | D Unidade | E Destino | F Observação.
// As quatro primeiras do caderno, na mesma ordem, mais a unidade separada do numero
// (para poder somar) e um campo livre para o que o caderno anota na margem.

var CAB_REGISTRO = ['Data', 'Produto', 'Quantidade', 'Unidade', 'Destino', 'Observação'];
// Larguras escolhidas para servir a DUAS coisas que dividem as colunas: a tabela de
// registro embaixo e os quatro cards em cima. Cards: A:B (265), C:D (210), E (210),
// F (210). Com larguras desiguais o card de pacotes caia em 95 px e cortava o texto.
var LARGURAS = [95, 170, 110, 100, 210, 210];
var N_COLS = CAB_REGISTRO.length;

function montarMes(ano, mes, registros) {
  var ss = SpreadsheetApp.getActive();
  var nome = nomeAba(mes);
  var aba = ss.getSheetByName(nome) || ss.insertSheet(nome);
  limparAba(aba);

  var ref = new Date(ano, mes - 1, 1);
  var ultima = LINHA_DADOS + Math.max((registros || []).length, 0) + 250;
  if (aba.getMaxRows() < ultima) aba.insertRowsAfter(aba.getMaxRows(), ultima - aba.getMaxRows());

  aba.getRange(1, 1, aba.getMaxRows(), N_COLS + 1).setBackground(COR.fundo);
  // O painel inclui a area do grafico, a direita da tabela (ate a coluna L).
  // Pintar SO aqui, no comeco: pintado depois, cobria o titulo, os cards e os
  // cabecalhos ja coloridos, e o titulo ficou branco sobre branco.
  aba.getRange(1, 1, LINHA_CAB - 1, 12).setBackground(COR.fundo);
  aba.getRange(LINHA_CAB, 1, aba.getMaxRows() - LINHA_CAB + 1, N_COLS).setBackground('#FFFFFF');

  // --- cabecalho -----------------------------------------------------------
  aba.getRange(1, 1, 1, N_COLS).merge()
    // Ano por formula: vem da Configuracao, como em todo o resto. Fica visivel no
    // titulo porque e o que aparece na impressao para o certificador.
    .setFormula('="COLHEITA - ' + MESES_LONGOS[mes - 1].toUpperCase() + ' DE "&'
      + refConfig('ano'))
    .setFontSize(15).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COR.escuro)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  aba.setRowHeight(1, 36);
  aba.getRange(2, 1, 1, N_COLS).merge()
    // A rede vem da Configuracao: o modelo nao leva o nome de nenhuma organizacao.
    .setFormula('=IF(' + refConfig('rede') + '="","Caderno de campo","Caderno de campo - "&'
      + refConfig('rede') + ')&"   |   o quadro abaixo se atualiza sozinho conforme a tabela é preenchida"')
    .setFontSize(9).setFontColor(COR.suave).setHorizontalAlignment('center');

  // --- cards ---------------------------------------------------------------
  var d = LINHA_DADOS;
  card(aba, 3, 1, 2, 'PACOTES COLHIDOS',
       '=SUMIF(D' + d + ':D,"pacote",C' + d + ':C)');
  card(aba, 3, 3, 2, 'UNIDADES COLHIDAS',
       '=SUMIF(D' + d + ':D,"unidade",C' + d + ':C)');
  // So conta se houver registro. Sem registro, o FILTER devolve erro, e o COUNTA
  // CONTA o erro como um item: o mes vazio mostrava "1 dia" e "1 produto". O
  // IFERROR em volta nao pegava, porque o COUNTA nao falha, devolve 1.
  card(aba, 3, 5, 1, 'DIAS DE COLHEITA',
       '=IF(COUNTA(A' + d + ':A)=0,0,COUNTA(UNIQUE(FILTER(A' + d + ':A,A' + d + ':A<>""))))');
  card(aba, 3, 6, 1, 'PRODUTOS DIFERENTES',
       '=IF(COUNTA(B' + d + ':B)=0,0,COUNTA(UNIQUE(FILTER(B' + d + ':B,B' + d + ':B<>""))))');

  // --- tabelinhas ----------------------------------------------------------
  // So pacotes: somar pacote com unidade daria um numero que nao significa nada.
  // Os nomes caem nas colunas largas: produto em B (170 px), destino em E (210 px).
  // Antes caiam em A e C (95 e 110 px) e cortavam: "Cliente direto / A".
  tituloSecao(aba, 6, 2, 2, 'Mais colhidos no mês');
  cabecalhoTabela(aba, 7, 2, ['Produto', 'Pacotes']);
  aba.getRange(8, 2).setFormula(
    '=IFERROR(QUERY(B' + d + ':E,"select B, sum(C) where B is not null '
    + 'and D=\'pacote\' group by B order by sum(C) desc limit 6 label sum(C) \'\'"),"")');

  tituloSecao(aba, 6, 5, 2, 'Por destino');
  cabecalhoTabela(aba, 7, 5, ['Destino', 'Pacotes']);
  aba.getRange(8, 5).setFormula(
    '=IFERROR(QUERY(B' + d + ':E,"select E, sum(C) where E is not null '
    + 'and D=\'pacote\' group by E order by sum(C) desc limit 6 label sum(C) \'\'"),"")');

  // Com as tabelinhas nas colunas largas, nao sobra espaco para o grafico dentro
  // da moldura A:F, e descer a tabela de registro nao e opcao (o script todo conta
  // com ela na linha 17). Ele fica a direita, a partir da coluna H.
  tituloSecao(aba, 6, 8, 4, 'Por semana (pacotes)');
  aba.setColumnWidth(7, 20);

  // --- area auxiliar que alimenta o grafico --------------------------------
  aba.getRange(1, COL_AUX).setValue(ref).setNumberFormat('dd/mm/yyyy'); // 1o dia do mes
  // R guarda a data da segunda-feira, que e o que soma; P guarda so o rotulo em
  // texto. Com data no eixo, o grafico usa escala continua e poucas semanas viram
  // barras finissimas. Com texto, o eixo e discreto.
  // As semanas do mes sao escritas todas de uma vez, inclusive as que ainda nao
  // tem colheita (que somam zero). Duas razoes: a faixa do grafico nunca tem linha
  // vazia (ver linhasComDado em Config.gs), e o mes aparece inteiro desde o dia 1,
  // entao uma semana sem colheita fica visivel como tal, em vez de sumir do eixo.
  var semanas = semanasDoMes(ano, mes);
  aba.getRange(3, COL_AUX, 1, 3).setValues([['Semana', 'Pacotes', 'data (cálculo)']]);
  aba.getRange(4, COL_AUX + 2, semanas.length, 1).setValues(semanas);
  // Rotulo fixo e um SUMIFS por linha: so funcoes classicas, que recalculam
  // offline e em qualquer aparelho. (Antes era MAP/LAMBDA, de 2022.)
  // Texto puro ('@') ANTES de escrever: o rotulo e escrito com a planilha em ingles
  // (comLocaleNeutro), e o Sheets lia "03/08" como data americana, 3 de marco,
  // exibida depois como "08/03". So semanas com dia ate 12 saiam trocadas.
  aba.getRange(4, COL_AUX, semanas.length, 1).setNumberFormat('@').setValues(semanas.map(function (s) {
    var dt = s[0];
    return [('0' + dt.getDate()).slice(-2) + '/' + ('0' + (dt.getMonth() + 1)).slice(-2)];
  }));
  aba.getRange(4, COL_AUX + 1, semanas.length, 1).setFormulas(semanas.map(function (s, i) {
    var r = 'R' + (4 + i);
    return ['=SUMIFS($C$' + d + ':$C,$D$' + d + ':$D,"pacote",$A$' + d + ':$A,">="&' + r
      + ',$A$' + d + ':$A,"<"&' + r + '+7)'];
  }));
  aba.getRange(4, COL_AUX + 2, semanas.length, 1).setNumberFormat('dd/mm/yyyy');
  // Nao ocultar: o grafico do Sheets ignora dados de linha ou coluna escondida,
  // e sairia em branco. Ficam longe o bastante para nao atrapalhar a leitura.
  aba.getRange(2, COL_AUX, 1, 3).merge()
    .setValue('área auxiliar - alimenta o gráfico por semana')
    .setFontSize(8).setFontColor(COR.suave);

  // Espera o recalculo antes de criar o grafico: sem isso ele pode nascer
  // apontando para um intervalo ainda vazio e ficar sem serie nenhuma.
  SpreadsheetApp.flush();
  aba.insertChart(aba.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(aba.getRange(3, COL_AUX, semanas.length + 1, 2))
    .setNumHeaders(1)
    .setPosition(7, 8, 4, 2)
    .setOption('title', '')
    .setOption('legend', { position: 'none' })
    .setOption('colors', [COR.medio])
    // A partir da coluna H; 390 px vai ate K, longe da area auxiliar (P em diante).
    .setOption('width', 390).setOption('height', 155)
    .setOption('backgroundColor', COR.fundo)
    .setOption('hAxis', { textStyle: { fontSize: 10, color: COR.suave } })
    .setOption('vAxis', { textStyle: { fontSize: 10, color: COR.suave },
                          gridlines: { color: COR.borda } })
    // Sem 'chartArea': medido que essa opcao esvazia o grafico (mistura px
    // em left/top com % em width/height). Ver o historico em Diagnostico.gs.
    .build());

  // --- tabela de registro --------------------------------------------------
  aba.getRange(LINHA_CAB - 1, 1, 1, N_COLS).merge()
    .setValue('REGISTRO DE COLHEITA   -   escreva na primeira linha vazia')
    .setFontSize(11).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COR.medio)
    .setVerticalAlignment('middle');
  aba.setRowHeight(LINHA_CAB - 1, 26);
  aba.getRange(LINHA_CAB, 1, 1, N_COLS).setValues([CAB_REGISTRO])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COR.escuro)
    .setVerticalAlignment('middle');
  aba.setRowHeight(LINHA_CAB, 24);

  if (registros && registros.length) {
    aba.getRange(d, 1, registros.length, N_COLS).setValues(registros.map(function (r) {
      // Data em branco e possivel quando os registros vem de uma aba ja preenchida
      // a mao: preservar a linha vale mais do que gravar uma data invalida.
      var data = (r[0] === '' || r[0] === null) ? '' : new Date(r[0], r[1] - 1, r[2]);
      return [data, r[3], r[4], r[5], r[6], r[7]];
    }));
  }

  var n = aba.getMaxRows() - d + 1;
  aba.getRange(d, 1, n, 1).setNumberFormat('dd/mm/yyyy');
  aba.getRange(d, 3, n, 1).setNumberFormat('0').setHorizontalAlignment('center');
  aba.getRange(d, 4, n, 1).setHorizontalAlignment('center');
  aba.getRange(d, 1, n, N_COLS).setFontSize(10).setFontColor(COR.texto);

  aplicarValidacoes(aba);
  aplicarAvisos(aba);

  LARGURAS.forEach(function (w, i) { aba.setColumnWidth(i + 1, w); });
  aba.getRange(LINHA_CAB, 1, aba.getMaxRows() - LINHA_CAB + 1, N_COLS)
    .setBorder(null, null, null, null, true, true, COR.borda, SpreadsheetApp.BorderStyle.SOLID);
  aba.getRange(LINHA_CAB, 1, aba.getMaxRows() - LINHA_CAB + 1, N_COLS).createFilter();
  return aba;
}

// As segundas-feiras que tocam o mes, da que contem o dia 1 ate a que contem o
// ultimo dia. Sao 5 ou 6 por mes. Devolve no formato de setValues (uma por linha).
function semanasDoMes(ano, mes) {
  var primeiro = new Date(ano, mes - 1, 1);
  var ultimo = new Date(ano, mes, 0);
  var seg = new Date(ano, mes - 1, 1 - ((primeiro.getDay() + 6) % 7));
  var fora = [];
  while (seg <= ultimo) {
    fora.push([new Date(seg.getTime())]);
    seg.setDate(seg.getDate() + 7);
  }
  return fora;
}

// Menus suspensos. Produto e destino avisam mas nao bloqueiam: no calor da colheita,
// travar o registro e pior do que um nome fora do padrao.
function aplicarValidacoes(aba) {
  var p = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS);
  var n = aba.getMaxRows() - LINHA_DADOS + 1;
  var deLista = function (faixa, permitirFora) {
    return SpreadsheetApp.newDataValidation()
      .requireValueInRange(faixa, true).setAllowInvalid(permitirFora).build();
  };
  aba.getRange(LINHA_DADOS, 2, n, 1).setDataValidation(deLista(p.getRange('I5:I'), true));
  aba.getRange(LINHA_DADOS, 5, n, 1).setDataValidation(deLista(p.getRange('J5:J'), true));
  aba.getRange(LINHA_DADOS, 4, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['pacote', 'unidade'], true).setAllowInvalid(false).build());

  // Quantidade RECUSA, ao contrario do resto, que so avisa. Nome fora do padrao ainda
  // e somado; ja "vinte" ou "-3" saem das somas sem ninguem ver, e o total do mes
  // fica errado em silencio. Recusar na hora, dizendo o que digitar, e mais barato.
  aba.getRange(LINHA_DADOS, 3, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0).setAllowInvalid(false)
    .setHelpText('Digite só o número, maior que zero. Ex.: 22 (a unidade vai na coluna ao lado).')
    .build());
}

// Avisos visuais. Nenhum deles impede o registro: so pedem uma conferida no caderno.
function aplicarAvisos(aba) {
  var d = LINHA_DADOS;
  var faixa = [aba.getRange(d, 1, aba.getMaxRows() - d + 1, N_COLS)];
  var regra = function (formula, cor) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula).setBackground(cor).setRanges(faixa).build();
  };
  aba.setConditionalFormatRules([
    // data fora do mes da aba
    regra('=AND($A' + d + '<>"",OR($A' + d + '<$P$1,$A' + d + '>=EOMONTH($P$1,0)+1))', COR.alerta),
    // produto sem quantidade
    regra('=AND($B' + d + '<>"",$C' + d + '="")', COR.alerta),
    // registro identico a outro do mesmo dia (pode ser digitacao repetida ou 2o lote)
    regra('=AND($B' + d + '<>"",COUNTIFS($A$' + d + ':$A,$A' + d + ',$B$' + d + ':$B,$B' + d
        + ',$C$' + d + ':$C,$C' + d + ',$E$' + d + ':$E,$E' + d + ')>1)', '#FFF4E0'),
    // produto que nao esta no cadastro
    regra('=AND($B' + d + '<>"",COUNTIF(INDIRECT("' + ABA_PRODUTOS + '!$A$5:$A"),$B' + d + ')=0)',
      '#FFF4E0')
  ]);
}

// Preenche a unidade padrao ao escolher o produto, e a data repetindo a linha de cima.
// Gatilho simples: roda sozinho, sem precisar autorizar nada.
function onEdit(e) {
  try {
    var aba = e.range.getSheet();
    if (!ehAbaDeMes(aba.getName())) return;
    var lin = e.range.getRow(), col = e.range.getColumn();
    if (col !== 2 || lin < LINHA_DADOS || e.range.getNumRows() > 1) return;
    var produto = e.range.getValue();
    if (!produto) return;

    var cadastro = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS)
      .getRange('A5:B500').getValues();
    for (var i = 0; i < cadastro.length; i++) {
      if (cadastro[i][0] === produto && cadastro[i][1]) {
        var unidade = aba.getRange(lin, 4);
        if (!unidade.getValue()) unidade.setValue(cadastro[i][1]);
        break;
      }
    }
    var data = aba.getRange(lin, 1);
    if (!data.getValue() && lin > LINHA_DADOS) {
      var acima = aba.getRange(lin - 1, 1).getValue();
      if (acima) data.setValue(acima);
    }
  } catch (err) {
    // Um erro aqui nao pode atrapalhar quem esta digitando.
  }
}
