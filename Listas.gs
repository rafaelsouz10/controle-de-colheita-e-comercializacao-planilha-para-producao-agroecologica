// Aba Produtos: cadastro que alimenta os menus suspensos das abas de mes.
// Colunas A-E sao editaveis pela agricultora; G-I sao auxiliares (ocultas)
// e guardam so os itens ativos, que e o que a validacao le.

// Sem argumentos, o cadastro nasce com as listas transcritas do caderno.
// Com argumentos (vindos de estilizar), preserva o que ja estava cadastrado.
function montarProdutos(produtosAtuais, destinosAtuais) {
  var ss = SpreadsheetApp.getActive();
  var aba = ss.getSheetByName(ABA_PRODUTOS) || ss.insertSheet(ABA_PRODUTOS);
  limparAba(aba);

  aba.getRange('A1:E1').merge().setValue('CADASTRO DE PRODUTOS E DESTINOS')
    .setFontSize(14).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COR.escuro)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  aba.setRowHeight(1, 34);
  aba.getRange('A2:E2').merge()
    .setValue('Para incluir um produto novo, escreva na primeira linha vazia. '
            + 'Para parar de usar um produto, escreva "não" em Ativo (não apague a linha: '
            + 'os registros antigos usam esse nome).')
    .setFontSize(9).setFontColor(COR.suave).setWrap(true);
  aba.setRowHeight(2, 32);

  cabecalhoTabela(aba, 4, 1, ['Produto', 'Unidade padrão', 'Ativo', 'Observação']);
  // Sem lista informada, comeca pela lista base do modelo (ver PRODUTOS_BASE).
  var prods = produtosAtuais && produtosAtuais.length ? produtosAtuais
    : PRODUTOS_BASE.map(function (p) { return [p[0], p[1], 'sim', '']; });
  aba.getRange(5, 1, prods.length, 4).setValues(prods);

  // Destinos, ao lado, com a mesma logica.
  cabecalhoTabela(aba, 4, 6, ['Destino', 'Ativo']);
  // Destinos nao tem lista base: sao os compradores de cada produtor.
  var dests = destinosAtuais || [];
  if (dests.length) aba.getRange(5, 6, dests.length, 2).setValues(dests);

  var simNao = SpreadsheetApp.newDataValidation()
    .requireValueInList(['sim', 'não'], true).setAllowInvalid(false).build();
  aba.getRange(5, 3, 500, 1).setDataValidation(simNao);
  aba.getRange(5, 7, 200, 1).setDataValidation(simNao);
  aba.getRange(5, 2, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['pacote', 'unidade'], true).setAllowInvalid(false).build());

  // Auxiliares: listas ja filtradas e ordenadas, fonte da validacao nas abas de mes.
  aba.getRange('I4').setValue('ativos (uso interno)').setFontColor(COR.suave).setFontSize(8);
  aba.getRange('I5').setFormula('=IFERROR(SORT(FILTER(A5:A,A5:A<>"",C5:C="sim")),"")');
  aba.getRange('J5').setFormula('=IFERROR(SORT(FILTER(F5:F,F5:F<>"",G5:G="sim")),"")');
  // Podem ficar ocultas: alimentam menu suspenso, nao grafico. Menu le faixa
  // escondida sem problema; grafico, nao (por isso as auxiliares dos paineis
  // ficam visiveis).
  aba.hideColumns(9, 2);

  // Produto repetido no cadastro: pinta para a agricultora notar na hora.
  var regra = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A5<>"",COUNTIF($A$5:$A,$A5)>1)')
    .setBackground(COR.alerta).setRanges([aba.getRange('A5:A500')]).build();
  aba.setConditionalFormatRules([regra]);

  aba.setColumnWidth(1, 170); aba.setColumnWidth(2, 110); aba.setColumnWidth(3, 60);
  aba.setColumnWidth(4, 220); aba.setColumnWidth(5, 20);
  aba.setColumnWidth(6, 180); aba.setColumnWidth(7, 60);
  aba.setFrozenRows(4);
  aba.getRange(5, 1, Math.max(prods.length, dests.length, 1), 7).setFontSize(10);
}

function montarLeiaMe() {
  var ss = SpreadsheetApp.getActive();
  var aba = ss.getSheetByName(ABA_LEIAME) || ss.insertSheet(ABA_LEIAME);
  limparAba(aba);

  var linhas = [
    ['COMO USAR ESTA PLANILHA', ''],
    ['', ''],
    ['1. Antes de começar',
     'Na aba "Configuração", preencha seu nome, núcleo, grupo e a rede ou associação. '
     + 'É só uma vez. Eles aparecem no topo da aba "Geral".'],
    ['2. Registrar a colheita',
     'Abra a aba do mês (ex.: "Ago"), desça até a tabela "REGISTRO DE COLHEITA" e escreva '
     + 'na primeira linha vazia: data, produto, quantidade, unidade e destino. '
     + 'O quadro de cima se atualiza sozinho.'],
    ['3. Data',
     'Escreva dia/mês/ano, ex.: 05/08/2026. Se a data não for do mês da aba, a célula fica '
     + 'vermelha.'],
    ['4. Quantidade',
     'Só o número, ex.: 22. A planilha não aceita palavras ("vinte") nem número negativo, '
     + 'porque eles sairiam das somas sem ninguém perceber.'],
    ['5. Unidade',
     'Escolha "pacote" ou "unidade" na lista. No computador ela aparece sozinha ao escolher o '
     + 'produto; no celular, escolha você. Pacote e unidade nunca se somam: por isso os totais '
     + 'aparecem separados.'],
    ['6. Produto ou destino novo',
     'Cadastre na aba "Produtos" e ele passa a aparecer na lista. Dá para digitar um nome '
     + 'fora da lista: a planilha avisa (fica laranja), mas aceita.'],
    ['7. Parar de usar um produto',
     'Na aba "Produtos", escreva "não" em Ativo. Não apague a linha nem mude o nome: os '
     + 'registros antigos usam esse nome. Para mudar de nome, cadastre o novo e desative o antigo.'],
    ['8. Ver um período',
     'Na aba "Geral", logo abaixo dos totais, escolha "PERÍODO: de ... até ...". Os totais, '
     + 'os mais colhidos e os destinos passam a mostrar só aqueles meses. O gráfico de '
     + 'colheita por mês continua mostrando o ano todo.'],
    ['9. Células coloridas',
     'Vermelho claro: data fora do mês ou produto sem quantidade. Laranja claro: produto '
     + 'fora do cadastro, ou um registro igual a outro do mesmo dia (pode ser repetição). '
     + 'São avisos: confira no caderno e corrija se for o caso.'],
    ['10. No celular e sem internet',
     'Abra a planilha uma vez com internet e marque "Disponível off-line" (nos três pontinhos '
     + 'do app). Depois ela abre e aceita registros sem sinal, e sincroniza quando a internet '
     + 'voltar. O menu "Colheita" só existe no computador; para registrar, não é preciso.'],
    ['11. Os meses do ano',
     'As doze abas, de janeiro a dezembro, já vêm prontas para o ano da aba "Configuração". '
     + 'Se faltar alguma, no computador: menu "Colheita" > "Manutenção" > "Recriar meses que faltam".'],
    ['', ''],
    ['Origem dos dados',
     'Ficha de controle de colheita e comercialização do caderno de campo. Quem é o(a) '
     + 'agricultor(a) e a rede estão na aba "Configuração".']
  ];
  aba.getRange(1, 1, linhas.length, 2).setValues(linhas);
  aba.getRange('A1:B1').merge().setFontSize(14).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground(COR.escuro)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  aba.setRowHeight(1, 34);
  aba.getRange(3, 1, linhas.length - 2, 1).setFontWeight('bold').setFontColor(COR.escuro);
  aba.getRange(1, 1, linhas.length, 2).setVerticalAlignment('top').setWrap(true);
  aba.setColumnWidth(1, 240);
  aba.setColumnWidth(2, 560);
}

// Aba Configuracao: quem e o produtor e de que ano e a planilha. E a unica coisa
// que muda de uma copia para outra; todo o resto e igual para todos os produtores.
// Quem le estes campos sao formulas (ver refConfig), nao o script.
function montarConfig(cfg) {
  var ss = SpreadsheetApp.getActive();
  var aba = ss.getSheetByName(ABA_CONFIG) || ss.insertSheet(ABA_CONFIG);
  limparAba(aba);

  aba.getRange('A1:B1').merge().setValue('CONFIGURAÇÃO')
    .setFontSize(14).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COR.escuro)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  aba.setRowHeight(1, 34);
  aba.getRange('A2:B2').merge()
    .setValue('Preencha uma vez. Nome, núcleo, grupo e rede aparecem no topo da aba Geral.')
    .setFontSize(9).setFontColor(COR.suave).setWrap(true);

  var linhas = CAMPOS_CONFIG.map(function (c) {
    var v = cfg && cfg[c[0]] !== undefined ? cfg[c[0]] : '';
    return [c[1], v];
  });
  var faixa = aba.getRange(LINHA_CONFIG, 1, linhas.length, 2).setValues(linhas);
  aba.getRange(LINHA_CONFIG, 1, linhas.length, 1)
    .setFontWeight('bold').setFontColor(COR.escuro);
  // Campo de digitar destacado, para nao confundir com rotulo.
  aba.getRange(LINHA_CONFIG, 2, linhas.length, 1)
    .setBackground('#FFFDE7').setFontSize(11);
  faixa.setBorder(true, true, true, true, true, true, COR.borda,
                  SpreadsheetApp.BorderStyle.SOLID);

  var linhaAno = LINHA_CONFIG + CAMPOS_CONFIG.length - 1;
  aba.getRange(linhaAno, 2).setNumberFormat('0').setHorizontalAlignment('left')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(2000, 2100).setAllowInvalid(false)
      .setHelpText('Ano com quatro dígitos, ex.: 2026').build());

  aba.setColumnWidth(1, 160);
  aba.setColumnWidth(2, 340);
  return aba;
}

// Le os valores atuais, para que estilizar() nao apague o que o produtor preencheu.
function lerConfig() {
  var aba = SpreadsheetApp.getActive().getSheetByName(ABA_CONFIG);
  if (!aba) return null;
  // Le pelo ROTULO da coluna A, nao pela posicao: quando um campo novo entra no
  // meio (como a Rede, antes do Ano), uma planilha antiga tem os campos em outras
  // linhas, e ler por posicao trocaria um pelo outro.
  var v = aba.getRange(LINHA_CONFIG, 1, 20, 2).getValues();
  var cfg = {};
  CAMPOS_CONFIG.forEach(function (c) {
    for (var i = 0; i < v.length; i++) {
      if (v[i][0] === c[1]) { cfg[c[0]] = v[i][1]; return; }
    }
    cfg[c[0]] = '';
  });
  return cfg;
}
