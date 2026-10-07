# Controle de colheita e comercialização - planilha para produção agroecológica

Planilha do Google Sheets para o produtor registrar a colheita e o destino de cada
produto, no lugar (ou ao lado) da ficha de colheita do caderno de campo. Cada produtor
recebe a própria cópia, em branco, na própria conta Google, e a preenche no celular ou
no computador.

O repositório guarda só o **modelo**: o código Apps Script que monta a planilha. Nenhum
dado de produtor faz parte dele.

## O que a planilha tem

- **Geral** - totais do ano (pacotes, unidades, dias de colheita, produtos), colheita
  por mês, produtos mais colhidos, distribuição por destino e comparação do último mês
  com colheita contra o anterior. Um seletor **PERÍODO: de ... até ...** recorta os
  totais, os rankings e os destinos.
- **Jan** a **Dez** - uma aba por mês. Em cima, o painel do mês (totais, mais colhidos,
  por destino, gráfico por semana). Embaixo, a tabela de registro.
- **Produtos** - cadastro de produtos (com unidade padrão) e de destinos. Alimenta as
  listas suspensas. Ativo = "não" tira da lista sem apagar o histórico.
- **Configuração** - nome, núcleo, grupo, rede ou associação, e o ano da planilha.
- **Leia-me** - instruções para quem preenche.
- **_Consolidado** - oculta. Empilha as abas de mês e alimenta a Geral.

A tabela de registro segue a ficha de campo: **Data, Produto, Quantidade, Unidade,
Destino, Observação**. A "Colheita" do papel ("22 pacotes") vira duas colunas para poder
ser somada.

## Estrutura do repositório

O repositório é a própria pasta do clasp: os `.gs` e o `appsscript.json` vão para o
Apps Script; o resto (README, `.gitignore`, exemplos) fica só no git.

```
├── .clasp.json.example     modelo do .clasp.json (o real, com o seu ID, fica fora do git)
├── .claspignore            o que o clasp nunca envia ao Apps Script (Dados.gs)
├── .gitignore              o que nunca vai para o git (.clasp.json, Dados.gs)
├── README.md
├── appsscript.json
├── Config.gs               paleta, layout, helpers, Configuração, modelo x caderno
├── Listas.gs               abas Produtos, Configuração e Leia-me
├── Mes.gs                  aba de mês + onEdit
├── Geral.gs                aba _Consolidado e a Geral (com o seletor de período)
├── Estilo.gs               estilizar(): refaz a formatação sem mexer nos dados
├── Diagnostico.gs          bisseção das opções de gráfico (submenu Teste)
└── Menu.gs                 menu "Colheita", nova planilha, importar caderno
```

Ficam **fora** do repositório (ver `.gitignore`): o seu `.clasp.json`, com o ID do seu
modelo, e qualquer caderno real transcrito (`Dados*.gs`).

## Instalar o clasp

O [clasp](https://github.com/google/clasp) é a ferramenta do Google que envia o código
desta pasta para o Apps Script. Testado com Node.js 24 e clasp 3.3.

1. Instale o **Node.js** (versão LTS) em <https://nodejs.org>. Confira:

   ```bash
   node --version
   ```

2. Instale o clasp:

   ```bash
   npm install -g @google/clasp
   ```

3. **Ligue a API do Apps Script** na sua conta Google, uma vez só, em
   <https://script.google.com/home/usersettings>. Sem isso, o clasp falha com erro de
   permissão.

4. Faça login. Abre o navegador para escolher a conta e autorizar:

   ```bash
   clasp login
   ```

5. **Confira com qual conta ficou logado.** Tem que ser a mesma que é dona da planilha
   modelo. Com a conta errada, o push falha com "The caller does not have permission",
   mesmo com o ID certo:

   ```bash
   clasp show-authorized-user
   ```

   Para trocar de conta: `clasp logout` e `clasp login` de novo, escolhendo a certa.

**No Windows**, se o PowerShell recusar com "a execução de scripts foi desabilitada",
use `clasp.cmd` no lugar de `clasp` (ex.: `clasp.cmd push`), ou rode os comandos pelo
Prompt de Comando (cmd) ou pelo Git Bash.

## Montar o seu modelo

O repositório não aponta para nenhuma planilha: cada um monta a sua.

1. Crie uma planilha em branco no Google Sheets e abra **Extensões > Apps Script**.
2. No editor, em **Configurações do projeto**, copie o **ID do script**.
3. Copie `.clasp.json.example` para `.clasp.json` e cole o ID no
   campo `scriptId`. O `.clasp.json` fica fora do git: o ID é seu.
4. Envie o código e autorize na primeira execução:

   ```bash
   clasp push
   ```

5. Recarregue a planilha e use **Colheita > Manutenção > Nova planilha em branco**.
6. Na aba **Configuração** do modelo, confira o **ano**. Deixe o resto vazio: nome,
   núcleo, grupo e rede são de cada produtor, que preenche na própria cópia.

## Desenvolver

O clasp aponta para o script do **seu modelo**. Mudança se faz aqui e sobe com:

```bash
clasp push
```

Depois, no modelo: **Colheita > Manutenção > Aplicar estilo**, para refazer as abas com o
código novo sem perder nada.

### Testar com dados reais

Nunca no modelo. Usa-se uma **planilha de bancada**, à parte, que tem um caderno real
transcrito num `Dados.gs`. Ela se identifica por uma propriedade do script, que não fica
no código: no editor dela, **Configurações do projeto > Propriedades do script >
`ID_PLANILHA_TESTE` = ID da própria planilha**. Só ali aparece o submenu **Teste**, com
"Importar caderno transcrito" e o diagnóstico de gráficos.

## Distribuir para os produtores

### O modelo

Uma planilha montada com **Colheita > Manutenção > Nova planilha em branco**, sem
`Dados.gs` no script, compartilhada como **"Qualquer pessoa com o link: Leitor"**. Leitor,
nunca editor: ninguém altera o modelo por engano, e a cópia funciona do mesmo jeito.

A Configuração do modelo fica vazia, exceto o ano: nome, núcleo, grupo e **rede ou
associação** são preenchidos por cada produtor na própria cópia.

### O link para o produtor

O endereço do **seu** modelo com `/copy` no lugar de `/edit`:

```
https://docs.google.com/spreadsheets/d/<ID-da-planilha-modelo>/copy
```

O produtor abre, clica em **Fazer uma cópia** (o Google avisa que um script vai junto), e
a planilha abre no Drive dele, com ele como dono. Depois: preencher a **Configuração**
(nome, núcleo, grupo e rede),
cadastrar os destinos na aba **Produtos** e registrar. Para registrar não é preciso
autorizar nada.

Testado de ponta a ponta com uma segunda conta: a cópia nasce no Drive do produtor, sem
`Dados.gs`, sem o submenu Teste, e o modelo continua intacto e em branco.

### Antes de divulgar o link

- **Cada cópia congela o código do dia.** Correções posteriores não chegam às cópias já
  distribuídas. Só divulgar com o modelo estável.
- **O modelo tem um ano**, o da Configuração. Na virada do ano, atualizar o modelo antes
  de mandar o link.
- O produtor precisa de conta Google.

## Menu Colheita

Só aparece no computador. No app do celular não existe menu de script, e o registro não
depende dele.

| Item | O que faz |
|---|---|
| Atualizar resumo geral | Refaz a `_Consolidado` e a Geral. Só é preciso se abas de mês forem criadas ou apagadas. |
| Manutenção > Recriar meses que faltam | Recria abas de mês apagadas, no ano da Configuração. Não toca nas que existem. Não serve para virar o ano. |
| Manutenção > Aplicar estilo | Refaz a formatação **mantendo** registros, cadastro e Configuração. |
| Manutenção > Nova planilha em branco | Apaga tudo e monta vazia, com uma lista base de produtos. Com registros na planilha, exige **digitar APAGAR**. |
| Teste > ... | Só na planilha de bancada (ver "Testar com dados reais"). |

## Decisões que valem lembrar

- **Pacote e unidade nunca se somam.** Os totais aparecem separados em toda parte.
- **As colunas são as da ficha de campo**, na mesma ordem. Uma coluna de Variedade chegou
  a existir e foi removida: a ficha não tem esse campo, e quem preenche copia do papel.
  "Alface crespa" é um produto, como está escrito lá.
- **Listas avisam, mas aceitam nome fora do cadastro.** No calor da colheita, travar o
  registro é pior do que um nome fora do padrão. A exceção é a **quantidade**, que recusa
  o que não for número maior que zero: texto ou negativo sairiam das somas em silêncio.
- **Produto combinado** ("Coentro/Cebolinha") fica como está. Dividir seria inventar
  número que o caderno não tem.
- **Só funções de planilha clássicas** onde dá. `MAP`/`LAMBDA` foram removidos para não
  arriscar o recálculo offline e em aparelhos antigos.
- **Faixa de gráfico justa e sem `chartArea`.** Medido: a opção `chartArea` (px misturado
  com %) deixava os gráficos vazios. Ver `Diagnostico.gs`.
- **Rótulo de eixo em texto**, não data: com data, o eixo vira contínuo e as barras somem.
- **Escrever fórmulas com a planilha em inglês** (`comLocaleNeutro`) e voltar para pt_BR
  no fim: em pt_BR o separador é `;`, e tudo que o script escreve usa `,`.
- **Nada aponta para dados de alguém.** A planilha de bancada se identifica por
  propriedade do script; o nome da rede vem da Configuração.

## Avisos visuais (não bloqueiam o registro)

| Cor | Significado |
|---|---|
| Vermelho claro | Data fora do mês da aba, ou produto sem quantidade. |
| Laranja claro | Registro idêntico a outro do mesmo dia, ou produto fora do cadastro. |

## Demandas futuras

### Impressão para o certificador

**Bloqueio:** falta saber se o organismo de certificação participativa (OPAC) aceita a
ficha impressa, e em que papel: substituindo o caderno, complementando, ou nada.

Perguntas: a ficha impressa é aceita? É por mês ou por período livre ("desde a última
visita")? Precisa de assinatura e de totais por produto?

Proposta, para quando houver resposta: uma aba **Ficha** que reproduz a folha do
caderno (Produto | Colheita | Data | Destino, com "22 pacotes" junto de novo),
cabeçalho com a identificação da Configuração, mês escolhido numa lista, preenchida por
fórmula, sem cores, cabeçalho repetido em cada página. Opcionais: assinaturas, totais por
produto, observações.

Limite de escopo: o caderno de campo tem outras fichas (insumos, manejo); a planilha
cobre só a de colheita e comercialização.

### Excel offline

Decidido seguir no Google Sheets e reavaliar depois de usar com produtores. Se precisar
migrar: gerar o `.xlsx` direto por Python, com doze abas fixas e só funções clássicas. O
que quebra hoje no Excel: `QUERY`, `ARRAYFORMULA`, `FILTER`, `UNIQUE`, `SORT`, o
empilhamento da `_Consolidado` e faixas abertas (`D17:D`).

### Virada de ano

Recomendado: uma planilha por ano. Falta uma opção "Virar o ano" que mantenha
Configuração, produtos e destinos e recrie as doze abas limpas.

### Teste offline no celular

Nunca foi feito. App do Google Sheets, planilha marcada como "Disponível off-line", modo
avião, registrar uma linha e ver se os quadros e a Geral mudam.
