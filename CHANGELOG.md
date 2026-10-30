# Changelog

Todas as mudanças notáveis deste projeto são documentadas aqui.

O formato segue [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/),
e este projeto segue [Semantic Versioning](https://semver.org/lang/pt-BR/).

## [1.9.2] - 2026-10-09

**Protótipo B — "trabalhar conforme hoje".** No app ao vivo, o status e os
"dias até vencer" do **ano corrente** passam a ser calculados com a **data de
hoje** (HOJE) direto sobre o FIM do contrato; a posição da planilha deixa de ser
a régua e vira **retrato** (a nota da barra de ano mostra até quando a planilha
foi recalculada e avisa quando está atrasada). Exercícios fechados continuam
congelados (posição 31/12 daquele ano) e o snapshot mantém o dia do payload.

### Changed
- `app.js`: split de `refAtiva()` em `refPlanilha()` (posição literal) +
  `refAtiva()` (régua): ano corrente usa `HOJE` (`STATUS_CONFORME_HOJE`); anos
  fechados e snapshot seguem a referência da planilha. `notaPosicao()` passou a
  usar `refPlanilha()` para o aviso de posição velha (o status já não depende dela).

> Observação: no modo base, o status em tela pode divergir do status gravado na
> planilha/Power BI enquanto ninguém recalcular (a planilha guarda a posição do
> último recálculo). É o efeito desejado do protótipo — o calendário anda com
> HOJE; o retrato anda quando alguém salva no Excel.

## [1.9.1] - 2026-10-09

**Abrir automático da última pasta.** No navegador, a pasta escolhida em "Abrir
planilha" fica guardada (File System Access API + IndexedDB) e, no próximo carregamento,
a planilha reabre sozinha se a permissão ainda estiver concedida — sem navegar na W:
de novo. Na 1ª vez (sem pasta lembrada), um aviso explicativo único orienta o 1º clique.
Publicado na W: Webapp em 09/10/2026 (hashs conferidos).

### Added
- `excel.js`: `abrirPastaHandle` (fluxo de abertura a partir de um handle já em mãos),
  persistência `guardarPastaPrincipal`/`recuperarPastaPrincipal` (IndexedDB
  `dashboard-contratos`) e `abrirPastaPrincipal` — usa `queryPermission` e só abre sem
  gesto quando a permissão está `granted`; `openFolder` passou a usar o mesmo fluxo e
  grava o handle a cada escolha de pasta.
- `app.js`: `autoOpenNavegador()` no boot (ramo navegador) com toasts — "Planilha reaberta
  automaticamente" ou, sem permissão persistida, aviso para dar um clique em "Abrir
  planilha".
- `app.js` (`updateFileStatus`): a sidebar deixou de exibir o nome cru do arquivo
  ("MATRIZ DE CONTRATOS 2025.xlsx") acima do botão "Abrir planilha"; mostra um resumo
  amigável ("Planilhas 2025 + 2026" / "Base única — exercício 2026") e o nome completo
  apenas como tooltip.

### Nota
- A permissão é por origem do navegador; se for revogada, volta a pedir 1 clique (a
  abertura em si não muda).

## [1.9.0] - 2026-10-09

**Saving de verdade.** A planilha base ganhou a aba `Saving` (tabela `tSaving`, molde
v3), e o app passa a ler e editar o saving dela. Os BIDs e o Saving da Executiva deixam
de ser só demonstração.

### Added
- `regras.js`: `savingBRL`, `savingUSD`, `savingDesde`, `savingMes` e `campoSaving`,
  com a mesma régua da planilha:
  - R$ por mês (`savJan`…`savDez`) mais o total sem detalhe mensal (`savSemMes`);
  - US$ do mês = R$ ÷ cotação do mês; o sem mês ÷ média das cotações do ano.
- `excel.js`:
  - `parseBase` lê também a `tSaving`. A tabela é opcional: planilha sem ela segue
    funcionando.
  - A gravação cirúrgica passa a valer para as duas tabelas (`alteracoesTabela`,
    `aplicarGravado`); `gravarBase` aplica as mudanças por aba (Contratos e Saving).
  - `save(contratos, removidos, {saving, savingRemovidos})`, com as mesmas travas e o
    mesmo backup.
- `app.js`:
  - na tela Gestão, a seção "Saving — BIDs e negociações", com "+ Novo saving",
    Editar e Excluir, num modal com 12 meses e o total sem mês;
  - a Executiva lê o saving real;
  - o snapshot leva o saving.

### Fixed (revisão do 1.8.1)
- O `SAVING` era um vetor preso à **posição** do contrato na lista: depois de incluir,
  excluir ou trocar de ano, o saving ia para o contrato errado. Agora cada item tem a
  sua linha na planilha.
- O ano dos BIDs estava fixo (2025 lia um campo, qualquer outro ano lia o `…26`).
  Agora cada ano lê o próprio arquivo.
- Com a planilha real, a Executiva mostrava BIDs vazio e Saving R$ 0 como se fosse
  dado. Agora mostra o real, ou diz que não há.
- No snapshot, a Home ficava sem fundo; agora há uma cor de reserva no `theme.css`.
- Comentário duplicado e indentação quebrada no `boot()` e no mock.

## [1.8.1] - 2026-10-08

**Executiva**: tabela **BIDs REALIZADOS** — colunas **Nome do Prestador | Tipo de
Serviço | Pagamento | Saving U$ | Saving R$** — mostrando **só o ano selecionado** no
seletor da página, com **total do ano no rodapé em verde**; e gráfico **Saving mensal**
(colunas verticais JAN..DEZ + total do ano), cada um em **card próprio**, em grade
2×2 com Status dos Contratos e Top 10 por Dispêndio. A fonte oficial
(`SAVING - CONTRATOS.xlsx`) traz os valores de 2025 e 2026; a aba `Saving` guarda os
dois anos e a tela alterna pelo seletor.

### Added
- `app.js`:
  - `BIDS`/`SAVING`/`SAVING_ANO` (dados manuais; `SAVING` é array alinhado a
    `CONTRATOS`, cada item = 12 valores; `SAVING_ANO` evita série de outro ano
    quando o seletor de exercício troca).
  - `loadPlanilhaResult` aceita `result.bids` / `result.saving` / `result.savingAno`.
  - Mock (`?mock=1`): `simularBids()` e `simularSaving()` — a tela aparece cheia
    mesmo sem a aba na planilha.
  - `bidsBlockHtml()` (ordena por mês, guarda mês inválido) e `savingMensal()`
    (soma por mês só quando o ano coincide com `SAVING_ANO`).
  - `renderExecutiva`: **cards próprios** na grid 2×2 (Status dos Contratos | Top 10 |
    BIDs REALIZADOS | Saving mensal).
- `theme.css`: `.bids-block/.bids-title/.bids-table/.bids-empty/.saving-title`.
- Snapshot regenerado.

### Nota
As abas `Saving`/`Bids` na planilha da base **ainda não existem** — roadmap
(seção 8 do mapa); hoje a tela mostra os dados sintéticos do `?mock=1`.

## [1.8.0] - 2026-10-08

**Home**: nova página de apresentação com o fundo da empresa (imagem
`assets/home-inventus.jpg`, aprovada em protótipo `html/docs/prototipo_home.html`),
título "Contratos", subtítulo "Gestão de Contratos - Inventus Power" e botões
laranjas de navegação para as quatro telas. Abre como primeira página do app.

### Added
- `app.js`, página Home:
  - `CURRENT_PAGE` inicia em `'home'` (era `'executiva'`).
  - `PAGE_META.home` (título "Home", subtítulo "Apresentação do sistema").
  - `renderHome()` monta o fundo + título + 4 botões de navegação e liga cada
    botão ao `goToPage`. Rende sempre, com ou sem planilha aberta — na Home a
    tela de "nenhuma planilha" não aparece; os botões levam às páginas que
    mostram o estado de cara.
- `theme.css`, bloco "HOME (V1.8.0)": `#page-home.active` em flex sem padding,
  `.home-page` com o fundo `home-inventus.jpg` a cobrir a área, `scrim` em
  gradiente para leitura, título/subtítulo com sombra e botões laranja
  (`--primary`/`--primary-dark`) com hover.
- `Dashboard_Contratos.html`: item "Home" no menu (primeiro, ativo) e a seção
  `#page-home` antes da executiva.
- `assets/home-inventus.jpg` (foto da empresa — dado visual, não versionada
  pelo script de snapshot, apenas referenciada no CSS).

### Changed
- A navegação do menu passa a começar por **Home**.

## [1.7.0] - 2026-10-07

Modo **"base do Power BI"**: o app lê e grava nas mesmas planilhas que o
Power BI lê (`MATRIZ DE POWERBI CONTRATOS <ano>.xlsx`, uma por exercício, na
pasta escolhida ou em qualquer subpasta dela), e mostra os mesmos números.

### Added
- `excel.js`, bloco "Modo base do Power BI":
  - `listarBaseDir` procura as planilhas na pasta e nas subpastas (a mesma
    regra do Power BI). Dois arquivos do mesmo ano dão erro.
  - `parseBase` lê a tabela `tContratos` com o ExcelJS, **só para leitura**.
    Também lê as cotações da aba Cotações e a data de posição da `tMeta`, e
    guarda as linhas livres.
  - `gravarBase` grava **direto no XML da aba** (JSZip), só as células de
    entrada que mudaram. Contrato novo ocupa a 1ª linha livre da tabela, e
    excluir limpa a linha. Ao gravar, liga `fullCalcOnLoad` e marca a
    conferência "recalculada no Excel" como ATENÇÃO, para o Power BI avisar
    até alguém abrir a planilha no Excel e salvar.
  - `saveBase` recusa gravar se a planilha estiver aberta no Excel (arquivo
    `~$`) ou se mudou desde que foi aberta, e grava o backup ao lado.
  - `descartarAlteracoes` faz trocar de ano sem salvar descartar de verdade.
- `assets/jszip.min.js` (JSZip 3.10.2, licença MIT).
- Campos **Código Fornecedor** (1ª coluna da tabela de Gestão, busca por
  código) e **Nome curto** no formulário, no modo base.

### Changed
- `regras.js`: `valorTotalContrato` passa a ser o valor negociado da
  planilha do Power BI (Valor Anual; senão Valor Mensal × 12; senão sem
  valor). `despesaAno` deixa de cair para o Valor Anual: dispêndio é a soma
  dos meses lançados.
- `app.js`, no modo base:
  - status e dias contam da **data de posição** da planilha (`refAtiva`),
    como no Power BI;
  - a barra de ano mostra "Posição em …";
  - o acumulado do ano anterior sai do arquivo daquele ano (`contratosDoAno`);
  - o formulário mostra só os meses do ano aberto;
  - somem "Status manual" e "+ Inserir ano" (ano novo é "Salvar como" no Excel).
- Snapshot: leva código, Nome curto, posição e os meses do ano anterior.

### Por que não ExcelJS para gravar
Num teste com a planilha real, mudando uma célula só, o `writeBuffer`
derrubou as 218 fórmulas de matriz dinâmica (a aba KPI inteira), as 44
colunas calculadas da tabela e 46 de 59 estilos de formatação condicional.
A gravação no XML deixa 29 de 31 partes do arquivo idênticas, e o Excel abre
e recalcula o arquivo sem reparo.

## [1.6.2] - 2026-10-02

Removido o KPI "Vencidos Críticos", que contava contratos vencidos com
dispêndio acima de R$ 100 mil e ocupava um card na Visão Executiva.

- `app.js`: fora o card `Vencidos Críticos (> R$ 100 mil)` da Visão
  Executiva. Os cards "Vencidos", "A Vencer (30d)" e "Em Análise" da faixa
  acima não mudam.
- `regras.js`: fora o limiar `VENCIDO_CRITICO_LIMIAR`, o contador
  `vencidosCriticos` e as duas exportações (`contratosVencidosCriticos` e o
  próprio limiar). `contratosVencidos` foi mantido, porque ainda alimenta o
  card "Vencidos" e a rosca de status. `despesaContratosVencidos` e
  `pctContratosVencidos` continuam exportados, mas já não eram usados pelo
  app antes desta mudança e ficaram como estavam.
- `README.md`: a lista de KPIs não cita mais o limiar configurável.
- `snapshot-template.js`: regerado por `build_snapshot_template.py`.
- O Power BI V2 não tem essa medida nem esse visual, então nada mudou lá. A
  medida continua existindo no PBIP do V1, que está congelado em
  `PowerBI\V1 - antigo (nao usar)` e não foi tocado.

## [1.6.1] - 2026-10-02

Interface sem emojis: os ícones coloridos foram removidos e os botões ficaram só
com o texto.

- `Dashboard_Contratos.html`: fora os `<span class="ic">` do menu (Executiva,
  Status do contrato, Financeiro, Gestão de Contratos) e da barra lateral
  (Abrir planilha, Salvar, Exportar HTML do dia), o `⚠` da faixa de aviso e o
  `⛶` dos dois botões de tela cheia.
- `app.js`: fora o `📂` do estado vazio (o `<div class="big">` foi junto) e o do
  botão "Abrir planilha"; `✏️`/`🗑️` viraram os textos "Editar"/"Excluir", que
  também resolve a dependência de emoji no botão.
- Mantidos de propósito, porque são controle e não decoração: `☰` (menu retrátil),
  `«` (recolher menu), `✕` (fechar modal) e `▲`/`▼` (ordem da tabela).
- `verify.ps1`: removidas as entradas `dashboard-html\assets\core.js` e
  `dashboard-html\assets\data.js`, que não existem desde a mudança de nome do app —
  o gate falhava com "FALTANDO" antes de qualquer checagem.

## [1.6.0] - 2026-09-29

Modo **"planilha por ano"**: pastas com uma `MATRIZ DE CONTRATOS <ano>.xlsx`
por exercício passam a ser detectadas e o seletor de ano troca de planilha,
em vez de blocos internos de um único arquivo multi-ano.

- `excel.js`: regex `ANUAL_FILE_RE` + `anoDoArquivo()`; `openFolder()` e
  `openDesktopDefault()` enumeram os arquivos anuais da pasta
  (`listarAnuaisDir` no navegador, `desktop:readdir` no Electron) e chamam
  `abrirAnuais()` / `abrirAnuaisDesktop()` — cada arquivo é lido com
  `parseWorkbook()` (extraído de `loadFromBuffer`), que devolve um substate
  próprio isolado por ano.
- `armarModoAnual` ativa a planilha do ano corrente quando existe, senão a do
  ano mais recente; `ativarAno(ano)` (exposto na API) troca a planilha ativa e
  devolve `{contratos, fileName, cotacoes, anosMensais, anoAtivo, anosPlanilhas}`.
  `atualizarSubstateSalvo()` mantém o substate do arquivo ativo em sincronia no
  `save()`/`saveDesktop()`: **salvar grava só na planilha do ano selecionado**
  (+ backup `nome.backup.xlsx` ao lado). `resetarModoAnual()` limpa o estado.
- `app.js`: `ANOS_PLANILHAS` é popularizada pelo `loadPlanilhaResult`; em modo
  anual o **seletor lista somente os anos com planilha**, e trocar o ano passa
  a carregar a planilha daquele exercício (`trocarPlanilhaDeAno`, com confirmação
  quando há alterações não salvas). Toasts de abertura citam o ano.
- Back-compat mantido: pasta com um único arquivo multi-ano funciona exatamente
  como antes (blocos por exercício) — o modo anual só se arma quando existe
  mais de um arquivo, ou quando um nome de arquivo casar com `ANUAL_FILE_RE`.
- `desktop/main.js` + `preload.js`: handler/exposta `desktop:readdir`
  (lista os arquivos anuais ao lado do app).
- Validação: `verify.ps1` OK; smoke desktop (`DASH_SMOKE`) com os 19 checks OK.
  O poll de auto-open subiu para 20 s (máquina lenta dava falso negativo acima
  de 10 s) e o assert de render passou a aceitar qualquer número de anos ≥ 1
  (no modo anual o seletor mostra só os anos com planilha).

## [1.5.3] - 2026-09-28

`Regras.ANOS_MENSAIS` (o padrão de anos usado antes de qualquer planilha
ser lida, e pelo boot autocontido do `?mock=1`) deixa de incluir `2024`:
vira `['2026', '2025']`. Pedido do Luis — 2024 não tem dado real (planilha
inteira vazia nesse exercício, ver seção 8 do MAPA) e não precisa aparecer
no seletor de ano da demonstração. Não afeta planilha real: os anos aí
continuam vindo de `detectAnosMensais` (cabeçalho da própria planilha), que
ainda reconhece um bloco `2024` se a Hendrya o mantiver na fonte de dados.
`snapshot-template.js` regerado (mudança em `regras.js`).

## [1.5.2] - 2026-09-28

Modo de **dados simulados** (`?mock=1`) enriquecido e autocontido:

- `boot()` detecta `?mock=1` **antes** de checar `showDirectoryPicker`/
  `DesktopBridge` e gera os contratos direto em memória
  (`gerarContratosSimulados`, 50 razões sociais fictícias) — não depende
  mais de "Abrir planilha". Isso corrige o caso real: em `file://` (duplo
  clique no HTML) o Chrome não expõe `showDirectoryPicker` (API exige
  contexto seguro), então o botão "Abrir planilha" não fazia nada e a tela
  ficava no estado vazio "Nenhuma planilha carregada" — parecia que
  `?mock=1` não funcionava.
- `simularDados` preenche todo campo em branco do contrato (valor anual/
  mensal, departamento, tipo de serviço, prazo de pagamento, data de
  início/fim com distribuição entre os 4 status, comentário) e
  `simularCotacoes` preenche a cotação do dólar dos meses sem lançamento —
  sem sobrescrever o que já estava preenchido, então também funciona sobre
  uma planilha real aberta com `?mock=1`.
- `handleSave` bloqueia o Salvar enquanto `?mock=1` estiver ativo, para não
  gravar valores fictícios de volta numa planilha real.
- `updateFileStatus` mostra "Demonstração (dados simulados)" no rodapé.
- `snapshot-template.js` regerado (mudanças em `app.js`).

## [1.5.1] - 2026-09-28

Rodada de ajustes: abas renomeadas ("Sla Operacional" → **Status do contrato**,
"Financeira" → **Financeiro**); "Top 10 Contratos por Valor" vira **Top 10
Contratos por Dispêndio** (soma dos meses do contrato); Financeira troca os
gráficos de Departamento e Evolução por **Dispêndio Mensal Acumulado** (2025 e
2026, em colunas); KPI "Contratos sem Valor Informado" removido; modo de
**dados simulados** (`?mock=1` ou `window.DASH_MOCK`) para visualização.

### Added

- **`Regras.somaMesesBRL(contrato, ano)`** (já existente) passa a alimentar o
  dispêndio: `topContratos(contratos, n, por)` agora aceita
  `por ∈ 'valor' | 'despesa' | 'dispendio'`.
- **`simularDados(contratos)`** em `app.js`: gera valores sintéticos
  determinísticos (FNV-1a) para os meses de 2025/2026 de cada contrato quando
  `SIMULAR` está ativo; 15% dos contratos ficam sem lançamentos (inativos).
  Nada é gravado na planilha.
- **`serieDispendioAcumulado(contratos, ano)`** em `app.js`: série acumulada
  mês a mês para os cards de dispêndio.
- **Modo simulação**: flag `SIMULAR` (parâmetro `?mock=1` ou `window.DASH_MOCK`);
  o status do arquivo ganha o sufixo "(dados simulados)" e o Demonstrativo da
  Financeira exibe o selo `[DADOS SIMULADOS]`.
- **Financeira**: cards "Dispêndio Mensal Acumulado 2025" (`chart-fin-disp-passado`,
  cor #1F6FB2) e "Dispêndio Mensal Acumulado 2026" (`chart-fin-disp-atual`, cor
  #F58220), desenhados como **gráfico de colunas** (`drawBar`, sem rótulos).

### Changed

- Renomeações de aba: `sla` → **Status do contrato** e `financeira` → **Financeiro**
  (nav e `PAGE_META`).
- "Top 10 Contratos por Valor" (Executiva `chart-exec-contratos` e Financeira
  `chart-fin-contratos`) passam a **Top 10 Contratos por Dispêndio**, ordenados
  por `somaMesesBRL` do ano ativo.
- Card "Despesa por Departamento" e "Evolução 2024/2025/2026" removidos da
  Financeira; no lugar, os dois cards de **Dispêndio Mensal Acumulado** com anos
  fixos (ano corrente 2026 e anterior 2025 via `anoAtual()`).

### Removed

- Financeira: KPI "Contratos sem Valor Informado"; cards `chart-fin-depto`
  (e uso de `Regras.sumBy`) e `chart-fin-evolucao`.
- `app.js`: helper `drawLine` (o acumulado usa colunas).

## [1.5.0] - 2026-09-28

Ajustes de tela pedidos pela Hendrya: "Despesa Total Anual" passa a ser o
**valor total do contrato**; Top 10 por valor/despesa; seletor de ano em
todas as telas com padrão no ano corrente; cards de Vencidos/Média e gráfico
mensal removidos.

### Added

- **`Regras.valorTotalContrato(contrato, ano)`**: valor anual do contrato
  (`Valor Anual do Contrato` → fallback `Valor Mensal do Contrato` →
  `despesaAno`). A medida `Despesa Total Anual` (KPIs e card "Valor Total dos
  Contratos") passa a somar isso. Hoje 83/84 contratos ainda estão com o
  valor não preenchido na planilha — o número só se diferencia do gasto
  quando a coluna começar a ser lançada.
- **`Regras.anoAtual()`**: ano corrente do sistema, usado como padrão do
  seletor de ano.
- **Barra de seletor de ano em todas as telas** (`barraAnoPagina()` +
  `renderSeletoresAno()`/`wireSeletorAno()` em `app.js`), sincronizada com o
  seletor da barra superior; anos sem colunas na planilha aparecem marcados
  como "em preparação" via `anoSemColunas(ano)`.
- **`topContratos(contratos, n, porValor)`** em `app.js`: top n por valor
  total do contrato ou por despesa do ano.
- **Executiva**: card "Top 10 Contratos por Valor" (`chart-exec-contratos`),
  no lugar do Ranking de Prestador de Serviço.
- **Financeira**: card "Valor Total dos Contratos", card "Total Faturado
  <ano>" (novo KPI `Regras.despesaTotalAno`), card "Contratos sem Valor
  Informado", tabela "Top 10 Contratos por Despesa" (`#fin-top10-tbody`) e
  gráfico "Top 10 Contratos por Valor" (`chart-fin-contratos`).
- Smoke test em `desktop/main.js`: exige padrão no ano corrente, seletor em
  cada tela e `render.anos.length >= 3`.

### Changed

- `ANO_ATIVO` em `app.js` deixou de ser `const`; cada tela renderiza a própria
  barra de ano e todos os seletores leem/escrevem o mesmo estado.
- Variável `ANO_SEM_COLUNAS` aposentada: o estado agora é derivado por
  `anoSemColunas(ano)`.
- "+ Inserir ano" (Gestão) subiu para a barra de filtros e o `prompt` usa
  `anoAtual()` como valor inicial.
- `#ano-ativo` na topbar ganhou a classe `ano-select`; estilos da barra de ano
  em `theme.css` (`.ano-bar*`).

### Removed

- Executiva: card "Vencidos" e "Ranking de Prestador de Serviço Top 10".
- Financeira: gráfico de gasto mensal, cards "Despesa Média por Contrato" e
  "Despesa Contratos Vencidos".
- Gestão: gráfico mensal (`graficoMensalCard`, `wireGraficoMensal`,
  `renderGraficoMensal`, `contratosDoGraficoMensal`).
- Código morto em `app.js`: `drawGroupedBar`, `distinctEmpresas`,
  `placeholderBlock`.

## [1.4.0] - 2026-08-11

A planilha deixa de ser "o ano de 2026 mais umas colunas históricas" e passa a ter
um **bloco regular de 26 colunas por exercício** — 12 meses em R$, 12 em dólar e
o par `Total Faturado <ano> R$` / `Total Faturado <ano> $`. A Hendrya lança o mês
no Excel e o app entende; cria um ano no app e ele nasce na planilha.

### Added

- **Despesa mensal por ano (2026, 2025, 2024)**: cada ano tem o seu bloco de
  colunas mensais e a sua cotação do dólar por mês (rótulo na linha 1, valor na
  linha 2 da coluna "$", como já era em 2026).
- **`Total Faturado <ano>` em R$ e em US$** fechando o bloco de cada ano,
  gravado como `=SOMA(...)` dos 12 meses daquele ano — a planilha continua viva.
- **Botão "+ Inserir ano"** (página Gestão): cria as 26 colunas de um ano novo no
  extremo direito da aba, com os rótulos de cotação, e o app já passa a ler/gravar
  ali. O caminho inverso é automático: um ano acrescentado à mão no Excel
  (`JAN 2027 R$`…) é detectado no cabeçalho ao abrir, sem recompilar nada.
- **Seletor de exercício na barra superior** — estado único de ano do app. Trocar
  o ano muda junto os KPIs, o Ranking de Prestador, o Dispêndio do SLA, os
  gráficos da Financeira e as 24 colunas mensais da tabela de Gestão.
- **Card "Contratos com Prazo de Pagamento Negociado"** na Visão Executiva:
  contagem (e % do total) dos contratos com a coluna `Prazo Negociado` preenchida.
- Formulário de contrato com um bloco recolhível de 12 meses **por ano**, com o
  valor em dólar de cada mês e o total de cada exercício calculados ao vivo.
- **Migração de layout automática e idempotente** ao abrir a planilha: leva o
  arquivo ao formato regular sem intervenção manual. Só altera o workbook em
  memória — nada vai a disco sem um "Salvar", que grava o `.backup.xlsx` antes.

### Changed

- **Colunas renomeadas na planilha**: `Valor Mensal 2` → **Valor Mensal do
  Contrato**; `Valor Anual2` → **Valor Anual do Contrato**; `TOTAL ANUAL` →
  **Total Faturado**; `JAN R$`/`JAN $` → **`JAN 2026 R$`/`JAN 2026 $`** (idem
  para os 12 meses). Os nomes antigos continuam sendo aceitos como *alias*, então
  uma cópia antiga da planilha ainda abre.
- **A coluna solta `Total Faturado` saiu do meio dos campos fixos**: o total
  agora mora no fim do bloco de cada ano. `Total Anual Reais`/`Total Anual
  Dolares`, que já ficavam logo após `DEZ 2026 $`, viraram
  `Total Faturado 2026 R$`/`$` — nenhuma célula de dado se moveu por causa disso.
- **Um grupo `[+]` por ano no Excel**: os meses continuam recolhíveis, mas como o
  par de totais fica fora do grupo, o Excel fecha um grupo por exercício em vez
  de um único bloco.
- `computeKPIs` passa a receber o ano; toda medida financeira (despesa total,
  média, de vencidos, em dólar) é do exercício selecionado. As medidas de
  contagem e status seguem independentes de ano.
- Rankings e tabelas passam a usar `Regras.despesaAno(contrato, ano)` no lugar da
  coluna `Valor Anual do Contrato` crua — que só existe para o ano corrente.

### Fixed

- **Perda silenciosa do Total Faturado ao inserir um ano.** O header do total era
  genérico (`Total Faturado`) e valia para o ano mais novo: ao inserir 2027, ele
  voltava a casar com a coluna do total de 2026 e o Salvar seguinte a
  sobrescrevia com a soma de 2027. O header passou a levar o ano, e o
  "+ Inserir ano" cria a coluna de total do ano novo.
- **Detecção de ano casava com o cabeçalho `FIM`.** O padrão de cabeçalho legado
  era "três letras quaisquer", então qualquer planilha declarava ter colunas de
  2026 — inclusive uma que só tivesse 2027. Agora só as 12 chaves de mês contam.
- **96 das 97 linhas da planilha estavam ocultas** (resíduo de um autofiltro
  salvo): quem abria o `.xlsx` no Excel via só o cabeçalho e uma linha. O Salvar
  passa a desocultar todas as linhas.
- **Rótulos soltos na faixa de cotação** (`2026` e `FATURAMENTO` perdidos na
  linha 2) eram formatados como se fossem cotação. A linha 2 agora só aceita
  valor nas colunas "$"; os rótulos de cotação que faltavam (SET–DEZ de 2026)
  passam a ser criados sozinhos.
- `colspan` da tabela de Gestão era literal (`37`/`38`) e saía de sincronia com o
  cabeçalho; agora é calculado.

## [1.3.0] - 2026-08-06

O Salvar deixa de degradar a planilha e passa a **entregar** um `.xlsx` formatado.

### Changed

- **Motor de leitura/escrita do `.xlsx`: SheetJS → ExcelJS.** O SheetJS community
  não grava estilo — cor de fundo, negrito, borda e painel congelado eram
  descartados na escrita, e as larguras de coluna se perdiam a cada Salvar. A API
  pública do `ExcelIO` não mudou; só a camada de I/O.
- **Salvar agora formata a planilha** (`aplicarFormatacao`, idempotente):
  cabeçalho laranja Inventus com fonte branca e quebra de linha, painel congelado
  no cabeçalho + Razão Social, autofiltro, larguras por tipo de coluna,
  `R$ #.##0,00` nas colunas em reais, `$ #.##0,00` nas em dólar, `dd/mm/aaaa` nas
  datas, zebra e borda nas linhas de dados.
- **As 24 colunas mensais viram um grupo recolhível** (`outlineLevel`): a planilha
  abre com 17 colunas em vez de 41, e os meses expandem no `[+]` do Excel.
- Datas passam a ser gravadas às 00:00 (antes vinham com o ruído `00:00:04` da
  planilha de origem) e no formato brasileiro (a fonte estava em `mm-dd-yy`).

### Removed

- `assets/xlsx.full.min.js` (SheetJS, ~861 KB) — substituído por
  `assets/exceljs.min.js` (ExcelJS 4.4.0, MIT, vendorizado).

### Fixed

- **Larguras de coluna e alturas de linha eram perdidas a todo Salvar** — a
  planilha voltava com todas as colunas no tamanho padrão.
- **Deslocamento de data por fuso horário**: o ExcelJS entrega a data da célula em
  UTC, e lê-la em horário local jogava o vencimento para o dia anterior
  (`29/10/2026` virava `28/10`), mudando o status calculado do contrato. Leitura e
  escrita agora normalizam explicitamente (`cellToDate` / `dateToCell`).

## [1.2.0] - 2026-08-06

Acompanha a reestruturação da planilha-fonte: as 12 colunas mensais viraram
**24 colunas pareadas** (`JAN R$` / `JAN $` … `DEZ R$` / `DEZ $`), mais
`Total Anual Reais` / `Total Anual Dolares`, com **uma cotação do dólar por mês**.

### Added

- **Cotação do dólar por mês**, lida da linha 2 da planilha (na coluna "MÊS $"
  de cada mês). Todo valor em dólar passa a usar a cotação do próprio mês.
- **Gravação por fórmula**: ao salvar, as colunas "MÊS $" recebem
  `=SEERRO(<mês R$>/<cotação do mês>;"")` e os dois totais anuais recebem
  `=SOMA(...)` — a planilha continua viva, recalculando sozinha quando a
  cotação é atualizada no Excel.
- **Botão "Exportar HTML do dia"**: gera `Dashboard_Contratos_AAAA-MM-DD.html`
  standalone (somente leitura) ao lado da planilha, sem sobrescrever o snapshot
  corrente. Funciona no navegador e no `.exe`.
- **Sidebar retrátil** (só ícones) e **modo tela cheia** (`Esc` para sair).
- **Gráfico "Gasto Mensal"** na página Gestão: colunas mês a mês em R$ e em US$
  (eixos independentes), com recorte por Ano, Empresa e Tipo de Serviço.
- Total de contratos escrito no centro da rosca "Status dos Contratos".
- Card **"Despesa Total Anual em Dólar"** (Executiva e Financeira).

### Changed

- **Tabela da página Gestão** vira o espelho da planilha: Razão Social, Tipo de
  Serviço, Departamento, Prazo Pagamento, Prazo Negociado, Valor do Contrato,
  TOTAL ANUAL R$, Total Anual $, VIGÊNCIA, INÍCIO, FIM, STATUS., Comentários e
  as 24 colunas mensais, com rolagem horizontal e a Razão Social fixa.
- Formulário de contrato passa a ter os 12 campos mensais em R$, com o valor em
  dólar de cada mês e os dois totais anuais calculados ao vivo. `Valor Anual2`
  volta a ser campo livre (deixou de ser derivado dos meses).
- "Top 8 Tipo de Serviço" → **"Ranking de Prestador de Serviço Top 10"**.
- "Sem data de fim" → **"Vencimento Indefinido"**.

### Removed

- Campo único **"Taxa de Câmbio (R$/US$)"** da Financeira, que gravava em
  `AD2`/`AE2` — células que no layout novo são `AGO R$`/`AGO $`. As cotações
  mensais o substituem.
- Container **"⚠ Alertas"** da Executiva e o card "Contratos Sem Valor Informado".
- Card **"Contratos Vigência Longa (>3 anos)"** da Financeira.
- Coluna e filtro de **Qualidade** na tabela de Gestão.

### Fixed

- **Botão "Salvar" não gravava no `.exe`**: o guard exigia `dirHandle`, que só
  existe no modo navegador — no modo desktop o app opera por caminho de arquivo.
- Casamento de cabeçalho confundia `JAN $` com a coluna legada `JAN` (ambos
  normalizavam para `jan`), fazendo o app ler o mês da coluna errada.

## [1.1.0] - 2026-07-31

### Added

- **Executável único `.exe` portátil** (Electron + electron-builder,
  `desktop/`): roda em outro PC sem instalar nada (Chromium embutido; sem
  depender de Edge/Chrome, Node ou Python no destino).
- **Auto-open da planilha**: ao abrir o executável, o app tenta carregar
  automaticamente a `MATRIZ DE CONTRATOS 2026.xlsx` que estiver ao lado do
  `.exe`; se não encontrar, abre um seletor de arquivo (`.xlsx`, não pasta).
- **Modo desktop em `excel.js`**: leitura/escrita por caminho via ponte
  mínima Electron (`contextBridge`), mantendo backup + snapshot no mesmo
  diretório da planilha. Modo navegador (File System Access API) intacto.
- Smoke test automatizado (`DASH_SMOKE=1`) que exercita auto-open, save com
  backup e snapshot sobre cópias temporárias — nunca toca na planilha real.

### Changed

- `verify.ps1` agora valida também `desktop/main.js`, `desktop/preload.js`
  e `desktop/scripts/sync-app.mjs` (node --check).
- `app.js`: no boot, em modo desktop, tenta abrir a planilha padrão e avisa
  via toast quando não encontra.

## [1.0.0] - 2026-07-29

Primeira versão pública.

### Added

- Dashboard client-side single-file para gestão de contratos, sem backend
  e sem instalação, lendo/escrevendo `.xlsx` direto no navegador via File
  System Access API + SheetJS.
- 4 páginas de análise: Visão Executiva, SLA Operacional, Análise
  Financeira e Gestão de Contratos.
- 14 KPIs calculados em tempo real (status, despesa total/média,
  vencidos críticos, vigência longa, prazo médio de pagamento, entre
  outros).
- Cálculo automático de status (No Prazo / A Vencer / Vencido / Em
  Análise) a partir da data de fim, com override manual opcional por
  contrato.
- CRUD completo de contratos com busca, filtros e ordenação.
- Flags de qualidade de dados (sem data, sem valor, status manual
  divergente).
- Gráficos (Chart.js): distribuição de status, despesa por tipo de
  serviço/departamento, top fornecedores.
- Backup automático (1 cópia do estado anterior) a cada salvamento.
- Geração de snapshot HTML standalone (somente leitura) para
  compartilhamento por e-mail/rede.
- Script auxiliar (`build_snapshot_template.py`) para regenerar o molde
  do snapshot após mudanças de UI.
- Documentação de arquitetura e decisões técnicas em `docs/`.
- Planilha de exemplo com dados fictícios para demonstração imediata.
