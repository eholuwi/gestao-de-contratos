# Arquitetura

## Visão geral

O Dashboard de Gestão de Contratos é uma **SPA client-side sem backend e sem
build step**: um único `Dashboard_Contratos.html` carrega alguns arquivos
JS/CSS estáticos de `assets/` e roda inteiramente no navegador. Não há
servidor, não há API, não há banco de dados — a "fonte da verdade" é um
arquivo `.xlsx` local, lido e escrito diretamente pelo navegador.

```
.xlsx (pasta local)
   │  ExcelIO.openFolder()  — File System Access API + ExcelJS
   ▼
CONTRATOS[] (estado em memória, normalizado por Regras.*)
   │  render das 4 páginas (KPIs, gráficos, tabelas)
   ▼
UI (Executiva / SLA / Financeira / Gestão de Contratos)
   │  CRUD do usuário marca estado "sujo"
   ▼
ExcelIO.save() — grava backup, grava .xlsx principal, gera snapshot HTML
```

## Por que sem backend

A arquitetura só funciona sem servidor porque três condições seguram esse
design:

1. **Leitura é sempre um gesto explícito do usuário** (seletor de
   pasta/arquivo) — não há necessidade de um processo sempre ativo
   escutando mudanças.
2. **A escrita é feita pelo próprio navegador**, via File System Access API
   (`showDirectoryPicker` + `FileSystemFileHandle.createWritable`),
   eliminando a necessidade de upload/download manual ou de um servidor
   intermediário.
3. **Um único editor por vez, por decisão de produto** — não há
   necessidade de lock de arquivo real nem de resolução de conflitos
   concorrentes, porque o fluxo de trabalho pressupõe uma pessoa
   responsável pela planilha por vez.

Se qualquer uma dessas três premissas deixar de valer (múltiplos editores
simultâneos, necessidade de auditoria centralizada, acesso fora da rede
local), a arquitetura precisa mudar para um modelo cliente-servidor
tradicional.

## Tradeoffs da File System Access API

- Suportada apenas em navegadores Chromium (Edge, Chrome) — Firefox e
  Safari não implementam a API. O app detecta isso em tempo de execução e
  esconde os botões de editar, mostrando um aviso.
- Pedir acesso a uma **pasta** (não a um arquivo isolado) permite, com uma
  única concessão de permissão do usuário, também escrever o backup e o
  snapshot HTML na mesma pasta — evita 3 diálogos de permissão separados.
- O fluxo de abrir/editar/salvar depende de gesto real do usuário no
  seletor de pasta do navegador; não é automatizável por scripts headless
  (Puppeteer, testes E2E tradicionais). A validação desse fluxo específico
  precisa ser manual.

## Padrão "um app, dois estados" (live vs. snapshot)

Em vez de dois artefatos separados (um "editor" e um "visualizador"), o
mesmo código roda em dois modos:

- **Modo live**: dados vêm de `ExcelIO.openFolder()`; botões de
  Abrir/Salvar e as ações de CRUD ficam visíveis (classe `.live-only`).
- **Modo snapshot**: os dados já vêm embutidos no HTML como
  `window.DASH_CONTRATOS` (gerados no momento do "Salvar"); os elementos
  `.live-only` ficam ocultos e o app roda 100% somente leitura.

Os módulos de renderização e cálculo (`Regras.*`, funções de página em
`app.js`) são os mesmos nos dois modos — a única diferença é a origem dos
dados e a visibilidade dos controles de edição. Isso evita duplicar lógica
entre "versão para editar" e "versão para distribuir".

## Geração do snapshot autocontido

O botão "Salvar" gera um HTML standalone (sem `<script src>` externos) para
ser distribuído por e-mail ou rede. Ele embute em base64: CSS, Chart.js, o
plugin de datalabels, a lógica de negócio, a lógica de app e o HTML do
shell — tudo pré-calculado em `assets/snapshot-template.js` por
`build_snapshot_template.py`, e reidratado em runtime por um pequeno
bootstrap que decodifica cada bloco e injeta como `<script>`/`<style>`.

Rode `python build_snapshot_template.py` de novo sempre que
`theme.css`, `regras.js`, `app.js` ou a estrutura de
`Dashboard_Contratos.html` mudarem — ver `README.md`, seção de manutenção.

## Preservação do round-trip do Excel

O app só entende um conjunto fixo de colunas "gerenciadas" (razão social,
tipo de serviço, departamento, prazos, valores, vigência, datas, status,
comentários — ver `COLUNAS_GERENCIADAS` em `assets/excel.js`). Todas as
outras colunas e abas da planilha (totais mensais, séries históricas,
outras tabelas) são lidas para a memória mas nunca reescritas — o workbook
inteiro é mantido em memória (via ExcelJS) e só as células gerenciadas são
alteradas antes de gravar de volta. Isso permite que o dashboard conviva
com uma planilha que também é usada por outras ferramentas (ex.: um
relatório de BI) sem corromper dados que ele não entende.

"Excluir" um contrato não remove a linha física — limpa apenas os campos
gerenciados daquela linha (inclusive a razão social, que é o sinal usado
no parser para pular a linha na próxima leitura). Isso evita desalinhar
colunas não gerenciadas nas linhas vizinhas.

## Base do Power BI: gravação cirúrgica (V1.7.0)

As planilhas do Power BI (`MATRIZ DE POWERBI CONTRATOS <ano>.xlsx`) têm
tabelas do Excel, colunas calculadas e fórmulas de matriz dinâmica: a aba
KPI, que o Power BI lê, é toda feita delas. O ExcelJS lê esse arquivo bem,
mas **não o regrava sem estrago**. Num teste com a planilha real, mudando
uma célula só, o `writeBuffer` perdeu:

- as 218 fórmulas de matriz dinâmica;
- as colunas calculadas da tabela;
- a maior parte da formatação condicional.

Por isso, nessa base, o ExcelJS só **lê**. Para gravar, `gravarBase()`
(`assets/excel.js`) abre o `.xlsx` como zip (JSZip) e troca o valor de cada
célula alterada direto no XML da aba Contratos:

- texto vai como `inlineStr`, e data como número de série;
- o estilo da célula é mantido;
- o resto do pacote sai idêntico.

O arquivo é preparado para não precisar criar nada:

- **Linhas livres.** A tabela já nasce com linhas vazias que têm as fórmulas
  (o molde é do `montar_base_pbi.py`, fora deste repositório). Contrato novo
  ocupa a 1ª livre, e excluir limpa a linha. O app nunca cria linha, fórmula
  nem redimensiona a tabela.
- **Excel recalcula, Power BI avisa.** O Power BI lê o valor que o Excel
  calculou, e o app não recalcula. Ao gravar, o app liga `fullCalcOnLoad` e
  escreve "ATENÇÃO" no valor calculado de uma conferência da planilha (nomes
  `pHtmlResultado`/`pHtmlStatus`). O Power BI fica vermelho até alguém abrir
  a planilha no Excel e salvar; o recálculo devolve "OK" sozinho.

Testado fora do navegador pelo `conferir_html_base.js` (raiz do projeto
privado): os KPIs do app têm de bater com a `tCards` da planilha, a gravação
de teste só pode mexer no que deve, e o Excel abre, recalcula e não "repara"
o arquivo.

## Riscos conhecidos e mitigações

| Risco | Mitigação atual |
|---|---|
| Dois programas abrindo o mesmo `.xlsx` ao mesmo tempo (ex.: Excel + navegador) | Layout antigo: nenhuma automática, só o aviso na interface. Base do Power BI (V1.7.0): o Salvar recusa se houver `~$<arquivo>` na pasta (aberto no Excel). |
| Um editor sobrescrever o trabalho de outro | Layout antigo: decisão de produto (um editor por vez). Base do Power BI: o Salvar recusa se o arquivo mudou (data/tamanho) desde a abertura; recusa, não mescla. |
| Power BI com número velho depois de gravar pelo app | Base do Power BI: conferência marcada como ATENÇÃO + `fullCalcOnLoad`; o aviso some quando a planilha é aberta e salva no Excel. |
| Perda de dados ao salvar | Backup automático de 1 cópia (estado anterior ao save) — não é histórico completo, não substitui controle de versão externo. |
| Sem trilha de auditoria | O app mostra flags de qualidade do estado atual, mas não quem mudou o quê e quando. |
| Navegador incompatível (Firefox/Safari) | Detecção em runtime + banner + ocultação dos botões de edição; leitura/visualização continuam funcionando. |
