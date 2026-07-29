# Dashboard de Gestão de Contratos

![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)
![Status](https://img.shields.io/badge/status-v1.0.0-brightgreen.svg)
![No backend](https://img.shields.io/badge/backend-none-lightgrey.svg)

Aplicação **client-side, single-file, sem instalação e sem servidor** que
lê e edita uma planilha `.xlsx` de contratos **direto no navegador**, sem
nunca enviar os dados para fora da máquina do usuário.

![Visão Executiva do dashboard](docs/screenshots/executiva.png)

> Os dados mostrados nas capturas de tela e na planilha de exemplo deste
> repositório são **100% fictícios**, gerados só para demonstração.

## Objetivo

Times que gerenciam dezenas ou centenas de contratos de fornecedores
normalmente dependem de uma planilha "viva", mantida manualmente, sem uma
visão consolidada de vencimentos, risco financeiro e qualidade dos dados.
Este dashboard cobre esse gap **sem exigir infraestrutura**: qualquer
pessoa com Microsoft Edge ou Google Chrome consegue abrir a planilha,
visualizar KPIs, cadastrar/editar/excluir contratos e gerar um relatório
estático para compartilhar — tudo a partir de um único arquivo `.xlsx` e
um único arquivo `.html`.

## Principais funcionalidades

- **4 páginas de análise**: Visão Executiva, SLA Operacional, Financeira e
  Gestão de Contratos.
- **14 KPIs calculados** em tempo real a partir da planilha: total de
  contratos, distribuição por status, despesa total/média, contratos
  vencidos críticos (acima de um limiar configurável), vigência longa,
  prazo médio de pagamento, entre outros.
- **Status automático**: calculado a partir da data de fim a cada
  carregamento (No Prazo / A Vencer / Vencido / Em Análise), com opção de
  override manual por contrato quando necessário.
- **Gráficos** (Chart.js): distribuição de status (donut), despesa por
  tipo de serviço e por departamento, top fornecedores por despesa.
- **CRUD completo de contratos**: cadastro, edição e exclusão via
  formulário modal, com busca (sem distinção de acento), filtro por
  departamento e por qualidade de dados, e ordenação de colunas.
- **Flags de qualidade de dados**: contratos sem data de fim, sem valor
  informado, ou com status manual divergente do calculado.
- **Leitura/escrita direta do `.xlsx`** via File System Access API — sem
  upload, sem conversão, sem servidor intermediário. Apenas as colunas
  gerenciadas pelo app são tocadas; todas as outras abas/colunas da
  planilha são preservadas intactas.
- **Backup automático** (1 cópia do estado anterior) a cada salvamento.
- **Snapshot standalone**: gera um HTML autocontido (dados embutidos,
  sem dependências externas) para compartilhar por e-mail ou rede — quem
  abre só visualiza, sem nenhum controle de edição.
- **100% offline**: nenhuma dependência é baixada em tempo de execução
  (sem CDN); tudo que o navegador precisa já está em `assets/`.

## Screenshots

| Visão Executiva | SLA Operacional |
|---|---|
| ![Executiva](docs/screenshots/executiva.png) | ![SLA](docs/screenshots/sla.png) |

| Análise Financeira | Gestão de Contratos |
|---|---|
| ![Financeira](docs/screenshots/financeira.png) | ![Gestão](docs/screenshots/gestao.png) |

## Arquitetura

SPA vanilla JS (sem framework, sem bundler) dividida em módulos por
responsabilidade: `Regras` (lógica de negócio pura), `ExcelIO`
(leitura/escrita do `.xlsx` via File System Access API + SheetJS) e a
camada de estado/renderização/CRUD em `app.js`. O mesmo código roda em dois
modos — **live** (edição via arquivo local) e **snapshot** (visualização
estática, dados embutidos) — sem duplicação de lógica.

Ver [`docs/ARQUITETURA.md`](docs/ARQUITETURA.md) para o fluxo de dados
completo, o racional de não ter backend e os tradeoffs da File System
Access API, e [`docs/DECISOES.md`](docs/DECISOES.md) para decisões
técnicas pontuais (ex.: mitigação de injeção de script na geração do
snapshot).

## Tecnologias

| Camada | Tecnologia |
|---|---|
| UI | HTML5 + CSS3 (design system próprio, sem framework) |
| Lógica | JavaScript (ES5-style, módulos IIFE, sem bundler) |
| Planilha | [SheetJS](https://sheetjs.com/) (`xlsx.full.min.js`), vendorizado |
| Gráficos | [Chart.js](https://www.chartjs.org/) v4.4.0 + [chartjs-plugin-datalabels](https://chartjs-plugin-datalabels.netlify.app/) v2.2.0, vendorizados |
| Leitura/escrita de arquivo | [File System Access API](https://developer.mozilla.org/en-US/docs/Web/API/File_System_API) (Chromium: Edge/Chrome) |
| Build auxiliar | Python 3 (`build_snapshot_template.py`), só para regenerar o molde do snapshot |

Todas as bibliotecas de terceiros são MIT-licensed e ficam vendorizadas em
`assets/` — nada é baixado de CDN em runtime.

## Estrutura do projeto

```
Código/
├── Dashboard_Contratos.html      ← abrir este arquivo para usar o app
├── MATRIZ DE CONTRATOS 2026.xlsx ← planilha de exemplo (dados fictícios)
├── build_snapshot_template.py    ← regenera assets/snapshot-template.js
├── assets/
│   ├── theme.css                 ← design system (paleta, tipografia, componentes)
│   ├── regras.js                 ← lógica de negócio (status, 14 KPIs, normalizações)
│   ├── excel.js                  ← leitura/escrita via File System Access API + SheetJS
│   ├── app.js                    ← estado, navegação, filtros, gráficos, CRUD, snapshot
│   ├── xlsx.full.min.js          ← SheetJS vendorizado
│   ├── chart.umd.min.js          ← Chart.js vendorizado
│   ├── chartjs-plugin-datalabels.min.js
│   └── snapshot-template.js      ← gerado por build_snapshot_template.py (não editar à mão)
└── docs/
    ├── ARQUITETURA.md            ← fluxo de dados, decisões de arquitetura
    ├── DECISOES.md               ← decisões técnicas pontuais (estilo ADR)
    └── screenshots/
```

## Como instalar

Não há instalação nem dependências para baixar. Clone o repositório:

```bash
git clone https://github.com/eholuwi/gest-o-de-contratos.git
cd gest-o-de-contratos
```

## Como executar

1. Use **Microsoft Edge ou Google Chrome** (obrigatório para editar — é a
   única família de navegador com suporte à File System Access API, que
   grava direto no arquivo). Em outro navegador o dashboard abre e mostra
   os dados, mas os botões de editar ficam ocultos.
2. Dê duplo clique em **`Dashboard_Contratos.html`**.
3. Clique em **📂 Abrir planilha** na barra lateral.
4. Selecione a **pasta** do repositório (não o arquivo — a pasta, para que
   o navegador também consiga escrever o backup e o snapshot). O
   Chrome/Edge vai pedir permissão de leitura/escrita nessa pasta uma vez;
   aceite.
5. Navegue pelas 4 páginas e explore os dados de exemplo.

## Como usar

- Cadastre, edite ou exclua contratos na página **Gestão de Contratos**.
- O **status é calculado automaticamente** a partir da data de Fim, toda
  vez que a página carrega. O campo "Status Manual" no formulário existe
  só para forçar um status diferente do calculado em casos excepcionais —
  deixe em "Automático" na maioria dos casos.
- O filtro **Qualidade: Status manual divergente** mostra contratos cujo
  status manual (herdado de planilhas legadas, ou forçado por alguém) não
  bate com o calculado — útil para revisão periódica.

## Como gerar o dashboard (snapshot)

Clique em **💾 Salvar**. Isso faz três coisas de uma vez:

1. grava as alterações de volta no `.xlsx` principal;
2. atualiza o `.backup.xlsx` (cópia de segurança, sempre com o estado
   *anterior* ao salvamento — não é histórico, é 1 cópia só, sobrescrita a
   cada vez);
3. gera/atualiza `Dashboard_Contratos_snapshot.html` na mesma pasta — esse
   arquivo é standalone (sem dependências externas) e pode ser anexado em
   e-mail ou publicado numa rede interna. Quem abrir só visualiza; não há
   nenhum botão de edição.

## Como atualizar

```bash
git pull
```

Se você alterar `theme.css`, `regras.js`, `app.js` ou a estrutura de
`Dashboard_Contratos.html`, **reexecute**:

```bash
python build_snapshot_template.py
```

Isso regenera `assets/snapshot-template.js`, que embute uma cópia em
base64 desses arquivos para que o botão "Salvar" consiga montar o snapshot
standalone. Se você esquecer, o app continua funcionando normalmente — só
o *snapshot* gerado ficará com uma versão desatualizada da interface (os
dados sempre saem atualizados; é só o "molde" visual que fica velho).

## Limitações conhecidas

- **Só funciona para editar em Edge/Chrome** (File System Access API). Em
  Firefox/Safari a visualização funciona, mas os botões Abrir/Salvar ficam
  ocultos e um aviso aparece no topo.
- **Um editor por vez, por decisão de produto** — não há lock real de
  arquivo; nada impede tecnicamente duas pessoas abrirem a pasta ao mesmo
  tempo. Ver [`docs/ARQUITETURA.md`](docs/ARQUITETURA.md).
- **Backup é 1 cópia só**, sobrescrita a cada "Salvar" — não é histórico.
  Não há como voltar duas versões atrás.
- **Sem trilha de auditoria** — o app mostra só o estado atual (flag de
  "Qualidade" na tabela), não quem mudou o quê e quando.
- **Excluir contrato não remove a linha física da planilha** — limpa os
  campos gerenciados daquela linha, para não desalinhar colunas não
  gerenciadas nas linhas vizinhas. Ver [`docs/DECISOES.md`](docs/DECISOES.md).
- **Fluxo de abrir/editar/salvar via File System Access API depende de
  gesto real do usuário** no seletor de pasta do navegador — não é
  automatizável em testes headless tradicionais.

## Roadmap / Funcionalidades futuras

- [ ] Série histórica de despesa mensal (JAN–DEZ) e comparativo
      ano-a-ano, hoje como placeholder aguardando preenchimento da fonte.
- [ ] Integração com ERP para sincronização automática de contratos.
- [ ] Alertas por e-mail para contratos próximos do vencimento.
- [ ] Histórico de versões (além do backup de 1 cópia atual).
- [ ] Trilha de auditoria (quem alterou o quê e quando).

## Licença

Distribuído sob a licença [MIT](LICENSE).

## Autor

**eholuwi** — [github.com/eholuwi](https://github.com/eholuwi)

## Contato

Dúvidas, sugestões ou bugs: abra uma [issue](https://github.com/eholuwi/gest-o-de-contratos/issues).
