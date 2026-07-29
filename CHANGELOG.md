# Changelog

Todas as mudanças notáveis deste projeto são documentadas aqui.

O formato segue [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/),
e este projeto segue [Semantic Versioning](https://semver.org/lang/pt-BR/).

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
