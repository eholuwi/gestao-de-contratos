# Contribuindo

Obrigado pelo interesse em contribuir! Este é um projeto pequeno,
sem build step e sem framework — o objetivo é manter assim.

## Como propor uma mudança

1. Abra uma [issue](https://github.com/eholuwi/gest-o-de-contratos/issues)
   descrevendo o bug ou a funcionalidade antes de investir tempo numa
   implementação grande — evita retrabalho caso a abordagem precise
   ajustar.
2. Faça um fork do repositório e crie uma branch a partir de `main`:
   `git checkout -b feat/nome-da-mudanca`.
3. Siga o estilo de código já existente (JS ES5-style, sem bundler,
   módulos IIFE — ver `assets/app.js`, `assets/regras.js`, `assets/excel.js`).
4. Se você alterar `theme.css`, `regras.js`, `app.js` ou a estrutura de
   `Dashboard_Contratos.html`, rode `python build_snapshot_template.py`
   antes de commitar, para manter `assets/snapshot-template.js` em dia.
5. Teste manualmente no navegador (Edge ou Chrome) — abrir a pasta do
   repositório, navegar pelas 4 páginas, e, se sua mudança afeta CRUD ou
   salvamento, testar o fluxo completo de abrir → editar → salvar.
6. Descreva no Pull Request **o que mudou e por quê**, e inclua
   screenshots se a mudança for visual.

## Commits

Este projeto usa [Conventional Commits](https://www.conventionalcommits.org/pt-br/):
`feat:`, `fix:`, `docs:`, `style:`, `refactor:`, `chore:`, `build:`.

## Reportando bugs

Use o template de [bug report](.github/ISSUE_TEMPLATE/bug_report.md) e
inclua: navegador/versão, passos para reproduzir, e se possível uma
planilha de exemplo (**nunca** anexe dados reais/sensíveis).

## Dúvidas de segurança

Não abra issue pública para vulnerabilidades — ver [SECURITY.md](SECURITY.md).
