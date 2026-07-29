# Política de Segurança

## Versões suportadas

| Versão | Suportada |
|---|---|
| 1.x   | ✅ |

## Modelo de ameaça

Este é um app **client-side, sem backend e sem servidor**. Os dados da
planilha nunca saem da máquina do usuário — não há upload, API ou
armazenamento remoto. As principais superfícies de risco são:

- Geração do snapshot HTML standalone (dados de contrato embutidos no
  HTML gerado) — ver mitigação de injeção de script em
  [`docs/DECISOES.md`](docs/DECISOES.md#3-mitigação-de-injeção-de-script-na-geração-do-snapshot).
- Bibliotecas de terceiros vendorizadas (SheetJS, Chart.js,
  chartjs-plugin-datalabels) — atualizadas manualmente, não via
  gerenciador de pacotes.

## Reportando uma vulnerabilidade

Se você encontrar uma vulnerabilidade de segurança, **não abra uma issue
pública**. Em vez disso, entre em contato diretamente através do perfil do
autor no GitHub ([@eholuwi](https://github.com/eholuwi)) ou abra uma
[Security Advisory](https://github.com/eholuwi/gest-o-de-contratos/security/advisories/new)
privada no repositório.

Inclua: passos para reproduzir, impacto esperado, e navegador/versão
utilizado. Você deve receber uma resposta em até 7 dias.
