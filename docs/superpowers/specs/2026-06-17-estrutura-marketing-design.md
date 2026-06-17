# Aba "Estrutura" — Painel visual da máquina de marketing

**Data:** 2026-06-17
**App:** app.paraser.com.br (`index.html`, single file) + Apps Script `meta-capi`

## Objetivo
Aba nova no app que mostra, com números reais, **como a campanha está montada** (estrutura) **e o resultado** (funil). Para o Felipe enxergar a máquina rodando.

## Decisões (brainstorm 17/06)
- **Conteúdo:** estrutura + resultado juntos
- **Onde:** aba nova dedicada (nome sugerido "🧩 Estrutura", ajustável)
- **Formato:** dois blocos empilhados — "A Máquina" em cima, "O Resultado" embaixo
- **Dados:** híbrido — funil usa o snapshot que já existe; tamanhos de público buscam ao vivo

## Bloco A — "A Máquina" (estrutura)
Esteira de cards ligados por setas, cada um com número real:

`Origem (Feegow + Site)` → `Eventos (Schedule · CompleteRegistration · PageView, contagem 7d)` → `Públicos (Clientes/Controle Agendamentos com selo 🟢 auto-sync + última atualização · Engajados)` → `Lookalike 1% (alcance ~1,2M)` → `Campanha ativa (nome + CPL)`

Dados ao vivo via endpoint novo `marketing-setup`.

## Bloco B — "O Resultado" (funil)
Funil em barras, **reaproveitando** o endpoint/snapshot `marketing-funnel` que já existe: impressões → cliques → contato/lead → schedule → cadastro → procedimento. Carrega instantâneo.

## Backend — endpoint `marketing-setup` (Apps Script meta-capi)
Retorna JSON:
- `audiences`: `{ clientes:{name,count,updatedAt}, lookalike:{name,count}, engajados:{name,count} }`
  - via `GET /{id}?fields=name,approximate_count_lower_bound,approximate_count_upper_bound,time_updated` (token META_ADS_TOKEN, já confirmado com permissão)
- `events7d`: `{ pageview, completeRegistration, schedule }`
  - Schedule/CompleteRegistration: reusar `_mktCountCapiEvents_` (lê MetaCapi_Log)
  - PageView: dataset stats API `GET /{dataset}/stats` ou via insights
- `config`: `{ engajado:"CompleteRegistration", cliente:"Controle Agendamentos" }`

IDs: Clientes `120242388951550375` · Lookalike `120242388972990375` · Engajados `120247022590510375` · Dataset `920108941023871`. Chave: `paraser2026`.

## Frontend — nova aba em `index.html`
- Botão de aba no nav + seção nova. Lazy-load ao abrir (padrão `_mktLoaded`).
- Render Bloco A (cards horizontais responsivos + setas, vira coluna no mobile) e Bloco B (barras; reusar lógica de `_mktRenderFunnel`).
- Estilo: tema escuro existente (`var(--card)`, gradiente `#00d4aa→#a855f7`).
- Fontes: `marketing-setup` (máquina) + `marketing-funnel` (resultado), em `Promise.all`.

## Deploy
Atualizar a deployment que o app usa (`AKfycbz3...`, hoje @10) para a última versão do script (ganha `marketing-setup`, mantém os endpoints atuais). `MKT_WEBAPP_URL` no `index.html` já aponta pra ela.

## Fora de escopo
- Não alterar a lógica do funil existente (só reusar).
- Sem edição de campanhas pela aba (somente visualização).
- Sem snapshot novo: máquina é ao vivo, resultado usa snapshot que já roda.
