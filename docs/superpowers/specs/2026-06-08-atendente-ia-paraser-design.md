# Atendente IA Paraser ("ANA") — Design Spec

**Date:** 2026-06-08
**Product:** Sistema de atendimento automatizado WhatsApp + Landing Page + Quiz qualificador + Dashboard
**Target:** Substituir Manychat na Paraser, atender ~100+ leads/dia (Meta Ads) com IA conversacional
**Language:** PT-BR
**Status:** Design aprovado pelo Felipe via brainstorming visual companion (16 telas)

---

## 1. Problem Statement

A Paraser recebe 100+ leads novos por dia via Meta Ads (formulário "Seu sonho de ser mãe começa aqui!"). Hoje, o primeiro atendimento via WhatsApp é feito pelo Manychat (R$ 500+/mês), que opera em árvore de decisão burra e não entende contexto. Resultado:

- Equipe humana sobrecarregada com perguntas repetitivas (preço, médicos, localização)
- Conversão lead → consulta em ~10-15% (abaixo do potencial)
- Custo alto (R$ 500+/mês) pra funcionalidade limitada
- Sem qualificação prévia: cada lead chega na equipe igual

**Objetivo do projeto:** substituir Manychat por sistema próprio com IA conversacional ("ANA"), Landing Page com Quiz qualificador, e Dashboard de gestão. Resultado esperado:

- **Economia direta:** ~R$ 200-300/mês após migração (Claude API < Manychat)
- **Capacidade 5×:** 24/7 + multiparalelo + leads qualificados pré-WhatsApp
- **Conversão maior:** lead chega no WhatsApp já com perfil capturado
- **Experiência humanizada:** persona acolhedora ("Camila") vs auto-responder genérico

---

## 2. Escopo — Fronteiras

### ✅ ANA atende
- Leads NOVOS via Meta Ads, site, Instagram (pré-paciente)
- Dúvidas sobre tratamentos, preços, médicos, processos
- Quiz qualificador conversacional (idade, tempo tentando, urgência)
- Pré-agendamento de 1ª consulta (humano confirma)
- Histórias e educação sobre fertilidade

### ❌ ANA NÃO atende
- Pacientes em tratamento (sigilo médico + LGPD)
- Diagnósticos clínicos
- Confirmação direta de agendamento (sempre humano)
- Negociação de preço/desconto
- Solicitação de prontuário/laudo
- Conversas pós-confirmação (saem do círculo da ANA, vão pra Confirmações Agenda + equipe humana)

### Linha divisória
Quando atendente humana **confirma o pré-agendamento** no Dashboard → contato vira "paciente" → próximas mensagens não passam por ANA, vão direto pra equipe humana.

---

## 3. Decisões Consolidadas (rastreabilidade)

| # | Decisão | Escolha do Felipe |
|---|---|---|
| Q1 | Escopo do MVP | **C** — Tudo (Bot + LP + Quiz + Dashboard) em 2 fases |
| Q2 | Persona do bot | **B** — Humanizada, nome **ANA** |
| Q3 | Capacidades | **B híbrido** — ANA pré-agenda, atendente humana confirma |
| Q4 | Notificação atendente | **C** — Slack #atendimento + Dashboard |
| Q5 | Horário operação | **A** — 24/7 com transparência (humano confirma em horário comercial) |
| Q6 | Tamanho do Quiz | **B** — 5 perguntas (~1 min) |
| Q7 | Cérebro da ANA | **C** — Prompt + memória por paciente |
| Q8 | Stack técnica | **A** — Apps Script + Sheets + GitHub Pages + Claude API |
| Q9 | Red flags | 5 cenários selecionados + ANA limitada a leads NOVOS |
| Q10 | Arquitetura código | **B** — 4 Apps Scripts especializados |

---

## 4. Arquitetura

### Diagrama geral

```
ENTRADAS
├─ Meta Ad "Saiba mais"  → LP/Quiz (GitHub Pages)  ───┐
├─ Meta Ad "Enviar Msg"  → WhatsApp Z-API             ├─→
├─ Site → botão WhatsApp → WhatsApp Z-API             │
└─ Instagram DM (manual) → WhatsApp Z-API             ┘
                                                       ▼
APPS SCRIPTS (cérebro)
├─ ana-quiz       : recebe submissão LP, dispara mensagem WhatsApp inicial
├─ ana-bot        : webhook Z-API + Claude API + memória + escalation
├─ ana-dashboard  : Web App (HTML) com UI da atendente humana
└─ ana-cron       : timeouts, métricas diárias, limpeza, heartbeat Z-API
                                                       ▼
SHEETS (bus + memória)
├─ ANA_Config             : persona, regras, tom (Felipe edita)
├─ ANA_FAQ                : base de conhecimento (atendente cura)
├─ ANA_Conversas          : memória por telefone (sistema escreve)
├─ ANA_Quiz_Submissoes    : capturas do quiz (LP grava via Web App)
├─ ANA_Pre_Agendamentos   : fila pendente (status PENDENTE/CONFIRMADO/RECUSADO)
└─ ANA_Logs               : erros, eventos, métricas brutas
                                                       ▼
SAÍDAS
├─ Z-API           : mensagem volta pra paciente
├─ Slack #atendimento : notificação push pra atendente humana
├─ Dashboard web   : ferramenta de trabalho da pessoa dedicada
└─ Feegow API      : leitura de disponibilidade (read-only)
```

### Stack

| Camada | Tecnologia | Status |
|---|---|---|
| IA | Claude API (Anthropic, modelo Sonnet) | Novo — R$ 200-300/mês |
| Bot WhatsApp | Apps Script + Z-API | Z-API já roda |
| Memória / Bus | Google Sheets | Padrão Paraser |
| LP + Quiz | GitHub Pages (HTML estático) | `app.paraser.com.br/ana` |
| Dashboard | HTML single-page (padrão `index.html`) | `app.paraser.com.br/ana/dashboard` |
| Notificação | Slack webhook (canal `#atendimento`) | Já configurado |
| Hospedagem backend | Apps Script Web Apps | Padrão atual (CRM, Brevo Lotes) |

**Custo infra: R$ 0** (tudo já pago). Custo total: ~R$ 200-300/mês em Claude API.

---

## 5. Cérebro da ANA

### System Prompt (todo chamada Claude inclui)

```
Você é a ANA, atendente do Instituto Paraser (clínica de fertilidade em Botafogo, RJ).

PERSONA:
- Acolhedora, calorosa, mas profissional.
- Trata por "você" (nunca senhora).
- Usa 💜 com parcimônia (1 por conversa, no máximo).
- Frases curtas. Sem listas longas no WhatsApp.

VOCÊ PODE:
- Responder dúvidas sobre tratamentos, preços, médicos, processos
- Qualificar perfil (tempo tentando, idade, etc)
- Pré-agendar consultas (humano confirma depois)
- Contar histórias / educar

VOCÊ NÃO PODE:
- Dar diagnóstico médico
- Confirmar agendamento sozinha (só pré-agenda)
- Negociar preço (sempre passa pra humano)
- Falar sobre paciente existente
- Acessar prontuário/laudo

ESCALAÇÃO IMEDIATA (chama humano AGORA + para de responder):
1. Crise emocional (suicídio, depressão grave)
2. Emergência médica (dor, sangramento, OHSS)
3. Reclamação/conflito ético
4. Tópico controverso (aborto, gestação substituição, etc)
5. Negociação de preço

QUANDO ESCALAR: pare de responder, marque o caso na fila do dashboard
com prioridade ALTA, notifique Slack #atendimento com 🚨 + telefone, espere humano entrar.

QUANDO PERGUNTAREM SE VOCÊ É HUMANA OU IA: admita — "aqui é a Ana, assistente
da Paraser, e respondo com IA. Mas qualquer momento a equipe humana entra."

[FAQ — injetado dinamicamente do Sheets ANA_FAQ]

HISTÓRICO DESTA PACIENTE:
[últimas 10 mensagens — do Sheets ANA_Conversas]
[perfil capturado no quiz — do Sheets ANA_Quiz_Submissoes, se houver]
```

### Memória por paciente — LGPD-friendly

**O que guarda:**
- Telefone (chave) + nome
- Últimas 10 mensagens da conversa (rolling window)
- Perfil do quiz (idade, tempo tentando, urgência, dúvida principal, plano de saúde)
- Status: `NOVO → CONVERSANDO → PRE_AGENDADO → CONFIRMADO → DESCARTADO`
- Última interação (timestamp)

**O que NÃO guarda:**
- Histórico clínico, diagnósticos, exames
- Dados de pagamento
- Conversas após `CONFIRMADO` (saem do escopo da ANA)

**LGPD:**
- Primeira mensagem inclui: *"ao continuar, você concorda com nossa Política de Privacidade"* + link
- Botão "Esquecer meus dados" no Dashboard (atendente pode acionar)
- Retenção: 6 meses sem interação → registro anonimizado (telefone hash, nome removido, mensagens deletadas)

---

## 6. Fluxo do Lead — Passo a Passo

### Cenário A: Lead vem via "Saiba mais" do Meta Ad

1. Lead clica "Saiba mais" → vai pra `app.paraser.com.br/ana`
2. **Landing Page** — hero editorial Paraser ("Vamos descobrir juntas o seu próximo passo") → botão "Começar agora"
3. **Quiz** — 5 perguntas, 1 por tela:
   - P1: Há quanto tempo tenta engravidar?
   - P2: Idade?
   - P3: Já consultou um especialista em fertilidade?
   - P4: Tem plano de saúde ou prefere particular?
   - P5: Qual sua principal dúvida ou expectativa?
4. **Resultado** — *"Tenho algumas ideias pra você"* + botão WhatsApp com payload qualificado
5. `ana-quiz` grava em `ANA_Quiz_Submissoes` e gera link `wa.me/552130344130?text=...&ref=quiz-{id}`
6. Lead clica → abre WhatsApp → mensagem inicial pré-preenchida
7. `ana-bot` recebe webhook Z-API → identifica via `ref=quiz-{id}` → carrega perfil → ANA cumprimenta com contexto: *"Oi, Maria! 💜 Vi que você tenta há 1-2 anos e ainda não consultou um especialista..."*

### Cenário B: Lead vem direto via "Enviar Mensagem" do Meta Ad

1. Lead clica "Enviar Mensagem" → WhatsApp abre com texto padrão
2. `ana-bot` recebe webhook → NÃO tem perfil de quiz → ANA cumprimenta genérico mas humano: *"Oi! 💜 Sou a Ana, do Instituto Paraser. Posso te ajudar com informações sobre fertilidade. Pra eu te orientar melhor — há quanto tempo você está tentando engravidar?"*
3. ANA faz quiz inline durante a conversa (5 perguntas espaçadas naturalmente)

### Pré-agendamento (final do funil ANA)

1. Paciente expressa intenção de agendar
2. ANA → Feegow API (read-only): "qual disponibilidade próximos 7 dias?"
3. ANA propõe 2 opções: *"Tenho terça 14h com a Dra. Marcelle ou quinta 10h com o Dr. Rodolfo"*
4. Paciente escolhe
5. ANA grava em `ANA_Pre_Agendamentos` (status=PENDENTE)
6. ANA → paciente: *"Anotei terça 14h. Em até 4h a equipe te confirma 💜"*
7. ANA → Slack #atendimento: `🔔 Maria Souza (21999999999) → terça 14h Dra. Marcelle · Perfil: 32 anos · 1-2 anos tentando`
8. Atendente abre Dashboard → vê card na Fila → clica "Ver conversa" → clica "✓ Confirmar"
9. Sistema (`ana-dashboard`) → Feegow API (write): cria agendamento real
10. `ana-bot` → paciente: *"✅ Confirmado! Te espero terça 14h. Endereço: Rua Prof. Álvaro Rodrigues 352, 10º andar"*
11. Status: `PENDENTE → CONFIRMADO`. Contato sai do escopo da ANA.
12. Confirmações Agenda (script existente) cuida de D-1 e lembretes.

---

## 7. Dashboard pra Atendente

Live em `app.paraser.com.br/ana/dashboard` — HTML single-page no estilo `index.html` (paleta roxa/dourada/creme, Playfair + Lora).

### 5 abas

| Aba | Conteúdo |
|---|---|
| **Fila** | KPIs do dia + cards de pré-agendamentos pendentes (com perfil do quiz visível) + botões Confirmar/Recusar/Ver conversa. Red flags destacados em vermelho com prioridade ALTA. |
| **Conversas ao vivo** | Lista de conversas ativas (igual WhatsApp Web). Botão "🤝 Tomar conversa" pra atendente assumir manualmente. |
| **Métricas** | Charts.js — leads/dia, taxa pré-agendamento, taxa confirmação, red flags, hora de pico, tempo médio até confirmação. |
| **FAQ** | Tabela Sheets editável inline. Atendente adiciona/edita perguntas. ANA usa atualização em ~5 min. |
| **Config ANA** | Persona, regras, palavras-chave de red flag. Controle Felipe (mudanças críticas pedem confirmação). |

### Quem usa
- **Pessoa dedicada** (a contratar — Roberta no mockup) opera diariamente
- **Felipe** edita Config ANA e revisa amostras

---

## 8. Red Flags (escalação imediata pra humano)

ANA detecta via 2 mecanismos: palavras-chave (lista) + análise de intenção pela Claude. Quando dispara: para de responder, marca caso como `PRIORIDADE_ALTA`, notifica Slack com 🚨.

| # | Cenário | Exemplo gatilho | Resposta ANA |
|---|---|---|---|
| 1 | Crise emocional grave | "não aguento mais", "não vejo saída", menção a suicídio | Acolhe + envia CVV 188 + chama humano |
| 2 | Emergência médica | "sangramento intenso", "dor forte", "febre alta", "OHSS" | Orienta procurar PA + chama plantão Paraser |
| 3 | Reclamação / conflito ético | "vou denunciar", "estou indignada" | Acolhe sem opinar + escala pra Felipe/gestor |
| 4 | Tópico controverso | aborto, gestação substituição, ovodoação anônima, seleção de sexo | Não opina, marca consulta médico |
| 5 | Negociação preço | "desconto?", "parcelar?", "outro lugar é mais barato" | Dá faixa geral + passa pra atendente fechar |

**Não incluídos** (decisão deliberada do Felipe — não ocorrem na fase de primeiro contato):
- Pergunta sobre paciente existente (LGPD) — não cabe pois ANA só atende leads NOVOS
- Solicitação de prontuário/laudo — não cabe pois ANA só atende leads NOVOS

---

## 9. KPIs de Sucesso

### Operação (sistema rodando)
- Uptime ANA ≥ 99%
- Tempo médio resposta ANA < 15 seg

### Qualidade (ANA atende bem)
- % conversas resolvidas sem humano: 60-70%
- Red flags acionados corretamente: ≥ 95%
- Pré-agendamentos confirmados (não recusados): ≥ 75%

### Negócio (gera valor)
- Taxa lead → pré-agendamento: ≥ 20% (Manychat hoje ~10-15%)
- Taxa pré-agendamento → consulta realizada: ≥ 65%

### Experiência (paciente fica feliz)
- % pacientes que reclamam de ser IA: < 5%
- NPS pós-1ª consulta: +10pts vs baseline Manychat

---

## 10. Custos Detalhados

| Item | Mês 1 (migração) | Mês 2+ (steady state) |
|---|---|---|
| Claude API (Sonnet, ~50k msgs/mês) | R$ 200-300 | R$ 200-300 |
| Z-API (já paga) | R$ 0 extra | R$ 0 extra |
| Apps Script + Sheets + GitHub Pages | R$ 0 | R$ 0 |
| Manychat (paralelo durante migração) | R$ 500 | R$ 0 (cancelado) |
| **Total infra técnica** | **R$ 700-800** | **R$ 200-300** |
| **Delta vs Manychat hoje (R$ 500+)** | **+R$ 200-300 temp** | **-R$ 200 a -R$ 300/mês** |

**Payback:** 2-3 meses após cancelar Manychat. **Economia 12 meses:** R$ 2.400-3.600.

---

## 11. Riscos e Mitigações

| Risco | Mitigação |
|---|---|
| **ANA "alucina"** — inventa info errada | FAQ no prompt. Quando não sabe: "não tenho essa info, vou pedir pra atendente". Auditoria de 50 conversas/semana na Fase 1. |
| **Red flag não detectado** — emergência vira problema | Palavras-chave + Claude detecta intenção. Plantão 24h via Slack (notificação celular do plantonista). |
| **LGPD** — paciente reclama do uso de dados | Consentimento explícito no início. Política publicada. Sheets com acesso restrito. "Esquecer meus dados" via Dashboard. |
| **Z-API cai** — leads ficam sem resposta | Heartbeat 10 min via `ana-cron`. Falha → Slack 🚨 + email Felipe. Manychat backup na Fase 1. |
| **Custos disparam** — Claude API explode | Limite mensal API (alerta R$ 400, corta R$ 600). Fallback: respostas cacheadas pras 50 perguntas mais comuns. |
| **Paciente descobre que é IA** e reclama | Se perguntar direto, ANA admite ("respondo com IA, mas qualquer momento equipe humana entra"). |

---

## 12. Fases de Implementação

### Fase 1 — Semanas 1-4 (ROI imediato)

| Semana | Foco | Entrega |
|---|---|---|
| 1 | Fundação | Conta Anthropic + Sheets ANA_* + FAQ estruturado + webhook Z-API. ANA "hello world" responde no WhatsApp. |
| 2 | Bot completo | `ana-bot` (webhook + Claude + memória + 5 red flags + Slack + pré-agendamento Feegow read). 20 conversas simuladas. |
| 3 | LP + Quiz | LP estática (GitHub Pages) + Quiz 5 perguntas + `ana-quiz` backend + integração Meta Ads. Live em `app.paraser.com.br/ana`. |
| 4 | Migração | Pilot 30% tráfego Meta → daily review → KPIs OK → 100% novo → **Manychat cancelado**. |

**Critério de sucesso Fase 1:** taxa lead → pré-agendamento ≥ 15% (igual Manychat) E zero incidente de red flag perdido.

### Fase 2 — Semanas 5-8 (Dashboard + Polish)

| Semana | Foco | Entrega |
|---|---|---|
| 5 | Dashboard v1 | `ana-dashboard` Web App + Tab Fila + KPIs + Confirmar/Recusar. Pessoa dedicada testa daily. |
| 6 | Dashboard v2 | Conversas ao vivo + Tomar conversa + Métricas Chart.js + FAQ editor. |
| 7 | Cron + Polish | `ana-cron` (timeouts, métricas, heartbeat). Tab Config ANA. Auditoria 100 conversas. Ajustes finos. |
| 8 | Handoff | Documentação + treinamento pessoa dedicada + runbook incidentes + plano de evolução (RAG?). |

**Critério de sucesso Fase 2:** pessoa dedicada opera sozinha, taxa lead → pré-agendamento ≥ 20%, NPS pós-consulta > baseline + 5pts.

---

## 13. Dependências Externas (Felipe fornece)

- **Antes Sem 1:** documento FAQ (rascunho aceito) + conta Anthropic criada + acesso Z-API verificado
- **Antes Sem 3:** definição do botão no Meta Ad (LP "Saiba mais" vs WhatsApp direto "Enviar Mensagem")
- **Antes Sem 4:** pessoa dedicada disponível pra daily review
- **Antes Sem 5:** política de privacidade publicada no site (LGPD)
- **Contínuo:** Felipe revisa amostras de conversas semanalmente nas Fases 1-2

---

## 14. Out of Scope (deixado pra projeto futuro)

- **RAG (vector DB)** — fica pra quando FAQ crescer > 20 páginas. Hoje cabe no prompt.
- **Memória longa multi-canal** — ANA só lembra conversas WhatsApp, não correlaciona com email Brevo. Quando precisar, vira projeto separado.
- **Bot proativo** — ANA não inicia conversa por conta própria, só responde quando paciente escreve.
- **Multi-idioma** — só PT-BR no MVP. Inglês/espanhol se algum dia atender estrangeiras.
- **Voz** — apenas texto. Áudio do paciente é transcrito mas resposta sempre em texto.
- **Integração CRM existente** — `ANA_Conversas` é Sheets isolada. Sincronização com CRM Comercial só depois de Fase 2.

---

## 15. Open Questions (a confirmar antes da Fase 1)

1. **FAQ formato** — Felipe vai entregar como? Google Doc, Sheets, Notion? (Resposta define se `ana-bot` lê direto do Sheets ou precisa importação.)
2. **Plantão noturno red flag** — quem recebe Slack 🚨 das 22h às 8h? Felipe direto? Atendente de plantão?
3. **Meta Ads** — Felipe vai criar campanhas com botão "Saiba mais" → LP? Ou começa só com migração das campanhas WhatsApp atuais?
4. **Acesso Feegow API** — `ana-bot` precisa de credencial separada ou usa a mesma do Confirmações Agenda?
5. **Domínio do Dashboard** — `app.paraser.com.br/ana/dashboard` ou subdomínio dedicado? (Preferência impacta CORS e config GitHub Pages.)

---

## 16. Connections (memórias relacionadas)

- `decisao_2026-05-19_brevo_reativacao_18k` — frente paralela de retenção
- `project_apps_scripts` — adiciona 4 novos scripts ao arsenal (ana-bot, ana-quiz, ana-dashboard, ana-cron)
- `reference_glossario_felipe` — termos como "atendente dedicada", "Lista A/B"
- `feedback_priorizar_automacao` — alinha com princípio de automação
- `feedback_materiais_equipe_jornada_nao_numeros` — tom acolhedor da ANA segue mesma linha dos emails Brevo
