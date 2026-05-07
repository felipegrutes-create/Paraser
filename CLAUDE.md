# Paraser — Clínica de Fertilidade RJ

> Sistema interno: dashboard financeiro/operacional + CRM comercial + automação WhatsApp.

## 👤 Sobre o usuário

Felipe Grutes (felipegrutes@yahoo.com.br). **Não é programador** — prefere instruções simples, comandos prontos pra colar, sem jargão técnico. Idioma: PT-BR.

## 🏥 Sobre a clínica

**Paraser** — clínica de fertilidade no Rio de Janeiro.
**Endereço:** Rua Prof. Álvaro Rodrigues, 352, 10º andar, Botafogo (próximo ao metrô Botafogo, saída E).

## 🌐 Onde tudo roda

| Camada | Localização |
|---|---|
| 📁 Pasta local desk | `C:\Users\USER\Paraser\` |
| 📁 Pasta local notebook | `C:\Users\felip\Paraser\` |
| 📦 Repo GitHub | `https://github.com/felipegrutes-create/Paraser` |
| 🌐 Dashboard live | `https://felipegrutes-create.github.io/Paraser/dashboard-feegow-v2.html` (GitHub Pages auto-deploy) |
| 📜 Apps Scripts (cloud) | "Confirmação de Agenda" + "CRM Paraser" no `script.google.com` |

## 📄 Arquivos principais

- `dashboard-feegow-v2.html` — **dashboard único** (single HTML file). Roda no GitHub Pages. Contém Chart.js + integração Google Sheets API + multiplas abas (Receitas, Médicos, Projeto ANA, etc).
- `crm-apps-script.gs` — código fonte do Apps Script "CRM Paraser" (cópia local; **roda no Google Cloud**).
- `confirmacoes-whatsapp.gs` — código fonte do Apps Script "Confirmação de Agenda" (envia WhatsApp via Z-API + notifica Slack).

## 📊 Integrações de dados

| Sistema | Pra que serve |
|---|---|
| **Feegow API** (`https://api.feegow.com/v1/api/`) | Buscar agendamentos, profissionais, procedimentos |
| **Google Sheets** (ID `1uthRnuWMk2A26dZ8GaMXinuvPyxY50EH8RX85NemXwg`) | Planilha financeira (LANÇAMENTOS + LANÇAMENTOS_I + PAGAMENTOS) |
| **Z-API** | Envia mensagens WhatsApp pros pacientes |
| **Slack** (canal `atendimento`) | Notifica equipe sobre confirmações enviadas |

## 🚫 REGRAS CRÍTICAS — NÃO FAZER

1. **NÃO commitar tokens** (FEEGOW_TOKEN, ZAPI_TOKEN, SLACK_TOKEN, etc) — sempre via PropertiesService no Apps Script
2. **NÃO esquecer de copiar `.gs` editado pro Apps Script Cloud** — o arquivo local é só cópia, não roda sozinho
3. **NÃO mexer em** `Quiz Nutriplan/`, `images/eyes`, `images/skin*`, `HTipo*.png`, `Tipo*.png`, `Olho*.png`, `extract_bonos.py`, `extract_recipes.py` (são fósseis do quiz antigo, não fazem parte do Paraser ativo)
4. **SEMPRE git pull antes de mexer** — Felipe edita de outras IAs/máquinas com frequência (vimos em 2026-05-06: 52 commits novos no remote)

## 🔧 Stack técnica

- **Frontend:** HTML + Chart.js (single file, sem build)
- **Backend:** Google Apps Script (sem servidor próprio)
- **Banco:** Google Sheets (planilhas) + Feegow API
- **Mensageria:** Z-API (WhatsApp) + Slack
- **Deploy:** GitHub Pages (auto-deploy do branch `main`)

## 🧠 Memórias completas

`~/.claude/projects/<hash>/memory/MEMORY.md` (lido automaticamente). Memórias incluem:
- Perfil do usuário
- Estrutura do sistema
- Apps Scripts ativos e como editar
- URLs e localizações

Memórias sincronizadas via Google Drive (`G:\Meu Drive\claude-memory-felipe\paraser\`) entre desk + notebook + Mac (futuro). Junction Point conecta Claude Code ao Drive transparentemente.

## 🤝 Como colaborar comigo (Claude)

1. **Início de sessão:** leio MEMORY.md + este CLAUDE.md automaticamente
2. **Antes de mexer:** git pull pra garantir que estou no estado mais recente do remote
3. **Ao editar `.gs`:** lembrar Felipe de copiar pro Apps Script Cloud (cópia local não roda)
4. **Ao tomar decisão técnica:** registrar como `decisao_YYYY-MM-DD_<assunto>.md` em memory/

---

*Última revisão: 2026-05-06.*
