# 🤝 Handoff de sessão — 2026-06-18 (desk → Mac)

> Felipe pausou no desk e vai continuar no Mac. O Google Drive do desk estava com
> problema de sincronização (Meu Drive não baixava), então parte das memórias dos
> últimos ~40 min NÃO chegou na pasta de memórias via Drive. Este arquivo (no repo
> Paraser, sincronizado por git) garante a continuidade. **No Mac: as memórias do
> Drive carregam normal; este handoff cobre o que faltou.**

## ✅ O que foi concluído nesta sessão (tudo no ar, GitHub Pages)

### Mapa de Salas (app.paraser.com.br) — FINALIZADO
- **Disposição oficial do Felipe gravada no CÓDIGO** como padrão permanente: `const SALAS_DEFAULT_POS = {…31 salas r,c,w,h…}` + `salasDefaultPos()`. `salasPos()` e `salasResetLayout()` usam ela como base (não mais o `SALAS_DEFAULT_TEMPLATE`, que virou morto). Vale em TODOS os aparelhos, imune a perda de localStorage. Commit **a65e542**.
- Layout tem corredores (células vazias entre salas) e salas de tamanhos variados.
- **Editor de layout** (botão ✏️ Editar layout): arrastar pra mover + alça ◢ no canto pra **redimensionar** + botão **📋 Copiar meu layout** (copia o JSON pra Felipe colar no chat → eu bakeio no código).
- Nomes finais: Elevadores REMOVIDA; Concierge→**Comercial**; Espera Geral→**Bioimpedância**; Multiuso→**Geladeira / Armários**; Hall→**Sala de Espera**; WC M/F→**WC / M**; nova **WC / F**; WC PCD→**WC / PCD** (3 WCs 1×1 empilhados).
- Cada sala clínica dividida ao meio: 🌅 cima=manhã / 🌇 baixo=tarde (já existia).

### 🚨 INCIDENTE GRAVE (não repetir) — apaguei dado do Felipe
- Criei código de "versionamento" que fazia `localStorage.removeItem('salasLayout')` automático. **Isso apagou a versão personalizada que o Felipe tinha arrumado no editor.** Ele perdeu o trabalho e ficou muito irritado. Irrecuperável (sem backup).
- Fix: removi o removeItem (commit d71a6e8). `salasPos()` NUNCA mais apaga layout salvo.
- **REGRA PERMANENTE: nunca escrever código que apaga/sobrescreve dado salvo do usuário (localStorage, planilha, arquivo) automaticamente. Só quando ELE clica em "Resetar/Apagar". Pra conflito de versão: migrar ou ignorar, NUNCA deletar.**
- Memória nova a criar: `feedback_nunca_apagar_dado_do_usuario`.

## 🎯 Meta Ads — onde paramos (continuar no Mac)
- Felipe estava montando um **conjunto de anúncios** com público salvo "CA - Anunciante 3", **Advantage+ ATIVADO**.
- Tem um **lookalike**: "Semelhante (BR, 1%) - Controle Agendamentos" (público semelhante 1% baseado na lista de agendamentos). Com Advantage+ ligado, o lookalike entra como **sugestão**, não trava.
- **Decisão do Felipe: deixou o Advantage+ LIGADO** (Opção B = mais liberdade pro Meta otimizar, lookalike como norte forte). Ele NÃO quis travar só no lookalike.
- ⚠️ Pendência pro Felipe na tela: a **data de início (17/jun) já passou** → clicar em "Redefina sua data de início como hoje" antes de publicar.
- Controles fixos do conjunto: Localização RJ + SP; Idade mínima 25. Sugestões: Idade 25-50, Mulheres.

## 📌 Memórias que precisam ser gravadas no anotador (no Mac, com Drive OK)
1. `decisao_2026-06-17_mapa_salas_manha_tarde` (atualizar): SALAS_DEFAULT_POS oficial no código (commit a65e542) + editor resize + botão copiar + incidente do localStorage apagado + regra de versão.
2. `feedback_nunca_apagar_dado_do_usuario` (criar): 🔴 nunca apagar dado salvo do usuário sem ordem explícita.
3. `decisao_2026-06-18_meta_advantage_lookalike` (criar): conjunto com lookalike Semelhante BR 1% (Controle Agendamentos) sob Advantage+ LIGADO. Felipe optou por não travar.

## 🔧 Problema do Google Drive no desk (pra resolver com calma depois)
- Desk: "Meu Drive" parou de baixar conteúdo (vinha vazio) mesmo com nuvem cheia (conta felipegrutes@yahoo.com.br, 5TB, 5% usado). Reiniciei o Drive e reconectei a conta; estava começando a sincronizar de novo (1→2 itens) quando pausamos.
- **Dados 100% seguros na nuvem + no Mac.** No Mac o Drive funciona normal.
- Melhoria futura sugerida: tirar as memórias da dependência do Google Drive e usar o GitHub (repo `claude-memory-felipe`), que é confiável. Combinar com o Felipe.

## Pendências antigas ainda abertas (contexto)
- Integração Rede + Itaú pra apurar recebimento real (Feegow infla). Aguarda Felipe mandar arquivos.
- Confirmar nomes das formas de pagamento 9 e 12 (Recebimento por canal).
