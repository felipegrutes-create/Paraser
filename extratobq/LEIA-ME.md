# Script da planilha FLUXO DE CX 2026

Espelho versionado do Apps Script vinculado à planilha (Script ID
`1ql4PYVnZfJ4s3P0XcejLbQHZt4yS29fGa5D2hWxpmhH4cp1RAgtoQ5hT`, nome "DATA ATUAL").

**O que roda é a cópia na nuvem, não esta.** Clone local com clasp em
`C:\Users\USER\apps-script-fluxocx` (usar `-u paraser`). Fluxo: editar lá →
`clasp -u paraser push --force` → copiar para cá → commit.

É um script **bound** (vinculado à planilha): `clasp run` não funciona nele, e o
web app é restrito ao domínio. Para executar uma função sem abrir o editor, o
caminho usado foi um endpoint temporário no script do CRM (que tem acesso à
mesma planilha por `openById`) com deployment descartável, removido depois.

## ⚠️ O fuso da planilha é America/Los_Angeles

Não é São Paulo. Toda data escrita como `Date` cai com 07:00 ou 08:00 (muda com
o horário de verão) e deixa de ser um dia puro, o que faz fórmula de outra aba
não encontrar a linha. Por isso existe `gravarDataPura_()` em `Extratos.js`:
grava o número de série do dia, sem fuso no meio. **Usar essa função em todo
lugar que escreve data na coluna A.**
