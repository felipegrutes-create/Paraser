"""
Limpa o CSV de leads do Meta Lead Center e gera um CSV pronto pra Brevo.

Input:  leads-meta.csv (encoding latin-1, 19k linhas)
Output: leads-meta-brevo.csv (UTF-8, deduplicado, validado)

Critérios de limpeza:
  - Email obrigatório e em formato válido (regex)
  - Deduplicar por email (mantém o lead MAIS ANTIGO — first touch)
  - Normalizar telefone BR: só dígitos, prefixo 55, formato E.164 sem +
  - Split full_name em FIRSTNAME + LASTNAME
  - Mapear campos extras pra atributos Brevo
  - Excluir leads SEM email (Brevo exige email pra cadastrar)

Atributos custom na Brevo:
  - TEMPO_TENTANDO (qualificador 1)
  - JA_CONSULTOU (qualificador 2)
  - PLATAFORMA (ig/fb)
  - CAMPANHA (campaign_name)
  - CRIATIVO (ad_name)
  - DATA_CADASTRO (created_time)
"""
import csv
import re
from collections import Counter
from datetime import datetime

INPUT_CSV  = r"c:/Users/USER/Paraser/leads-meta.csv"
OUTPUT_CSV = r"c:/Users/USER/Paraser/leads-meta-brevo.csv"

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def norm_email(s):
    if not s: return None
    e = s.strip().lower()
    if not EMAIL_RE.match(e): return None
    # Excluir emails óbvios de teste/lixo
    if any(x in e for x in ['teste@', 'test@', 'noemail', 'naoinformado', 'email@email']):
        return None
    return e

def norm_phone(s):
    if not s: return ""
    p = re.sub(r'\D', '', s)
    if not p: return ""
    # Adiciona 55 se vier só DDD+número
    if len(p) == 10 or len(p) == 11:
        p = '55' + p
    # Tira "0" extra (021 → 21)
    if p.startswith('550') and len(p) >= 13:
        p = '55' + p[3:]
    if len(p) < 12 or len(p) > 13:
        return ""
    return p

def split_name(full):
    parts = (full or '').strip().split()
    if not parts: return ('', '')
    if len(parts) == 1: return (parts[0].capitalize(), '')
    first = parts[0].capitalize()
    last  = ' '.join(p.capitalize() for p in parts[1:])
    return (first, last)

def fmt_date(s):
    """2025-09-10T21:46:22-05:00 → 2025-09-10"""
    if not s: return ""
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', s)
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)}" if m else ""

# === LER ===
with open(INPUT_CSV, encoding='utf-8', newline='') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

print(f"Total bruto: {len(rows)}")

# === LIMPAR / DEDUPLICAR ===
seen = {}              # email -> row mais antigo
sem_email = 0
invalido = 0
duplicados = 0
tempo_count = Counter()
consulta_count = Counter()
plat_count = Counter()
camp_count = Counter()

# Mapeia nomes de coluna problemáticos (com acentos)
def get_col(row, *candidates):
    for c in candidates:
        if c in row:
            return row[c]
    # fallback: busca por substring
    for k, v in row.items():
        for c in candidates:
            if c.lower().replace('?','').replace('_','') in k.lower().replace('?','').replace('_',''):
                return v
    return ""

for r in rows:
    email = norm_email(r.get('email', ''))
    if not r.get('email'):
        sem_email += 1
        continue
    if not email:
        invalido += 1
        continue

    created = r.get('created_time', '')
    if email in seen:
        # mantém o mais antigo
        if created < seen[email]['created_time']:
            seen[email] = r
        duplicados += 1
        continue
    seen[email] = r

print(f"Sem email:          {sem_email}")
print(f"Email inválido:     {invalido}")
print(f"Duplicados:         {duplicados}")
print(f"ÚNICOS p/ Brevo:    {len(seen)}")

# === ESTATÍSTICAS ===
for r in seen.values():
    tempo = get_col(r, 'há_quanto_tempo_você_tenta_engravidar?', 'h�_quanto_tempo_voc�_tenta_engravidar?')
    consulta = get_col(r, 'você_já_consultou_um_especialista_em_fertilidade?', 'voc�_j�_consultou_um_especialista_em_fertilidade?')
    tempo_count[tempo or '(vazio)'] += 1
    consulta_count[consulta or '(vazio)'] += 1
    plat_count[r.get('platform', '(vazio)')] += 1
    # Pega só o mês/ano da campanha (campaign_name termina em "YYYY-MM" geralmente)
    camp = r.get('campaign_name', '')
    m = re.search(r'(\d{4}-\d{2})', camp)
    camp_count[m.group(1) if m else 'outro'] += 1

print("\n=== Tempo tentando engravidar ===")
for k, v in tempo_count.most_common():
    pct = v/len(seen)*100
    print(f"  {k:30s} {v:6d}  ({pct:.1f}%)")

print("\n=== Já consultou especialista ===")
for k, v in consulta_count.most_common():
    pct = v/len(seen)*100
    print(f"  {k:10s} {v:6d}  ({pct:.1f}%)")

print("\n=== Plataforma ===")
for k, v in plat_count.most_common():
    print(f"  {k:5s} {v:6d}")

print("\n=== Distribuição por mês de campanha ===")
for k in sorted(camp_count.keys()):
    print(f"  {k:10s} {camp_count[k]:6d}")

# === GERAR CSV BREVO ===
COLS_BREVO = [
    'EMAIL', 'FIRSTNAME', 'LASTNAME', 'SMS',
    'TEMPO_TENTANDO', 'JA_CONSULTOU',
    'PLATAFORMA', 'CAMPANHA', 'CRIATIVO',
    'DATA_CADASTRO', 'LEAD_ID_META'
]

with open(OUTPUT_CSV, 'w', encoding='utf-8', newline='') as f:
    w = csv.DictWriter(f, fieldnames=COLS_BREVO)
    w.writeheader()
    for email, r in seen.items():
        first, last = split_name(r.get('full_name', ''))
        phone = norm_phone(r.get('whatsapp_number', ''))
        w.writerow({
            'EMAIL': email,
            'FIRSTNAME': first,
            'LASTNAME': last,
            'SMS': phone,
            'TEMPO_TENTANDO': get_col(r, 'há_quanto_tempo_você_tenta_engravidar?', 'h�_quanto_tempo_voc�_tenta_engravidar?'),
            'JA_CONSULTOU':   get_col(r, 'você_já_consultou_um_especialista_em_fertilidade?', 'voc�_j�_consultou_um_especialista_em_fertilidade?'),
            'PLATAFORMA':     r.get('platform', ''),
            'CAMPANHA':       r.get('campaign_name', ''),
            'CRIATIVO':       r.get('ad_name', ''),
            'DATA_CADASTRO':  fmt_date(r.get('created_time', '')),
            'LEAD_ID_META':   r.get('id', '').replace('l:', '')
        })

print(f"\n[OK] CSV pronto pra Brevo: {OUTPUT_CSV}")
print(f"     {len(seen)} leads unicos | 11 colunas | UTF-8")
