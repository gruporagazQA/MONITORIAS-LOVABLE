#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 RAGAZ — EXPORTADOR DE TRANSCRIÇÕES PARA LOVABLE
================================================================================
 Gera um JSON estruturado com todas as transcrições casadas com o XLS,
 enriquecidas com gestor, área, supervisor, data, hora, telefone etc.

 Destinado à IA do Lovable para responder perguntas como:
   - "Mostre as ligações da Aline no dia 14/05"
   - "Quais ligações do Atendimento tiveram problema de saudação?"
   - "Resumo do Ariel na semana de 19/05"

 SAÍDA:
   RELATÓRIOS DD.MM.AA\
     transcricoes_maio_YYYYMMDD_HHMMSS.json    ← arquivo principal
     transcricoes_index_YYYYMMDD_HHMMSS.json   ← índice leve (sem texto)

 USO:
   python "G:\Meu Drive\ARQUITETURA LOVABLE\LINGUAGEM EM CÓDIGO\exportar_transcricoes.py"
================================================================================
"""

from __future__ import annotations

import json
import logging
import os
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

# ── dependências ──────────────────────────────────────────────────────────────
def _ensure(mod, pkg=""):
    try:
        __import__(mod)
    except ImportError:
        p = pkg or mod
        print(f"[setup] instalando {p}...")
        os.system(f'"{sys.executable}" -m pip install -q {p}')

for _m, _p in [("pandas","pandas"), ("openpyxl","openpyxl"), ("xlrd","xlrd==1.2.0")]:
    _ensure(_m, _p)

import pandas as pd

# ── configuração ──────────────────────────────────────────────────────────────
BASE_DIR     = Path(r"G:\Meu Drive\ARQUITETURA LOVABLE")
TRANSCR_DIR  = BASE_DIR / "TRANSCRIÇÕES"
AUDIOS_DIR   = BASE_DIR / "AUDIOS"

AREAS: Dict[str, Dict[str, Any]] = {
    "SUCESSO_CORRETOR": {
        "nome":       "Sucesso do Corretor",
        "supervisor": "Diego",
        "agentes":    ["Aline De Souza", "Ariel Maia", "Tatiana Costa",
                       "Taynnã Alves", "Ligiane Maia", "Eduardo Novais"],
    },
    "ATENDIMENTO": {
        "nome":       "Atendimento",
        "supervisor": "Danielle Moura",
        "agentes":    ["Deuzeni", "Lhaine", "Talita", "Jéssica"],
    },
    "COMERCIAL_CADASTRO": {
        "nome":       "Comercial Cadastro",
        "supervisor": "Rodrigo",
        "agentes":    ["Natália", "Ariely", "Ana Lúcia"],
    },
}

STATUS_OK    = "Bem sucedida"
DUR_MIN_SEG  = 30
TIPO_EF      = ("efetuad", "saída", "saida", "outbound")
TIPO_REC     = ("recebid", "entrada", "inbound")

PAT_TXT = re.compile(r"^(\d{4}-\d{2}-\d{2})_([\d-]+)_\+?(\d+)\.txt$")

# ── logging ───────────────────────────────────────────────────────────────────
LOG = logging.getLogger("exportar_transcricoes")
LOG.setLevel(logging.INFO)
_h = logging.StreamHandler(sys.stdout)
_h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S"))
LOG.addHandler(_h)

# ── utilitários ───────────────────────────────────────────────────────────────

def norm(s: str) -> str:
    nfkd = unicodedata.normalize("NFKD", str(s or ""))
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower().strip()

def u8(tel: str) -> str:
    d = re.sub(r"\D", "", tel)
    return d[-8:] if len(d) >= 8 else d

def normalizar_tel(t: Any) -> str:
    if t is None or str(t).strip() in ("", "-", "nan"):
        return ""
    d = re.sub(r"\D", "", str(t))
    if not d:
        return ""
    if not d.startswith("55") and len(d) >= 10:
        d = "55" + d[-10:]
    return "+" + d

def dur_seg(d: Any) -> int:
    if not d or str(d).strip() in ("", "-", "nan"):
        return 0
    s = str(d).lower()
    seg = 0
    m = re.search(r"(\d+)\s*m", s)
    if m: seg += int(m.group(1)) * 60
    m = re.search(r"(\d+)\s*s(?![a-z])", s)
    if m: seg += int(m.group(1))
    if seg == 0:
        raw = re.sub(r"\D", "", s)
        seg = int(raw) if raw else 0
    return seg

def parse_dt(s: Any) -> Optional[datetime]:
    for fmt in ("%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S",
                "%d/%m/%Y %H:%M",    "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(str(s).strip(), fmt)
        except Exception:
            continue
    return None

def col(df: pd.DataFrame, candidatas: List[str]) -> Optional[str]:
    cols_n = {norm(c): c for c in df.columns}
    for cand in candidatas:
        n = norm(cand)
        if n in cols_n:
            return cols_n[n]
    for cand in candidatas:
        n = norm(cand)
        for k, real in cols_n.items():
            if n in k:
                return real
    return None

def classificar_agente(nome: str) -> Optional[str]:
    alvo = norm(nome)
    for area_id, info in AREAS.items():
        for ag in info["agentes"]:
            if norm(ag) == alvo:
                return area_id
    for area_id, info in AREAS.items():
        for ag in info["agentes"]:
            ag_n = norm(ag)
            if len(ag_n) >= 4 and len(alvo) >= 4:
                if ag_n in alvo or alvo in ag_n:
                    return area_id
    return None

# ── leitura do XLS ────────────────────────────────────────────────────────────

def encontrar_xls() -> Optional[Path]:
    candidatos = [*BASE_DIR.glob("calls_detail_*.xlsx"),
                  *BASE_DIR.glob("calls_detail_*.xls"),
                  *BASE_DIR.glob("*.xlsx"),
                  *BASE_DIR.glob("*.xls")]
    candidatos = [c for c in candidatos if c.is_file()]
    return max(candidatos, key=lambda p: p.stat().st_mtime) if candidatos else None

def ler_ligacoes(xls: Path) -> List[Dict]:
    LOG.info(f"Lendo XLS: {xls.name}")
    try:
        df = pd.read_excel(xls, engine="openpyxl" if xls.suffix == ".xlsx" else "xlrd")
    except Exception:
        df = pd.read_html(str(xls))[0]

    c_col  = col(df, ["Colaborador"])
    c_tel  = col(df, ["Telefone"])
    c_data = col(df, ["Data da chamada", "Data", "Data/Hora"])
    c_dur  = col(df, ["Duração da chamada", "Duraçãoda chamada", "Duração"])
    c_tipo = col(df, ["Tipo Chamada", "Tipo de chamada", "Tipo"])
    c_stat = col(df, ["Status"])
    c_crm  = col(df, ["CRM", "Contato CRM"])

    ligacoes = []
    for _, row in df.iterrows():
        gestor  = str(row.get(c_col, "") or "").strip()
        tel     = normalizar_tel(row.get(c_tel, ""))
        dt      = parse_dt(row.get(c_data, ""))
        dur_s   = dur_seg(row.get(c_dur, "")) if c_dur else 0
        tipo_r  = norm(row.get(c_tipo, "") or "") if c_tipo else ""
        status  = str(row.get(c_stat, "") or "").strip()
        crm     = str(row.get(c_crm,  "") or "").strip() if c_crm else ""

        is_ef  = any(t in tipo_r for t in TIPO_EF)
        is_rec = any(t in tipo_r for t in TIPO_REC)
        is_ok  = status == STATUS_OK and dur_s > DUR_MIN_SEG
        area   = classificar_agente(gestor)
        if not area or not is_ok:
            continue

        ligacoes.append({
            "gestor":      gestor,
            "area_id":     area,
            "area_nome":   AREAS[area]["nome"],
            "supervisor":  AREAS[area]["supervisor"],
            "telefone":    tel,
            "data":        dt.strftime("%Y-%m-%d") if dt else "",
            "hora":        dt.strftime("%H:%M:%S") if dt else "",
            "datetime_iso":dt.isoformat() if dt else "",
            "duracao_seg": dur_s,
            "tipo":        "efetuada" if is_ef else ("recebida" if is_rec else "outro"),
            "status":      status,
            "crm":         crm,
        })

    LOG.info(f"Ligações bem-sucedidas em escopo: {len(ligacoes)}")
    return ligacoes

# ── indexar TXTs ──────────────────────────────────────────────────────────────

def indexar_txts() -> Dict[str, Path]:
    idx = {}
    if TRANSCR_DIR.exists():
        for arq in TRANSCR_DIR.glob("*.txt"):
            m = PAT_TXT.match(arq.name)
            if m:
                chave = f"{m.group(1)}|{u8(m.group(3))}"
                idx[chave] = arq
    LOG.info(f"TXTs indexados: {len(idx)}")
    return idx

# ── casar e montar registros ──────────────────────────────────────────────────

def montar_registros(ligacoes: List[Dict], idx_txt: Dict[str, Path]) -> List[Dict]:
    registros = []
    sem_txt   = 0

    for i, lig in enumerate(ligacoes):
        chave = f"{lig['data']}|{u8(lig['telefone'])}"
        arq   = idx_txt.get(chave)
        if not arq:
            sem_txt += 1
            continue

        texto = arq.read_text(encoding="utf-8", errors="replace").strip()
        if len(texto) < 30:
            sem_txt += 1
            continue

        registros.append({
            # ── identificação ──────────────────────────────────────────
            "id":           f"{lig['data']}_{lig['hora'].replace(':','')}_{u8(lig['telefone'])}",
            # ── quem ───────────────────────────────────────────────────
            "gestor":       lig["gestor"],
            "area_id":      lig["area_id"],
            "area_nome":    lig["area_nome"],
            "supervisor":   lig["supervisor"],
            # ── quando ─────────────────────────────────────────────────
            "data":         lig["data"],
            "hora":         lig["hora"],
            "datetime_iso": lig["datetime_iso"],
            "dia_semana":   _dia_semana(lig["data"]),
            # ── como ───────────────────────────────────────────────────
            "telefone":     lig["telefone"],
            "duracao_seg":  lig["duracao_seg"],
            "duracao_fmt":  _fmt_dur(lig["duracao_seg"]),
            "tipo":         lig["tipo"],
            "crm":          lig["crm"],
            # ── conteúdo ───────────────────────────────────────────────
            "transcricao":  texto,
            "palavras":     len(texto.split()),
        })

    LOG.info(f"Registros com transcrição: {len(registros)} | Sem TXT: {sem_txt}")
    return registros

def _dia_semana(data_str: str) -> str:
    DIAS = ["Segunda","Terça","Quarta","Quinta","Sexta","Sábado","Domingo"]
    try:
        dt = datetime.strptime(data_str, "%Y-%m-%d")
        return DIAS[dt.weekday()]
    except Exception:
        return ""

def _fmt_dur(seg: int) -> str:
    m, s = divmod(seg, 60)
    return f"{m}min {s}s"

# ── montar JSON principal ─────────────────────────────────────────────────────

def montar_json(registros: List[Dict]) -> Dict:
    # índices para a IA navegar rapidamente
    por_gestor: Dict[str, List[str]] = {}
    por_area:   Dict[str, List[str]] = {}
    por_data:   Dict[str, List[str]] = {}

    for r in registros:
        rid = r["id"]
        por_gestor.setdefault(r["gestor"],   []).append(rid)
        por_area.setdefault(r["area_nome"],  []).append(rid)
        por_data.setdefault(r["data"],       []).append(rid)

    return {
        "meta": {
            "gerado_em":        datetime.now().isoformat(),
            "mes_referencia":   "2026-05",
            "total_registros":  len(registros),
            "areas":            [info["nome"] for info in AREAS.values()],
            "instrucao_ia": (
                "Este JSON contém transcrições de ligações da equipe Ragaz em maio/2026. "
                "Use 'por_gestor', 'por_area' ou 'por_data' para localizar IDs rapidamente, "
                "depois consulte 'transcricoes' pelo campo 'id' para obter o texto completo."
            ),
        },
        "indices": {
            "por_gestor": por_gestor,
            "por_area":   por_area,
            "por_data":   por_data,
        },
        "transcricoes": {r["id"]: r for r in registros},
    }

def montar_index(registros: List[Dict]) -> List[Dict]:
    """Versão leve sem o texto da transcrição — para carregamento rápido."""
    return [
        {k: v for k, v in r.items() if k != "transcricao"}
        for r in registros
    ]

# ── main ──────────────────────────────────────────────────────────────────────

def main() -> int:
    ts         = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = BASE_DIR / f"RELATÓRIOS {datetime.now():%d.%m.%y}"
    output_dir.mkdir(parents=True, exist_ok=True)

    xls = encontrar_xls()
    if not xls:
        LOG.error(f"Nenhum XLS encontrado em {BASE_DIR}")
        return 2

    ligacoes = ler_ligacoes(xls)
    idx_txt  = indexar_txts()
    registros = montar_registros(ligacoes, idx_txt)

    if not registros:
        LOG.error("Nenhuma transcrição encontrada após casamento. Verifique TRANSCRIÇÕES\ e o XLS.")
        return 1

    # ── arquivo principal ──────────────────────────────────────────────────
    json_completo = montar_json(registros)
    path_completo = output_dir / f"transcricoes_maio_{ts}.json"
    with path_completo.open("w", encoding="utf-8") as f:
        json.dump(json_completo, f, ensure_ascii=False, indent=2, default=str)

    # ── índice leve (sem texto) ────────────────────────────────────────────
    json_index = montar_index(registros)
    path_index = output_dir / f"transcricoes_index_{ts}.json"
    with path_index.open("w", encoding="utf-8") as f:
        json.dump(json_index, f, ensure_ascii=False, indent=2, default=str)

    tam_mb = path_completo.stat().st_size / 1_048_576

    print("\n" + "=" * 60)
    print("  EXPORTAÇÃO DE TRANSCRIÇÕES — RAGAZ")
    print("=" * 60)
    print(f"  Total de registros  : {len(registros)}")
    print(f"  Arquivo completo    : {path_completo.name}")
    print(f"  Tamanho             : {tam_mb:.2f} MB")
    print(f"  Índice leve         : {path_index.name}")
    print()
    print("  POR ÁREA:")
    for area_id, info in AREAS.items():
        n = sum(1 for r in registros if r["area_id"] == area_id)
        print(f"   - {info['nome']:<26} {n:>4} transcrições")
    print()
    print("  POR DATA:")
    from collections import Counter
    datas = Counter(r["data"] for r in registros)
    for data in sorted(datas):
        print(f"   - {data}  {datas[data]:>4} transcrições")
    print("=" * 60)
    return 0

if __name__ == "__main__":
    sys.exit(main())
