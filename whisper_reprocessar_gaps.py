#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 RAGAZ — WHISPER REPROCESSAR GAPS
================================================================================
 Reprocessa com Whisper os áudios que o Google Speech não conseguiu transcrever.
 Pega as ligações da aba "Sem Transcrição" do relatório V7,
 localiza os MP3s correspondentes e tenta retranscrever usando Whisper local.

 Os TXTs gerados caem na pasta TRANSCRICOES normal — na próxima rodada do V7
 (ou via pipeline V9 Fase 6) eles serão aproveitados automaticamente.

 USO (PowerShell):
   python "G:\\Meu Drive\\ARQUITETURA LOVABLE\\LINGUAGEM EM CODIGO\\whisper_reprocessar_gaps.py"

 OPÇÕES:
   --relatorio   Caminho do xlsx V7 (padrão: mais recente em RELATÓRIOS*)
   --audios      Pasta de MP3s (padrão: AUDIOS\JUNHO_2026)
   --modelo      Modelo Whisper: tiny | base | small | medium (padrão: small)
   --max         Limita N áudios (para testes)
   --so-inint    Só reprocessa os "Áudio ininteligível", pula os sem áudio
================================================================================
"""
from __future__ import annotations
import argparse, os, re, sys, tempfile, time, unicodedata
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

def _ensure(mod, pkg=""):
    try: __import__(mod)
    except ImportError:
        p = pkg or mod
        print(f"[setup] instalando {p}...")
        os.system(f'"{sys.executable}" -m pip install -q {p}')

_ensure("tqdm")
_ensure("librosa")
_ensure("soundfile")
_ensure("scipy")

# Whisper instalado separado pois é grande
try:
    import whisper as _w
except ImportError:
    print("[setup] instalando openai-whisper (pode demorar ~1 min)...")
    os.system(f'"{sys.executable}" -m pip install -q openai-whisper')
    import whisper as _w

import pandas as pd
from tqdm import tqdm

# ── configuração ──────────────────────────────────────────────────────────────
BASE_DIR    = Path(r"G:\Meu Drive\ARQUITETURA LOVABLE")
TRANSCR_DIR = BASE_DIR / "TRANSCRIÇÕES"
AUDIOS_DIR  = BASE_DIR / "AUDIOS" / "JUNHO_2026"

PAT_MP3 = re.compile(r"^(\d{4}-\d{2}-\d{2}) [\d-]+ \+?(\d+)\.mp3$")

def norm(s):
    nfkd = unicodedata.normalize("NFKD", str(s or ""))
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower().strip()

def u8(tel):
    d = re.sub(r"\D", "", str(tel))
    return d[-8:] if len(d) >= 8 else d

# ── localizar relatório mais recente ─────────────────────────────────────────

def encontrar_relatorio() -> Optional[Path]:
    candidatos = []
    for pasta in BASE_DIR.glob("RELATÓRIOS*"):
        candidatos += list(pasta.glob("RELATORIO_*.xlsx"))
    return max(candidatos, key=lambda p: p.stat().st_mtime) if candidatos else None

# ── indexar MP3s ──────────────────────────────────────────────────────────────

def indexar_mp3s(audios_dir: Path) -> Dict[str, Path]:
    idx = {}
    for arq in audios_dir.glob("*.mp3"):
        m = PAT_MP3.match(arq.name)
        if m:
            chave = f"{m.group(1)}|{u8(m.group(2))}"
            idx[chave] = arq
    return idx

# ── transcrever com Whisper ───────────────────────────────────────────────────

def transcrever_whisper(mp3: Path, modelo) -> Tuple[Optional[str], Optional[str]]:
    try:
        result = modelo.transcribe(str(mp3), language="pt", fp16=False)
        texto = result.get("text", "").strip()
        if not texto or len(texto) < 30:
            return None, "whisper_texto_vazio"
        return texto, None
    except Exception as e:
        return None, f"whisper_erro: {e}"

# ── salvar TXT no padrão do V6 ───────────────────────────────────────────────

def salvar_txt(mp3: Path, texto: str) -> Path:
    # Formato: 2026-06-12_171255_+5511985385800.txt
    nome_txt = mp3.stem.replace(" ", "_") + ".txt"
    destino = TRANSCR_DIR / nome_txt
    destino.write_text(texto, encoding="utf-8")
    return destino

# ── main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--relatorio", default="")
    parser.add_argument("--audios",    default=str(AUDIOS_DIR))
    parser.add_argument("--modelo",    default="small",
                        choices=["tiny","base","small","medium"])
    parser.add_argument("--max",       type=int, default=0)
    parser.add_argument("--so-inint",  action="store_true",
                        help="Só reprocessa áudios ininteligíveis (pula os sem áudio)")
    args = parser.parse_args()

    audios_dir = Path(args.audios)
    TRANSCR_DIR.mkdir(parents=True, exist_ok=True)

    # 1) Localizar relatório
    rel_path = Path(args.relatorio) if args.relatorio else encontrar_relatorio()
    if not rel_path or not rel_path.exists():
        print("ERRO: Relatório V6 não encontrado. Use --relatorio <caminho>")
        return 2
    print(f"Relatório: {rel_path.name}")

    # 2) Ler aba Sem Transcrição
    df = pd.read_excel(rel_path, sheet_name="Sem Transcrição")
    print(f"Gaps no relatório: {len(df)}")

    if args.so_inint:
        df = df[df["Motivo"].str.contains("ininteligível|erro_google", case=False, na=False)]
        print(f"Filtrado para ininteligíveis/erro API: {len(df)}")

    # Excluir: sem áudio (nada a fazer) e TXTs filtrados pelo --novos-desde (já têm transcrição)
    df = df[~df["Motivo"].str.contains(
        "não baixado|txt_anterior_ao_whisper", case=False, na=False
    )]
    print(f"Com áudio potencial para reprocessar: {len(df)}")

    if args.max and args.max > 0:
        df = df.head(args.max)
        print(f"Limitado a {args.max} para teste")

    # 3) Indexar MP3s
    idx_mp3 = indexar_mp3s(audios_dir)
    print(f"MP3s indexados em {audios_dir.name}: {len(idx_mp3)}")

    # 4) Casar gaps com MP3s
    pendentes = []
    for _, row in df.iterrows():
        data = str(row["Data"])[:10]
        tel  = str(row["Telefone"])
        chave = f"{data}|{u8(tel)}"
        mp3 = idx_mp3.get(chave)
        if mp3:
            pendentes.append({"data": data, "tel": tel, "mp3": mp3, "motivo": row["Motivo"]})

    print(f"MP3s localizados para reprocessar: {len(pendentes)}")
    nao_localizados = len(df) - len(pendentes)
    if nao_localizados:
        print(f"Não localizados (sem MP3 na pasta): {nao_localizados}")

    if not pendentes:
        print("Nenhum MP3 para reprocessar. Encerrando.")
        return 0

    # 5) Carregar modelo Whisper
    print(f"\nCarregando modelo Whisper '{args.modelo}'... (primeira vez faz download)")
    modelo = _w.load_model(args.modelo)
    print("Modelo carregado.\n")

    # 6) Transcrever
    ok = 0; falhou = 0; pulou = 0
    log_linhas = []
    ts_inicio = datetime.now()

    for item in tqdm(pendentes, desc="Whisper", unit="áudio"):
        mp3: Path = item["mp3"]

        # Pular se já existe TXT (rodada anterior gerou)
        nome_txt = mp3.stem.replace(" ", "_") + ".txt"
        if (TRANSCR_DIR / nome_txt).exists():
            pulou += 1
            continue

        texto, erro = transcrever_whisper(mp3, modelo)

        if erro or not texto:
            falhou += 1
            log_linhas.append(f"FALHOU | {mp3.name} | {erro}")
            continue

        destino = salvar_txt(mp3, texto)
        ok += 1
        log_linhas.append(f"OK     | {mp3.name} | {len(texto.split())} palavras → {destino.name}")
        time.sleep(0.2)

    # 7) Resumo
    duracao = (datetime.now() - ts_inicio).seconds // 60
    print("\n" + "=" * 60)
    print("  WHISPER — REPROCESSAMENTO DE GAPS")
    print("=" * 60)
    print(f"  Reprocessados com sucesso : {ok}")
    print(f"  Falhou                    : {falhou}")
    print(f"  Já tinha TXT (pulou)      : {pulou}")
    print(f"  Sem MP3 na pasta          : {nao_localizados}")
    print(f"  Tempo total               : ~{duracao} min")
    print()
    print(f"  TXTs salvos em: {TRANSCR_DIR}")
    print("  → Rode o V6 com --so-txts para aproveitar na próxima análise")
    print("=" * 60)

    # Salvar log
    log_path = BASE_DIR / f"whisper_gaps_{datetime.now():%Y%m%d_%H%M%S}.log"
    log_path.write_text("\n".join(log_linhas), encoding="utf-8")
    print(f"\n  Log salvo: {log_path.name}")

    return 0

if __name__ == "__main__":
    sys.exit(main())
