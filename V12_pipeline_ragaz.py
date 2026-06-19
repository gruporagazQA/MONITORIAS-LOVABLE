#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 RAGAZ — PIPELINE UNIFICADO V12
================================================================================
 Orquestra o ciclo completo de monitorias em 7 fases:

   1. Download XLS de ligações do Bitrix24 (Playwright → headless)
   2. Download áudios MP3 do Bitrix24 Drive (REST API)
   3. Mover áudios para pasta do V7  (AUDIOS/MES_ANO)
   4. Análise V7  (transcrição Google Speech + Claude Haiku)
        → gera relatório com aba "Sem Transcrição"
   5. Whisper nos gaps do relatório da Fase 4
        → reprocessa áudios que o Google não conseguiu transcrever
   6. Análise V7 --so-txts --novos-desde TIMESTAMP
        → Claude analisa APENAS os TXTs criados pelo Whisper (Fase 5)
        → NÃO re-analisa as ligações já analisadas na Fase 4
   7. Exportar JSON de transcrições para Lovable AI

 Cobertura: sempre d-1 do mês de referência
   (ex: rodando em 18/06 cobre 01/06 → 17/06)

 USO (PowerShell):
   # Ciclo completo (mês corrente)
   python "...\\V12_pipeline_ragaz.py"

   # Mês específico
   python "...\\V12_pipeline_ragaz.py" --mes 2026-07

   # Pular fases 1-3 (downloads já feitos)
   python "...\\V12_pipeline_ragaz.py" --fases 4,5,6,7

   # Pular Whisper (só Google + Claude)
   python "...\\V12_pipeline_ragaz.py" --fases 1,2,3,4,7

   # Modelo Whisper mais rápido (tiny/base/small/medium, padrão: small)
   python "...\\V12_pipeline_ragaz.py" --modelo-whisper small

   # Primeiro login manual Bitrix (necessário apenas uma vez)
   python "...\\V12_pipeline_ragaz.py" --setup-sessao

   # Registrar no Agendador de Tarefas do Windows (22:00 diário)
   python "...\\V12_pipeline_ragaz.py" --agendar --hora 22:00

 PRÉ-REQUISITOS:
   .env em AUTOMAÇÃO EXPORTAÇÃO com BITRIX_WEBHOOK
   $env:ANTHROPIC_API_KEY = "sk-ant-..."

 TEMPO ESTIMADO (referência junho/2026 — ~800 ligações):
   Fase 1-3  ~40 min   (download + move)
   Fase 4    ~60 min   (Google transcription + Claude)
   Fase 5    ~2h       (Whisper small em ~50 gaps)
   Fase 6    ~2 min    (Claude só nos gaps novos do Whisper — não re-analisa tudo)
   Fase 7    <1 min
   TOTAL     ~4h       → agendar às 22:00 para ter o relatório pela manhã

 NOVIDADES V12 vs V11:
   - Fase 5 recebe o timestamp de início da Fase 4 e usa-o para filtrar
     o relatório de referência: garante que o Whisper sempre leia o Excel
     gerado pela Fase 4 desta rodada, nunca um relatório dry-run anterior
   - whisper_reprocessar_gaps.py: exclui "txt_anterior_ao_whisper" da lista
     de gaps, eliminando o falso positivo de 789 gaps (era 30 reais)
================================================================================
"""

from __future__ import annotations

import argparse
import datetime
import logging
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

# ── configuração ──────────────────────────────────────────────────────────────
BASE_DIR       = Path(r"G:\Meu Drive\ARQUITETURA LOVABLE")
SCRIPTS_DIR    = BASE_DIR / "LINGUAGEM EM CÓDIGO"
AUDIOS_BASE    = BASE_DIR / "AUDIOS"
AUTOMACAO_DIR  = Path(r"G:\Meu Drive\AUTOMAÇÃO EXPORTAÇÃO")
BITRIX_DL_BASE = AUTOMACAO_DIR / "downloads" / "audio"

_MESES_PT = {
    "01": "JANEIRO",  "02": "FEVEREIRO", "03": "MARÇO",    "04": "ABRIL",
    "05": "MAIO",     "06": "JUNHO",     "07": "JULHO",    "08": "AGOSTO",
    "09": "SETEMBRO", "10": "OUTUBRO",   "11": "NOVEMBRO", "12": "DEZEMBRO",
}

# ── logging ───────────────────────────────────────────────────────────────────
LOG = logging.getLogger("ragaz_pipeline_v12")
LOG.setLevel(logging.INFO)
_fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
_sh  = logging.StreamHandler(sys.stdout)
_sh.setFormatter(_fmt)
LOG.addHandler(_sh)


def _add_file_log(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    fh = logging.FileHandler(
        output_dir / f"pipeline_{datetime.datetime.now():%Y%m%d_%H%M%S}.log",
        encoding="utf-8",
    )
    fh.setFormatter(_fmt)
    LOG.addHandler(fh)


# ── utilitários ───────────────────────────────────────────────────────────────

def _run(cmd: list, desc: str, env_extra: dict | None = None) -> int:
    LOG.info(f"▶ {desc}")
    env = {**os.environ, **(env_extra or {})}
    result = subprocess.run(cmd, env=env)
    status = "OK" if result.returncode == 0 else f"FALHOU (exit {result.returncode})"
    LOG.info(f"  {status}: {desc}")
    return result.returncode


def _pasta_audios_v7(mes_ym: str) -> Path:
    ano, mes = mes_ym.split("-")
    return AUDIOS_BASE / f"{_MESES_PT[mes]}_{ano}"


def _encontrar_relatorio_mais_recente(depois_de: datetime.datetime | None = None) -> Path | None:
    """Retorna o Excel V7 mais recente. Se depois_de for informado, só considera
    relatórios criados a partir desse instante (evita pegar dry-runs anteriores)."""
    candidatos: list[Path] = []
    for pasta in BASE_DIR.glob("RELATÓRIOS*"):
        candidatos += list(pasta.glob("RELATORIO_*.xlsx"))
    if depois_de:
        min_ts = depois_de.timestamp()
        candidatos = [c for c in candidatos if c.stat().st_mtime >= min_ts]
    return max(candidatos, key=lambda p: p.stat().st_mtime) if candidatos else None


def calcular_mes(mes_arg: str) -> str:
    if mes_arg:
        if not re.match(r"^\d{4}-\d{2}$", mes_arg):
            raise ValueError(f"--mes deve ser YYYY-MM (ex: 2026-07), recebido: {mes_arg!r}")
        return mes_arg
    return datetime.date.today().strftime("%Y-%m")


def _checar_api_key() -> bool:
    if not os.environ.get("ANTHROPIC_API_KEY", "").strip():
        LOG.error(
            "ANTHROPIC_API_KEY não configurada.\n"
            "  No PowerShell: $env:ANTHROPIC_API_KEY = 'sk-ant-...'"
        )
        return False
    return True


# ── FASE 1 — Download XLS ─────────────────────────────────────────────────────

def fase1_xls(headed: bool = False) -> bool:
    LOG.info("=" * 65)
    LOG.info("FASE 1 — Download XLS  (Bitrix24 → local, d-1)")
    LOG.info("=" * 65)

    script = AUTOMACAO_DIR / "bitrix_export_drive.py"
    if not script.exists():
        LOG.error(f"Script não encontrado: {script}"); return False

    rc = _run(
        [sys.executable, str(script)],
        "Download XLS Bitrix24",
        {"HEADLESS": "0" if headed else "1", "DRIVE_DEST": str(BASE_DIR)},
    )
    if rc != 0:
        if not headed:
            LOG.warning("  → Se a sessão expirou: python V12_pipeline_ragaz.py --setup-sessao")
        return False

    xls = sorted(BASE_DIR.glob("*.xls*"), key=lambda p: p.stat().st_mtime, reverse=True)
    if xls:
        LOG.info(f"  XLS disponível: {xls[0].name}")
    return True


# ── FASE 2 — Download Áudios ──────────────────────────────────────────────────

def fase2_audios(mes_ym: str) -> bool:
    LOG.info("=" * 65)
    LOG.info(f"FASE 2 — Download Áudios  [{mes_ym}]  (Bitrix24 Drive)")
    LOG.info("=" * 65)

    script = AUTOMACAO_DIR / "bitrix_audio_download.py"
    if not script.exists():
        LOG.error(f"Script não encontrado: {script}"); return False

    rc = _run(
        [sys.executable, str(script)],
        f"Download áudios {mes_ym}",
        {"MES_ALVO": mes_ym, "AUDIO_DEST": str(BITRIX_DL_BASE)},
    )
    return rc == 0


# ── FASE 3 — Mover Áudios ─────────────────────────────────────────────────────

def fase3_mover(mes_ym: str) -> bool:
    LOG.info("=" * 65)
    LOG.info(f"FASE 3 — Mover áudios → pasta V7  [{mes_ym}]")
    LOG.info("=" * 65)

    origem  = BITRIX_DL_BASE / mes_ym
    destino = _pasta_audios_v7(mes_ym)
    destino.mkdir(parents=True, exist_ok=True)

    if not origem.exists():
        LOG.warning(f"Pasta de origem inexistente: {origem}"); return True

    movidos = pulados = 0
    for f in origem.glob("*.mp3"):
        dst = destino / f.name
        if dst.exists():
            pulados += 1
        else:
            shutil.move(str(f), str(dst))
            movidos += 1

    LOG.info(f"  Movidos: {movidos} | Já existiam: {pulados} | Total em {destino.name}: {movidos+pulados}")
    return True


# ── FASE 4 — V7 primeira passagem (Google + Claude) ───────────────────────────

def fase4_v7_primeira(so_txts: bool = False) -> bool:
    LOG.info("=" * 65)
    LOG.info("FASE 4 — V7 primeira passagem  (Google Speech + Claude)")
    LOG.info("=" * 65)

    if not _checar_api_key(): return False

    script = SCRIPTS_DIR / "V7_monitorias_ragaz.py"
    if not script.exists():
        LOG.error(f"Script não encontrado: {script}"); return False

    cmd = [sys.executable, str(script)]
    if so_txts:
        cmd.append("--so-txts")

    rc = _run(cmd, "V7 primeira passagem")
    return rc == 0


# ── FASE 5 — Whisper nos gaps ─────────────────────────────────────────────────

def fase5_whisper(
    mes_ym: str,
    modelo: str = "small",
    depois_de_fase4: datetime.datetime | None = None,
) -> bool:
    LOG.info("=" * 65)
    LOG.info(f"FASE 5 — Whisper reprocessamento de gaps  (modelo: {modelo})")
    LOG.info("=" * 65)

    script = SCRIPTS_DIR / "whisper_reprocessar_gaps.py"
    if not script.exists():
        LOG.error(f"Script não encontrado: {script}"); return False

    # Usa apenas relatórios gerados durante/após a Fase 4 desta rodada,
    # evitando que um dry-run anterior sirva de referência errada.
    relatorio = _encontrar_relatorio_mais_recente(depois_de=depois_de_fase4)
    if not relatorio:
        LOG.error("Nenhum relatório V7 encontrado. Execute a Fase 4 antes.")
        return False

    LOG.info(f"  Relatório de referência: {relatorio.name}")
    pasta_audios = _pasta_audios_v7(mes_ym)

    rc = _run(
        [
            sys.executable, str(script),
            "--relatorio", str(relatorio),
            "--audios",    str(pasta_audios),
            "--modelo",    modelo,
        ],
        f"Whisper gaps ({modelo})",
    )
    return rc == 0


# ── FASE 6 — V7 segunda passagem (só TXTs novos do Whisper) ──────────────────

def fase6_v7_final(novos_desde: str = "") -> bool:
    LOG.info("=" * 65)
    if novos_desde:
        LOG.info(f"FASE 6 — V7 passagem final  (só TXTs do Whisper após {novos_desde})")
    else:
        LOG.info("FASE 6 — V7 passagem final  (--so-txts)")
    LOG.info("=" * 65)

    if not _checar_api_key(): return False

    script = SCRIPTS_DIR / "V7_monitorias_ragaz.py"
    if not script.exists():
        LOG.error(f"Script não encontrado: {script}"); return False

    cmd = [sys.executable, str(script), "--so-txts"]
    if novos_desde:
        cmd += ["--novos-desde", novos_desde]

    desc = f"V7 passagem final (Whisper TXTs após {novos_desde})" if novos_desde else "V7 passagem final (--so-txts)"
    rc = _run(cmd, desc)
    return rc == 0


# ── FASE 7 — Exportar JSON ────────────────────────────────────────────────────

def fase7_exportar(mes_ym: str) -> bool:
    LOG.info("=" * 65)
    LOG.info(f"FASE 7 — Exportar transcrições JSON  [{mes_ym}]")
    LOG.info("=" * 65)

    script = SCRIPTS_DIR / "exportar_transcricoes.py"
    if not script.exists():
        LOG.error(f"Script não encontrado: {script}"); return False

    rc = _run([sys.executable, str(script), "--mes", mes_ym], "Exportar JSON Lovable")
    return rc == 0


# ── Setup de sessão ───────────────────────────────────────────────────────────

def cmd_setup_sessao() -> None:
    LOG.info("Abrindo Bitrix24 em janela visível para autenticação...")
    LOG.info("Faça login, marque 'Confiar neste dispositivo' e aguarde o XLS ser baixado.")
    fase1_xls(headed=True)
    LOG.info("Sessão salva. Próximas rodadas usarão headless automaticamente.")


# ── Agendamento ───────────────────────────────────────────────────────────────

def cmd_agendar(hora: str) -> None:
    script   = str(Path(__file__).resolve())
    nome     = "RagazPipelineV12"
    task_cmd = f'"{sys.executable}" "{script}"'
    ps_cmd   = (
        f'schtasks /Create /F /TN "{nome}" '
        f'/TR \'cmd /c "{task_cmd}"\' '
        f'/SC DAILY /ST {hora} /RL HIGHEST'
    )
    LOG.info(f"Criando tarefa '{nome}' às {hora} todos os dias...")
    rc = subprocess.run(ps_cmd, shell=True).returncode
    if rc == 0:
        LOG.info(f"Tarefa '{nome}' criada. Roda diariamente às {hora}.")
        LOG.info("Gerenciar: Painel de Controle → Agendador de Tarefas")
    else:
        LOG.error("Falha ao criar tarefa. Execute o PowerShell como Administrador.")


# ── main ──────────────────────────────────────────────────────────────────────

LABELS = {
    1: "XLS Download      ",
    2: "Áudios Download   ",
    3: "Mover Áudios      ",
    4: "V7 (Google+Claude)",
    5: "Whisper gaps      ",
    6: "V7 final (so-txts)",
    7: "Exportar JSON     ",
}

def main() -> int:
    parser = argparse.ArgumentParser(
        description="RAGAZ Pipeline Unificado V12",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--mes",            default="",
                        help="Mês YYYY-MM (padrão: mês corrente)")
    parser.add_argument("--fases",          default="1,2,3,4,5,6,7",
                        help="Fases a executar, ex: --fases 4,5,6,7")
    parser.add_argument("--modelo-whisper", default="small",
                        choices=["tiny", "base", "small", "medium"],
                        help="Modelo Whisper para Fase 5 (padrão: small)")
    parser.add_argument("--setup-sessao",   action="store_true",
                        help="Abre Bitrix24 em janela para renovar login")
    parser.add_argument("--agendar",        action="store_true",
                        help="Registra no Agendador de Tarefas do Windows")
    parser.add_argument("--hora",           default="22:00",
                        help="Hora do agendamento HH:MM (padrão: 22:00)")
    args = parser.parse_args()

    if args.setup_sessao:
        cmd_setup_sessao(); return 0

    mes_ym = calcular_mes(args.mes)

    if args.agendar:
        cmd_agendar(args.hora); return 0

    log_dir = BASE_DIR / f"RELATÓRIOS {datetime.datetime.now():%d.%m.%y}" / "logs"
    _add_file_log(log_dir)

    fases_rodar = sorted({
        int(f.strip()) for f in args.fases.split(",") if f.strip().isdigit()
    })

    # Checar API key antes de iniciar qualquer download se as fases 4 ou 6 estão no escopo
    if {4, 6} & set(fases_rodar) and not _checar_api_key():
        return 1

    LOG.info("=" * 65)
    LOG.info(f"RAGAZ PIPELINE V12 — {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    LOG.info(f"Mês de referência  : {mes_ym}")
    LOG.info(f"Fases              : {fases_rodar}")
    LOG.info(f"Modelo Whisper     : {args.modelo_whisper}")
    LOG.info("=" * 65)

    resultados: dict[int, bool] = {}
    ts_antes_whisper: str = ""
    ts_inicio_fase4: datetime.datetime | None = None  # usado pela Fase 5 para achar o Excel certo

    if 1 in fases_rodar:
        resultados[1] = fase1_xls()

    if 2 in fases_rodar:
        resultados[2] = fase2_audios(mes_ym)

    if 3 in fases_rodar:
        resultados[3] = fase3_mover(mes_ym)

    if 4 in fases_rodar:
        pre_falhas = [f for f in (1, 2) if f in fases_rodar and not resultados.get(f, True)]
        if pre_falhas:
            LOG.warning(f"Fase 4 pulada: fases {pre_falhas} falharam.")
            resultados[4] = False
        else:
            ts_inicio_fase4 = datetime.datetime.now()
            resultados[4] = fase4_v7_primeira()

    if 5 in fases_rodar:
        # Registra timestamp ANTES do Whisper — Fase 6 usará para filtrar TXTs novos
        ts_antes_whisper = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        LOG.info(f"  Timestamp pré-Whisper: {ts_antes_whisper}")
        resultados[5] = fase5_whisper(
            mes_ym, args.modelo_whisper,
            depois_de_fase4=ts_inicio_fase4,  # None se Fase 4 não rodou → pega mais recente
        )

    if 6 in fases_rodar:
        if 5 in fases_rodar and not resultados.get(5, True):
            LOG.warning("Fase 6 pulada: Whisper (Fase 5) falhou.")
            resultados[6] = False
        else:
            novos_desde = ts_antes_whisper if (5 in fases_rodar and ts_antes_whisper) else ""
            resultados[6] = fase6_v7_final(novos_desde=novos_desde)

    if 7 in fases_rodar:
        resultados[7] = fase7_exportar(mes_ym)

    # ── resumo ────────────────────────────────────────────────────────────────
    LOG.info("\n" + "=" * 65)
    LOG.info("  PIPELINE V12 — RESULTADO FINAL")
    LOG.info("=" * 65)
    for f in fases_rodar:
        ok     = resultados.get(f, False)
        status = "✓ OK     " if ok else "✗ FALHOU "
        LOG.info(f"  Fase {f} — {LABELS[f]} {status}")
    LOG.info("=" * 65)

    falhas = [f for f in fases_rodar if not resultados.get(f, False)]
    return 0 if not falhas else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        LOG.warning("Interrompido (Ctrl+C).")
        sys.exit(130)
    except Exception as e:
        LOG.error(f"ERRO FATAL: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
