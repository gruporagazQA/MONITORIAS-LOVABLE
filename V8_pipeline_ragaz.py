#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 RAGAZ — PIPELINE UNIFICADO V8
================================================================================
 Orquestra o ciclo completo de monitorias em 5 fases:

   1. Download XLS de ligações do Bitrix24 (Playwright → headless)
   2. Download áudios MP3 do Bitrix24 Drive (REST API)
   3. Mover áudios para pasta do V7  (AUDIOS/MES_ANO)
   4. Análise V7  (transcrição Google Speech + Claude Haiku)
   5. Exportar JSON de transcrições para Lovable AI

 Cobertura: sempre d-1 do mês de referência
   (ex: rodando em 18/06 cobre 01/06 → 17/06)

 USO (PowerShell):
   # Ciclo completo (mês corrente)
   python "...\V8_pipeline_ragaz.py"

   # Mês específico
   python "...\V8_pipeline_ragaz.py" --mes 2026-07

   # Pular fases 1 e 2 (downloads já feitos)
   python "...\V8_pipeline_ragaz.py" --fases 3,4,5

   # Primeiro login manual (necessário apenas uma vez)
   python "...\V8_pipeline_ragaz.py" --setup-sessao

   # Registrar no Agendador de Tarefas do Windows (07:00 diário)
   python "...\V8_pipeline_ragaz.py" --agendar --hora 07:00

 PRÉ-REQUISITOS:
   .env em AUTOMAÇÃO EXPORTAÇÃO\ com BITRIX_WEBHOOK e (opcional) BITRIX_URL
   $env:ANTHROPIC_API_KEY = "sk-ant-..."
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
import time
from pathlib import Path

# ── configuração ──────────────────────────────────────────────────────────────
BASE_DIR       = Path(r"G:\Meu Drive\ARQUITETURA LOVABLE")
SCRIPTS_DIR    = BASE_DIR / "LINGUAGEM EM CÓDIGO"
AUDIOS_BASE    = BASE_DIR / "AUDIOS"
AUTOMACAO_DIR  = Path(r"G:\Meu Drive\AUTOMAÇÃO EXPORTAÇÃO")

# Onde bitrix_audio_download.py salva os MP3s
BITRIX_DL_BASE = AUTOMACAO_DIR / "downloads" / "audio"

_MESES_PT = {
    "01": "JANEIRO",  "02": "FEVEREIRO", "03": "MARÇO",    "04": "ABRIL",
    "05": "MAIO",     "06": "JUNHO",     "07": "JULHO",    "08": "AGOSTO",
    "09": "SETEMBRO", "10": "OUTUBRO",   "11": "NOVEMBRO", "12": "DEZEMBRO",
}

# ── logging ───────────────────────────────────────────────────────────────────
LOG = logging.getLogger("ragaz_pipeline")
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
    """Executa subprocesso, exibe saída em tempo real, retorna exit code."""
    LOG.info(f"▶ {desc}")
    env = {**os.environ, **(env_extra or {})}
    result = subprocess.run(cmd, env=env)
    status = "OK" if result.returncode == 0 else f"FALHOU (exit {result.returncode})"
    LOG.info(f"  {status}: {desc}")
    return result.returncode


def _pasta_audios_v7(mes_ym: str) -> Path:
    """Pasta V7 de áudios: AUDIOS/JUNHO_2026"""
    ano, mes = mes_ym.split("-")
    return AUDIOS_BASE / f"{_MESES_PT[mes]}_{ano}"


def calcular_mes(mes_arg: str) -> str:
    if mes_arg:
        if not re.match(r"^\d{4}-\d{2}$", mes_arg):
            raise ValueError(f"--mes deve ser YYYY-MM (ex: 2026-07), recebido: {mes_arg!r}")
        return mes_arg
    return datetime.date.today().strftime("%Y-%m")


def _checar_api_key() -> bool:
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key:
        LOG.error(
            "ANTHROPIC_API_KEY não configurada.\n"
            "  No PowerShell: $env:ANTHROPIC_API_KEY = 'sk-ant-...'"
        )
        return False
    return True


# ── FASE 1 — Download XLS ─────────────────────────────────────────────────────

def fase1_xls(headed: bool = False) -> bool:
    """Baixa o XLS de Informações de Chamada do Bitrix24 via Playwright."""
    LOG.info("=" * 65)
    LOG.info("FASE 1 — Download XLS (Bitrix24 → local)")
    LOG.info("=" * 65)

    script = AUTOMACAO_DIR / "bitrix_export_drive.py"
    if not script.exists():
        LOG.error(f"Script não encontrado: {script}")
        return False

    env_extra = {"HEADLESS": "0" if headed else "1",
                 "DRIVE_DEST": str(BASE_DIR)}   # salva direto na BASE_DIR

    rc = _run([sys.executable, str(script)], "Download XLS Bitrix24", env_extra)
    if rc != 0:
        if not headed:
            LOG.warning(
                "  → Se a sessão expirou, renove o login com:\n"
                "    python pipeline_ragaz.py --setup-sessao"
            )
        return False

    # Confirmar que o XLS chegou em BASE_DIR
    xls_list = sorted(BASE_DIR.glob("*.xls*"), key=lambda p: p.stat().st_mtime, reverse=True)
    if xls_list:
        LOG.info(f"  XLS disponível: {xls_list[0].name}")
    return True


# ── FASE 2 — Download Áudios ──────────────────────────────────────────────────

def fase2_audios(mes_ym: str) -> bool:
    """Baixa áudios MP3 do Bitrix24 Drive via REST API."""
    LOG.info("=" * 65)
    LOG.info(f"FASE 2 — Download Áudios  [{mes_ym}]  (Bitrix24 Drive → local)")
    LOG.info("=" * 65)

    script = AUTOMACAO_DIR / "bitrix_audio_download.py"
    if not script.exists():
        LOG.error(f"Script não encontrado: {script}")
        return False

    rc = _run(
        [sys.executable, str(script)],
        f"Download áudios {mes_ym}",
        env_extra={
            "MES_ALVO":   mes_ym,
            "AUDIO_DEST": str(BITRIX_DL_BASE),
        },
    )
    return rc == 0


# ── FASE 3 — Mover Áudios ─────────────────────────────────────────────────────

def fase3_mover(mes_ym: str) -> int:
    """Move MP3s de AUTOMAÇÃO EXPORTAÇÃO/.../YYYY-MM para AUDIOS/MES_ANO."""
    LOG.info("=" * 65)
    LOG.info(f"FASE 3 — Mover áudios → pasta V7  [{mes_ym}]")
    LOG.info("=" * 65)

    origem  = BITRIX_DL_BASE / mes_ym
    destino = _pasta_audios_v7(mes_ym)
    destino.mkdir(parents=True, exist_ok=True)

    if not origem.exists():
        LOG.warning(f"Pasta de origem inexistente (download pode ter falhado): {origem}")
        return 0

    mp3s = list(origem.glob("*.mp3"))
    movidos = pulados = 0
    for f in mp3s:
        dst = destino / f.name
        if dst.exists():
            pulados += 1
            continue
        shutil.move(str(f), str(dst))
        movidos += 1

    total = (destino.glob("*.mp3"))
    LOG.info(
        f"  Movidos: {movidos} | Já existiam: {pulados} | "
        f"Total em {destino.name}: {movidos + pulados}"
    )
    return movidos + pulados


# ── FASE 4 — Análise V7 ───────────────────────────────────────────────────────

def fase4_analise(so_txts: bool = False) -> bool:
    """Roda o V7 (transcrição Google Speech + análise Claude)."""
    LOG.info("=" * 65)
    LOG.info("FASE 4 — Análise V7  (transcrição + Claude Haiku)")
    LOG.info("=" * 65)

    if not _checar_api_key():
        return False

    script = SCRIPTS_DIR / "V7_monitorias_ragaz.py"
    if not script.exists():
        LOG.error(f"Script não encontrado: {script}")
        return False

    cmd = [sys.executable, str(script)]
    if so_txts:
        cmd.append("--so-txts")

    rc = _run(cmd, "V7 análise completa")
    return rc == 0


# ── FASE 5 — Exportar JSON ────────────────────────────────────────────────────

def fase5_exportar(mes_ym: str) -> bool:
    """Exporta JSON de transcrições para Lovable AI."""
    LOG.info("=" * 65)
    LOG.info(f"FASE 5 — Exportar transcrições JSON  [{mes_ym}]")
    LOG.info("=" * 65)

    script = SCRIPTS_DIR / "exportar_transcricoes.py"
    if not script.exists():
        LOG.error(f"Script não encontrado: {script}")
        return False

    rc = _run([sys.executable, str(script), "--mes", mes_ym], "Exportar JSON Lovable")
    return rc == 0


# ── Setup de sessão Playwright ────────────────────────────────────────────────

def cmd_setup_sessao() -> None:
    """Abre o Chromium em modo headed para login manual no Bitrix24."""
    LOG.info("Abrindo Bitrix24 em janela visível para autenticação...")
    LOG.info("Faça login, marque 'Confiar neste dispositivo' e aguarde o download do XLS.")
    LOG.info("Após concluir, feche o navegador. As próximas execuções serão automáticas.")
    fase1_xls(headed=True)
    LOG.info("Sessão salva. Próximas rodadas usarão headless automaticamente.")


# ── Agendamento Windows Task Scheduler ────────────────────────────────────────

def cmd_agendar(hora: str, mes_ym: str) -> None:
    """Registra tarefa diária no Agendador de Tarefas do Windows."""
    script   = str(Path(__file__).resolve())
    nome     = "RagazPipelineMonitorias"
    # O pipeline sempre usa mês corrente quando chamado sem --mes
    task_cmd = f'"{sys.executable}" "{script}"'

    ps_cmd = (
        f'schtasks /Create /F /TN "{nome}" '
        f'/TR \'cmd /c "{task_cmd}"\' '
        f'/SC DAILY /ST {hora} /RL HIGHEST'
    )
    LOG.info(f"Criando tarefa '{nome}' às {hora} todos os dias...")
    rc = subprocess.run(ps_cmd, shell=True).returncode
    if rc == 0:
        LOG.info(f"Tarefa '{nome}' criada com sucesso.")
        LOG.info("Para editar/remover: Painel de Controle → Agendador de Tarefas")
    else:
        LOG.error("Falha ao criar tarefa. Execute o PowerShell como Administrador.")


# ── main ──────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(
        description="RAGAZ Pipeline Unificado",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--mes",          default="",
                        help="Mês alvo YYYY-MM (padrão: mês corrente)")
    parser.add_argument("--fases",        default="1,2,3,4,5",
                        help="Fases a executar, ex: --fases 3,4,5")
    parser.add_argument("--so-txts",      action="store_true",
                        help="Fase 4: pula transcrição de MP3, usa só TXTs existentes")
    parser.add_argument("--setup-sessao", action="store_true",
                        help="Abre Bitrix24 em janela para renovar login (run once)")
    parser.add_argument("--agendar",      action="store_true",
                        help="Registra no Agendador de Tarefas do Windows")
    parser.add_argument("--hora",         default="07:00",
                        help="Hora do agendamento HH:MM (padrão: 07:00)")
    args = parser.parse_args()

    # ── comandos de setup ────────────────────────────────────────────────────
    if args.setup_sessao:
        cmd_setup_sessao()
        return 0

    mes_ym = calcular_mes(args.mes)

    if args.agendar:
        cmd_agendar(args.hora, mes_ym)
        return 0

    # ── log de arquivo ───────────────────────────────────────────────────────
    log_dir = BASE_DIR / f"RELATÓRIOS {datetime.datetime.now():%d.%m.%y}" / "logs"
    _add_file_log(log_dir)

    fases_rodar = sorted({
        int(f.strip()) for f in args.fases.split(",") if f.strip().isdigit()
    })

    LOG.info("=" * 65)
    LOG.info(f"RAGAZ PIPELINE UNIFICADO — {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    LOG.info(f"Mês de referência : {mes_ym}")
    LOG.info(f"Fases             : {fases_rodar}")
    LOG.info(f"Só-TXTs (f4)      : {args.so_txts}")
    LOG.info("=" * 65)

    resultados: dict[int, bool] = {}

    if 1 in fases_rodar:
        resultados[1] = fase1_xls(headed=False)

    if 2 in fases_rodar:
        resultados[2] = fase2_audios(mes_ym)

    if 3 in fases_rodar:
        n = fase3_mover(mes_ym)
        resultados[3] = True   # mover nunca bloqueia o pipeline

    if 4 in fases_rodar:
        # Aborta fase 4 somente se 1 ou 2 falharam E faziam parte desta rodada
        fases_pre_falhas = [f for f in (1, 2) if f in fases_rodar and not resultados.get(f, True)]
        if fases_pre_falhas:
            LOG.warning(f"Fase 4 pulada: fases {fases_pre_falhas} falharam.")
            resultados[4] = False
        else:
            resultados[4] = fase4_analise(args.so_txts)

    if 5 in fases_rodar:
        resultados[5] = fase5_exportar(mes_ym)

    # ── resumo final ─────────────────────────────────────────────────────────
    labels = {
        1: "XLS Download    ",
        2: "Áudios Download ",
        3: "Mover Áudios    ",
        4: "Análise V7      ",
        5: "Exportar JSON   ",
    }
    LOG.info("\n" + "=" * 65)
    LOG.info("  PIPELINE — RESULTADO FINAL")
    LOG.info("=" * 65)
    for f in fases_rodar:
        ok     = resultados.get(f, False)
        status = "✓ OK     " if ok else "✗ FALHOU "
        LOG.info(f"  Fase {f} — {labels[f]} {status}")
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
