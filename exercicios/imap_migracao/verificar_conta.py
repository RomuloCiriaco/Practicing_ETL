#!/usr/bin/env python3
"""
Verifica estrutura IMAP Localweb × KingHost/Titan.

Modos:
  - unica  → 1 conta do .env
  - lote   → várias contas da planilha de alinhamento

Somente leitura — não copia mensagens.
Gera CSV + Excel em saidas/.
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
from dotenv import load_dotenv

from imap_estrutura import (
    analisar_conta,
    estrutura_para_rows,
)
from planilha_contas import (
    carregar_alinhamento,
    filtrar_contas,
    reconstruir_de_fontes,
)

PASTA = Path(__file__).resolve().parent
PASTA_LOGS = PASTA / "logs"
PASTA_SAIDAS = PASTA / "saidas"

def _path_from_env(key: str, default: str = "") -> Path:
    return Path(os.getenv(key, default).strip() or default)


# Defaults genéricos — sobrescreva no .env (dados reais ficam fora do Git)
DEFAULT_ALINHAMENTO = _path_from_env(
    "PLANILHA_ALINHAMENTO", "dados/planilhas/alinhamento.xlsx"
)
DEFAULT_KING = _path_from_env("PLANILHA_KINGHOST", "dados/planilhas/destino.xlsx")
DEFAULT_SENHAS = _path_from_env("PLANILHA_NOVASSENHAS", "dados/planilhas/senhas.xlsx")
USB_OUTPUT_DIR = os.getenv("USB_OUTPUT_DIR", "").strip()


def configurar_logging(prefixo: str = "verificacao") -> Path:
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    PASTA_SAIDAS.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = PASTA_LOGS / f"{prefixo}_{stamp}.log"

    root = logging.getLogger()
    root.handlers.clear()
    root.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    for handler in (
        logging.FileHandler(log_path, encoding="utf-8"),
        logging.StreamHandler(),
    ):
        handler.setFormatter(fmt)
        root.addHandler(handler)
    return log_path


def hosts_from_env() -> Dict[str, Any]:
    return {
        "lw_host": os.getenv("LOCALWEB_IMAP_HOST", "email-ssl.com.br").strip(),
        "lw_port": int(os.getenv("LOCALWEB_IMAP_PORT", "993")),
        "kh_host": os.getenv("KINGHOST_IMAP_HOST", "imap.titan.email").strip(),
        "kh_port": int(os.getenv("KINGHOST_IMAP_PORT", "993")),
        "modo": (os.getenv("COMPARE_MODE") or "ALL").strip().upper(),
    }


def cfg_conta(email: str, senha_lw: str, senha_kh: str, hosts: Dict[str, Any]) -> dict:
    return {
        "localweb": {
            "email": email,
            "password": senha_lw,
            "host": hosts["lw_host"],
            "port": hosts["lw_port"],
        },
        "kinghost": {
            "email": email,
            "password": senha_kh,
            "host": hosts["kh_host"],
            "port": hosts["kh_port"],
        },
        "modo": hosts["modo"],
    }


def carregar_config_unica() -> dict:
    load_dotenv(PASTA / ".env")
    obrigatorios = [
        "LOCALWEB_EMAIL",
        "LOCALWEB_PASSWORD",
        "KINGHOST_EMAIL",
        "KINGHOST_PASSWORD",
    ]
    faltando = [k for k in obrigatorios if not os.getenv(k)]
    if faltando:
        raise SystemExit(
            "Modo unica: preencha o .env. Faltando: " + ", ".join(faltando)
        )
    hosts = hosts_from_env()
    if hosts["modo"] not in {"INBOX", "ALL"}:
        raise SystemExit("COMPARE_MODE deve ser INBOX ou ALL")
    return {
        "localweb": {
            "email": os.getenv("LOCALWEB_EMAIL", "").strip(),
            "password": os.getenv("LOCALWEB_PASSWORD", ""),
            "host": hosts["lw_host"],
            "port": hosts["lw_port"],
        },
        "kinghost": {
            "email": os.getenv("KINGHOST_EMAIL", "").strip(),
            "password": os.getenv("KINGHOST_PASSWORD", ""),
            "host": hosts["kh_host"],
            "port": hosts["kh_port"],
        },
        "modo": hosts["modo"],
    }


def salvar_csv(path: Path, rows: List[Dict[str, Any]], fieldnames: List[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def salvar_excel(path: Path, abas: Dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for nome, df in abas.items():
            safe = nome[:31] or "aba"
            (df if not df.empty else pd.DataFrame({"info": ["sem dados"]})).to_excel(
                writer, sheet_name=safe, index=False
            )


def resumo_de_resultado(res: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "email": res["email"],
        "ok": "SIM" if res["ok"] else "NÃO",
        "login_locaweb": res["login_locaweb"],
        "login_kinghost": res["login_kinghost"],
        "qtd_pastas_locaweb": len(res.get("origem") or []),
        "qtd_pastas_kinghost": len(res.get("destino") or []),
        "inbox_locaweb": res.get("inbox_locaweb", ""),
        "inbox_kinghost": res.get("inbox_kinghost", ""),
        "status_inbox": res.get("status_inbox", ""),
        "status_geral": res.get("status_geral", ""),
        "erro": res.get("erro", ""),
    }


def detalhes_mapeamento(res: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows = []
    for m in res.get("mapeamento") or []:
        row = {"email": res["email"]}
        row.update(m)
        rows.append(row)
    return rows


def imprimir_arvore(titulo: str, pastas: List[Dict[str, Any]]) -> None:
    print(f"\n=== {titulo} ===")
    for p in pastas:
        delim = p.get("delimiter") or "."
        profundidade = max(0, (p.get("nivel") or 1) - 1)
        indent = "  " * profundidade
        msgs = p.get("mensagens")
        msgs_txt = "?" if msgs is None else str(msgs)
        print(
            f"{indent}- {p['nome']}  [{p['papel']}]  "
            f"msgs={msgs_txt}  delim={delim!r}"
        )


def rodar_unica(verbose: bool = True) -> Dict[str, Path]:
    log_path = configurar_logging("estrutura")
    cfg = carregar_config_unica()
    logging.info("Modo UNICA | %s | compare=%s", cfg["localweb"]["email"], cfg["modo"])

    res = analisar_conta(cfg, verbose=verbose)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    resumo_df = pd.DataFrame([resumo_de_resultado(res)])
    map_rows = detalhes_mapeamento(res)
    map_df = pd.DataFrame(map_rows) if map_rows else pd.DataFrame()
    lw_df = pd.DataFrame(estrutura_para_rows(res.get("origem") or []))
    kh_df = pd.DataFrame(estrutura_para_rows(res.get("destino") or []))

    xlsx = PASTA_SAIDAS / f"verificacao_{stamp}.xlsx"
    csv_resumo = PASTA_SAIDAS / f"resumo_{stamp}.csv"
    csv_map = PASTA_SAIDAS / f"mapeamento_pastas_{stamp}.csv"
    json_path = PASTA_SAIDAS / f"estrutura_completa_{stamp}.json"

    salvar_excel(
        xlsx,
        {
            "resumo": resumo_df,
            "mapeamento": map_df,
            "estrutura_locaweb": lw_df,
            "estrutura_kinghost": kh_df,
        },
    )
    resumo_df.to_csv(csv_resumo, index=False)
    if not map_df.empty:
        map_df.to_csv(csv_map, index=False)

    json_path.write_text(
        json.dumps(
            {
                "gerado_em": datetime.now().isoformat(timespec="seconds"),
                "modo_execucao": "unica",
                "email": res["email"],
                "resumo": resumo_de_resultado(res),
                "mapeamento": map_rows,
                "locaweb": estrutura_para_rows(res.get("origem") or []),
                "kinghost": estrutura_para_rows(res.get("destino") or []),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    if res.get("origem"):
        imprimir_arvore("Estrutura Localweb", res["origem"])
    if res.get("destino"):
        imprimir_arvore("Estrutura KingHost / Titan", res["destino"])

    print("\n=== Mapeamento ===")
    for row in map_rows:
        src = row.get("pasta_locaweb") or "(—)"
        dst = row.get("pasta_kinghost") or "(criar?)"
        print(
            f"  {src:<28} → {dst:<20} | "
            f"{row.get('msgs_locaweb', '-'):>6} → {row.get('msgs_kinghost', '-'):>6} | "
            f"{row.get('status')}"
        )

    print(f"\nPlanilha: {xlsx}")
    print(f"Log:      {log_path}")
    return {"xlsx": xlsx, "log": log_path}


def carregar_contas_lote(args: argparse.Namespace) -> List[Dict[str, Any]]:
    load_dotenv(PASTA / ".env")

    alinhamento = Path(
        args.planilha
        or os.getenv("PLANILHA_ALINHAMENTO")
        or DEFAULT_ALINHAMENTO
    )
    king = Path(args.planilha_king or os.getenv("PLANILHA_KINGHOST") or DEFAULT_KING)
    senhas = Path(
        args.planilha_senhas or os.getenv("PLANILHA_NOVASSENHAS") or DEFAULT_SENHAS
    )

    if alinhamento.exists():
        logging.info("Usando alinhamento: %s", alinhamento)
        df = carregar_alinhamento(alinhamento)
    else:
        logging.warning(
            "Alinhamento não encontrado (%s). Reconstruindo de King + NOVASSENHAS…",
            alinhamento,
        )
        df = reconstruir_de_fontes(king, senhas)

    contas = filtrar_contas(
        df,
        apenas_nos_dois=not args.todas,
        somente_com_senhas=not args.sem_filtro_senha,
        pular_migrados=args.pular_migrados,
        email=args.email,
        limite=args.limite,
    )
    logging.info("Contas selecionadas para o lote: %d", len(contas))
    if not contas:
        raise SystemExit("Nenhuma conta elegível na planilha com os filtros atuais.")
    return contas


def rodar_lote(args: argparse.Namespace) -> Dict[str, Path]:
    log_path = configurar_logging("lote")
    load_dotenv(PASTA / ".env")
    hosts = hosts_from_env()
    if args.compare_mode:
        hosts["modo"] = args.compare_mode.upper()
    if hosts["modo"] not in {"INBOX", "ALL"}:
        raise SystemExit("COMPARE_MODE deve ser INBOX ou ALL")

    logging.info("Modo LOTE | compare=%s", hosts["modo"])
    contas = carregar_contas_lote(args)

    resumos: List[Dict[str, Any]] = []
    detalhes: List[Dict[str, Any]] = []
    estruturas_lw: List[Dict[str, Any]] = []
    estruturas_kh: List[Dict[str, Any]] = []

    total = len(contas)
    for i, conta in enumerate(contas, 1):
        email = conta["email"]
        logging.info("==== [%d/%d] %s ====", i, total, email)
        cfg = cfg_conta(
            email, conta["senha_locaweb"], conta["senha_kinghost"], hosts
        )
        try:
            res = analisar_conta(cfg, verbose=False)
        except Exception as exc:
            logging.exception("Falha inesperada em %s: %s", email, exc)
            res = {
                "email": email,
                "ok": False,
                "erro": str(exc),
                "login_locaweb": "?",
                "login_kinghost": "?",
                "origem": [],
                "destino": [],
                "mapeamento": [],
                "inbox_locaweb": "",
                "inbox_kinghost": "",
                "status_inbox": "ERRO",
                "status_geral": "ERRO",
            }
        resumos.append(resumo_de_resultado(res))
        detalhes.extend(detalhes_mapeamento(res))
        for row in estrutura_para_rows(res.get("origem") or []):
            row = dict(row)
            row["email"] = email
            estruturas_lw.append(row)
        for row in estrutura_para_rows(res.get("destino") or []):
            row = dict(row)
            row["email"] = email
            estruturas_kh.append(row)
        logging.info(
            "  → inbox %s→%s | geral=%s | ok=%s",
            res.get("inbox_locaweb"),
            res.get("inbox_kinghost"),
            res.get("status_geral"),
            res.get("ok"),
        )

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    resumo_df = pd.DataFrame(resumos)
    map_df = pd.DataFrame(detalhes) if detalhes else pd.DataFrame()
    lw_df = pd.DataFrame(estruturas_lw) if estruturas_lw else pd.DataFrame()
    kh_df = pd.DataFrame(estruturas_kh) if estruturas_kh else pd.DataFrame()

    # ordenar resumo pelo status (falhas primeiro)
    if not resumo_df.empty and "status_geral" in resumo_df.columns:
        ordem = {
            "LOGIN_FALHOU": 0,
            "ERRO": 1,
            "SEM_PAR_NO_DESTINO": 2,
            "SO_ORIGEM": 3,
            "FALTA_MIGRAR_OU_PARCIAL": 4,
            "SO_DESTINO": 5,
            "SO_DESTINO_SEM_ORIGEM": 6,
            "AMBOS_VAZIOS": 7,
            "OK_PROVAVELMENTE_MIGRADA": 8,
        }
        resumo_df["_ord"] = resumo_df["status_geral"].map(lambda s: ordem.get(s, 99))
        resumo_df = resumo_df.sort_values(["_ord", "email"]).drop(columns=["_ord"])

    xlsx = PASTA_SAIDAS / f"lote_verificacao_{stamp}.xlsx"
    csv_resumo = PASTA_SAIDAS / f"lote_resumo_{stamp}.csv"
    csv_map = PASTA_SAIDAS / f"lote_mapeamento_{stamp}.csv"
    # cópia opcional se USB_OUTPUT_DIR estiver definido no .env
    usb_dir = Path(USB_OUTPUT_DIR) if USB_OUTPUT_DIR else None
    xlsx_usb = (
        usb_dir / f"lote_verificacao_{stamp}.xlsx"
        if usb_dir is not None and usb_dir.exists()
        else None
    )

    salvar_excel(
        xlsx,
        {
            "resumo": resumo_df,
            "mapeamento": map_df,
            "estrutura_locaweb": lw_df,
            "estrutura_kinghost": kh_df,
        },
    )
    if xlsx_usb is not None:
        try:
            salvar_excel(
                xlsx_usb,
                {
                    "resumo": resumo_df,
                    "mapeamento": map_df,
                    "estrutura_locaweb": lw_df,
                    "estrutura_kinghost": kh_df,
                },
            )
        except Exception as exc:
            logging.warning("Não gravou no pendrive: %s", exc)
            xlsx_usb = None

    resumo_df.to_csv(csv_resumo, index=False)
    if not map_df.empty:
        map_df.to_csv(csv_map, index=False)

    print("\n=== Resumo do lote ===")
    if not resumo_df.empty:
        counts = resumo_df["status_geral"].value_counts().to_dict()
        for status, qtd in counts.items():
            print(f"  {status}: {qtd}")
        print("\nContas:")
        for _, row in resumo_df.iterrows():
            print(
                f"  {row['email']:<40} inbox "
                f"{row['inbox_locaweb']}→{row['inbox_kinghost']} | {row['status_geral']}"
            )

    print(f"\nPlanilha: {xlsx}")
    if xlsx_usb:
        print(f"Cópia USB: {xlsx_usb}")
    print(f"CSV:      {csv_resumo}")
    print(f"Log:      {log_path}")
    return {"xlsx": xlsx, "log": log_path}


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Verificação IMAP Localweb × KingHost (unica ou lote)."
    )
    p.add_argument(
        "--modo",
        choices=["unica", "lote"],
        default=os.getenv("EXEC_MODE", "unica"),
        help="unica=conta do .env | lote=planilha (default: unica ou EXEC_MODE)",
    )
    p.add_argument(
        "--planilha",
        default=None,
        help="alinhamento_contas_king_locaweb.xlsx",
    )
    p.add_argument("--planilha-king", default=None, help="listagemsatoKingHOST.xlsx")
    p.add_argument("--planilha-senhas", default=None, help="NOVASSENHAS.xlsx")
    p.add_argument("--limite", type=int, default=None, help="Máx. de contas no lote")
    p.add_argument("--email", default=None, help="Filtra um e-mail específico")
    p.add_argument(
        "--todas",
        action="store_true",
        help="Não exige NOS_DOIS (ainda precisa senhas, salvo --sem-filtro-senha)",
    )
    p.add_argument(
        "--sem-filtro-senha",
        action="store_true",
        help="Inclui contas sem senha (vão falhar no login)",
    )
    p.add_argument(
        "--pular-migrados",
        action="store_true",
        help="Ignora linhas com migrado=SIM na planilha",
    )
    p.add_argument(
        "--compare-mode",
        choices=["INBOX", "ALL"],
        default=None,
        help="Sobrescreve COMPARE_MODE do .env",
    )
    return p


def main() -> None:
    load_dotenv(PASTA / ".env")
    args = build_parser().parse_args()
    if args.modo == "lote":
        rodar_lote(args)
    else:
        rodar_unica(verbose=True)


if __name__ == "__main__":
    main()
