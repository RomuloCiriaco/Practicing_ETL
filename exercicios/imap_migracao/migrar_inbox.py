#!/usr/bin/env python3
"""
Migração IMAP em modo TESTE: Localweb → KingHost/Titan.

Por padrão:
  - só INBOX
  - poucas mensagens (--limite)
  - não duplica (checa Message-ID)
  - gera planilha/CSV em saidas/

Não apaga nada na origem.
"""

from __future__ import annotations

import argparse
import email
import imaplib
import logging
import os
import re
from datetime import datetime
from email.header import decode_header, make_header
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from dotenv import load_dotenv

from imap_estrutura import conectar_imap, status_mensagens
from planilha_contas import carregar_alinhamento, filtrar_contas

PASTA = Path(__file__).resolve().parent
PASTA_LOGS = PASTA / "logs"
PASTA_SAIDAS = PASTA / "saidas"

DEFAULT_ALINHAMENTO = Path(
    os.getenv("PLANILHA_ALINHAMENTO", "dados/planilhas/alinhamento.xlsx").strip()
    or "dados/planilhas/alinhamento.xlsx"
)
DEFAULT_CONTAS_TESTE = Path(
    os.getenv(
        "CONTAS_TESTE_CSV",
        str(PASTA / "contas_teste.csv"),
    ).strip()
    or str(PASTA / "contas_teste.csv")
)
USB_OUTPUT_DIR = os.getenv("USB_OUTPUT_DIR", "").strip()


def configurar_logging() -> Path:
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    PASTA_SAIDAS.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = PASTA_LOGS / f"migracao_{stamp}.log"
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
    load_dotenv(PASTA / ".env")
    return {
        "lw_host": os.getenv("LOCALWEB_IMAP_HOST", "email-ssl.com.br").strip(),
        "lw_port": int(os.getenv("LOCALWEB_IMAP_PORT", "993")),
        "kh_host": os.getenv("KINGHOST_IMAP_HOST", "imap.titan.email").strip(),
        "kh_port": int(os.getenv("KINGHOST_IMAP_PORT", "993")),
    }


def cfg_par(email_addr: str, senha_lw: str, senha_kh: str, hosts: Dict[str, Any]) -> dict:
    return {
        "localweb": {
            "email": email_addr,
            "password": senha_lw,
            "host": hosts["lw_host"],
            "port": hosts["lw_port"],
        },
        "kinghost": {
            "email": email_addr,
            "password": senha_kh,
            "host": hosts["kh_host"],
            "port": hosts["kh_port"],
        },
    }


def decodificar_assunto(raw: Optional[str]) -> str:
    if not raw:
        return ""
    try:
        return str(make_header(decode_header(raw)))
    except Exception:
        return raw


def extrair_message_id(raw: bytes) -> str:
    msg = email.message_from_bytes(raw)
    mid = msg.get("Message-ID") or msg.get("Message-Id") or ""
    return mid.strip()


def extrair_assunto(raw: bytes) -> str:
    msg = email.message_from_bytes(raw)
    return decodificar_assunto(msg.get("Subject"))


def selecionar_pasta(client: imaplib.IMAP4_SSL, pasta: str, readonly: bool = True) -> None:
    for ref in (f'"{pasta}"', pasta):
        st, _ = client.select(ref, readonly=readonly)
        if st == "OK":
            return
    raise RuntimeError(f"Não foi possível abrir a pasta {pasta!r}")


def listar_uids(client: imaplib.IMAP4_SSL, pasta: str) -> List[bytes]:
    selecionar_pasta(client, pasta, readonly=True)
    st, data = client.uid("search", None, "ALL")
    if st != "OK" or not data or data[0] is None:
        return []
    return data[0].split()


def carregar_message_ids_destino(
    client: imaplib.IMAP4_SSL, pasta: str, max_fetch: int = 5000
) -> set:
    """Indexa Message-IDs já presentes no destino (mais confiável que SEARCH HEADER)."""
    ids: set = set()
    uids = listar_uids(client, pasta)
    if len(uids) > max_fetch:
        logging.warning(
            "Destino tem %d msgs; indexando só as %d mais recentes p/ anti-dupe",
            len(uids),
            max_fetch,
        )
        uids = uids[-max_fetch:]
    selecionar_pasta(client, pasta, readonly=True)
    for uid in uids:
        try:
            st, data = client.uid("fetch", uid, "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)])")
            if st != "OK" or not data:
                continue
            for item in data:
                if isinstance(item, tuple) and len(item) >= 2 and isinstance(item[1], bytes):
                    mid = extrair_message_id(item[1])
                    if mid:
                        ids.add(mid.strip().lower())
                        if mid.startswith("<") and mid.endswith(">"):
                            ids.add(mid[1:-1].strip().lower())
        except Exception:
            continue
    logging.info("Indexados %d Message-IDs no destino (%s)", len(ids), pasta)
    return ids


def message_id_no_indice(message_id: str, indice: set) -> bool:
    if not message_id or not indice:
        return False
    mid = message_id.strip().lower()
    if mid in indice:
        return True
    if mid.startswith("<") and mid.endswith(">") and mid[1:-1] in indice:
        return True
    if f"<{mid}>" in indice:
        return True
    return False


def message_id_existe(client: imaplib.IMAP4_SSL, pasta: str, message_id: str) -> bool:
    if not message_id:
        return False
    selecionar_pasta(client, pasta, readonly=True)
    candidatos = [message_id]
    if message_id.startswith("<") and message_id.endswith(">"):
        candidatos.append(message_id[1:-1])
    else:
        candidatos.append(f"<{message_id}>")

    for mid in candidatos:
        try:
            st, data = client.uid("search", None, "HEADER", "Message-ID", mid)
            if st == "OK" and data and data[0] and data[0].split():
                return True
        except imaplib.IMAP4.error:
            continue
        try:
            crit = f'(HEADER Message-ID "{mid}")'
            st, data = client.uid("search", None, crit)
            if st == "OK" and data and data[0] and data[0].split():
                return True
        except imaplib.IMAP4.error:
            continue
    return False


def fetch_rfc822(client: imaplib.IMAP4_SSL, uid: bytes) -> Tuple[bytes, str, str]:
    """Retorna (raw, flags_str, internaldate_str)."""
    st, data = client.uid("fetch", uid, "(FLAGS INTERNALDATE BODY.PEEK[])")
    if st != "OK" or not data:
        raise RuntimeError(f"FETCH falhou uid={uid!r}")

    raw = b""
    flags = "()"
    internaldate = None

    for item in data:
        if not isinstance(item, tuple) or len(item) < 2:
            continue
        meta = item[0]
        body = item[1]
        if isinstance(body, bytes) and len(body) > 0:
            raw = body
        if isinstance(meta, bytes):
            meta_s = meta.decode("utf-8", errors="replace")
        else:
            meta_s = str(meta)
        m_flags = re.search(r"FLAGS\s*(\([^)]*\))", meta_s)
        if m_flags:
            flags = m_flags.group(1)
        m_date = re.search(r'INTERNALDATE\s+"([^"]+)"', meta_s)
        if m_date:
            internaldate = m_date.group(1)

    if not raw:
        # fallback RFC822
        st, data = client.uid("fetch", uid, "(RFC822)")
        if st != "OK" or not data or not isinstance(data[0], tuple):
            raise RuntimeError(f"FETCH RFC822 falhou uid={uid!r}")
        raw = data[0][1]

    # remove \Recent (não se APPEND)
    flags_limpos = flags
    for bad in ("\\Recent", "\\recent"):
        flags_limpos = flags_limpos.replace(bad, "")
    flags_limpos = re.sub(r"\s+", " ", flags_limpos).replace("( ", "(").replace(" )", ")")
    if flags_limpos in {"()", "( )", ""}:
        flags_limpos = None

    return raw, flags_limpos, internaldate or ""


def _date_for_append(internaldate: str):
    """Converte INTERNALDATE IMAP → struct_time (imaplib.append no Py3.8)."""
    if not internaldate:
        return None
    try:
        resp = f'INTERNALDATE "{internaldate}"'.encode("ascii", "replace")
        tt = imaplib.Internaldate2tuple(resp)
        return tt
    except Exception:
        return None


def append_mensagem(
    client: imaplib.IMAP4_SSL,
    pasta: str,
    raw: bytes,
    flags: Optional[str],
    internaldate: str,
) -> None:
    date_arg = _date_for_append(internaldate)
    tentativas = [
        (flags, date_arg),
        (None, date_arg),
        (flags, None),
        (None, None),
    ]
    ultimo_erro = None
    for fl, dt in tentativas:
        try:
            st, resp = client.append(pasta, fl, dt, raw)
            if st == "OK":
                return
            ultimo_erro = f"{st} {resp}"
        except Exception as exc:
            ultimo_erro = str(exc)
            continue
    raise RuntimeError(f"APPEND falhou: {ultimo_erro}")


def migrar_inbox_conta(
    cfg: dict,
    *,
    limite: int,
    dry_run: bool,
    pasta_origem: str = "INBOX",
    pasta_destino: str = "INBOX",
) -> Dict[str, Any]:
    email_addr = cfg["localweb"]["email"]
    resultado: Dict[str, Any] = {
        "email": email_addr,
        "ok": False,
        "dry_run": dry_run,
        "limite": limite,
        "inbox_antes_lw": "",
        "inbox_antes_kh": "",
        "inbox_depois_kh": "",
        "copiadas": 0,
        "ja_existiam": 0,
        "erros": 0,
        "detalhes": [],
        "erro": "",
    }

    lw = kh = None
    try:
        lw = conectar_imap("Localweb", cfg["localweb"])
        kh = conectar_imap("KingHost", cfg["kinghost"])

        q_lw, _ = status_mensagens(lw, pasta_origem)
        q_kh, _ = status_mensagens(kh, pasta_destino)
        resultado["inbox_antes_lw"] = q_lw if q_lw is not None else ""
        resultado["inbox_antes_kh"] = q_kh if q_kh is not None else ""
        logging.info(
            "INBOX antes | Localweb=%s | KingHost=%s | limite=%s | dry_run=%s",
            resultado["inbox_antes_lw"],
            resultado["inbox_antes_kh"],
            limite,
            dry_run,
        )

        uids = listar_uids(lw, pasta_origem)
        if not uids:
            logging.info("Origem sem mensagens na %s", pasta_origem)
            resultado["ok"] = True
            return resultado

        # pega as N mais recentes (UIDs maiores no fim)
        alvo = uids[-limite:] if limite > 0 else uids
        logging.info(
            "Selecionadas %d de %d mensagens da %s",
            len(alvo),
            len(uids),
            pasta_origem,
        )

        selecionar_pasta(lw, pasta_origem, readonly=True)
        selecionar_pasta(kh, pasta_destino, readonly=False)

        # índice anti-duplicata no destino (Titan nem sempre responde bem a SEARCH HEADER)
        indice_dest = carregar_message_ids_destino(kh, pasta_destino)

        for uid in alvo:
            uid_s = uid.decode() if isinstance(uid, bytes) else str(uid)
            detalhe = {
                "email": email_addr,
                "uid_origem": uid_s,
                "message_id": "",
                "assunto": "",
                "acao": "",
                "erro": "",
            }
            try:
                raw, flags, internaldate = fetch_rfc822(lw, uid)
                mid = extrair_message_id(raw)
                assunto = extrair_assunto(raw)
                detalhe["message_id"] = mid
                detalhe["assunto"] = assunto[:200]

                ja_existe = message_id_no_indice(mid, indice_dest) or (
                    mid and message_id_existe(kh, pasta_destino, mid)
                )
                if ja_existe:
                    detalhe["acao"] = "JA_EXISTIA"
                    resultado["ja_existiam"] += 1
                    logging.info("UID %s já existe no destino (%s)", uid_s, mid)
                elif dry_run:
                    detalhe["acao"] = "DRY_RUN_COPIARIA"
                    logging.info(
                        "DRY-RUN UID %s copiaria | %s | %s",
                        uid_s,
                        mid or "(sem Message-ID)",
                        assunto[:60],
                    )
                else:
                    append_mensagem(kh, pasta_destino, raw, flags, internaldate)
                    detalhe["acao"] = "COPIADA"
                    resultado["copiadas"] += 1
                    if mid:
                        indice_dest.add(mid.strip().lower())
                        if mid.startswith("<") and mid.endswith(">"):
                            indice_dest.add(mid[1:-1].strip().lower())
                    logging.info("UID %s COPIADA | %s", uid_s, assunto[:60])
            except Exception as exc:
                detalhe["acao"] = "ERRO"
                detalhe["erro"] = str(exc)
                resultado["erros"] += 1
                logging.error("UID %s erro: %s", uid_s, exc)

            resultado["detalhes"].append(detalhe)

        q_kh2, _ = status_mensagens(kh, pasta_destino)
        resultado["inbox_depois_kh"] = q_kh2 if q_kh2 is not None else ""
        resultado["ok"] = resultado["erros"] == 0
        logging.info(
            "Resultado %s | copiadas=%s | ja_existiam=%s | erros=%s | King depois=%s",
            email_addr,
            resultado["copiadas"],
            resultado["ja_existiam"],
            resultado["erros"],
            resultado["inbox_depois_kh"],
        )
        return resultado
    except Exception as exc:
        resultado["erro"] = str(exc)
        logging.exception("Falha na conta %s: %s", email_addr, exc)
        return resultado
    finally:
        for client in (lw, kh):
            if client is None:
                continue
            try:
                client.logout()
            except Exception:
                pass


def carregar_contas_args(args: argparse.Namespace) -> List[Dict[str, str]]:
    """Prioridade: senhas CLI → --contas CSV → planilha alinhamento → .env unica."""

    if args.senha_locaweb and args.senha_king and args.email:
        emails = args.email if isinstance(args.email, list) else [args.email]
        return [
            {
                "email": e.strip().lower(),
                "senha_locaweb": args.senha_locaweb,
                "senha_kinghost": args.senha_king,
            }
            for e in emails
        ]

    if args.contas:
        path = Path(args.contas)
        if not path.exists():
            raise SystemExit(f"Arquivo de contas não encontrado: {path}")
        df = pd.read_csv(path)
        cols = {c.lower().strip(): c for c in df.columns}
        need = ["email", "senha_locaweb", "senha_kinghost"]
        for n in need:
            if n not in cols:
                raise SystemExit(f"CSV precisa das colunas: {need}. Achou: {list(df.columns)}")
        rows = []
        for _, r in df.iterrows():
            email_addr = str(r[cols["email"]]).strip().lower()
            if not email_addr or "@" not in email_addr:
                continue
            rows.append(
                {
                    "email": email_addr,
                    "senha_locaweb": str(r[cols["senha_locaweb"]]).strip(),
                    "senha_kinghost": str(r[cols["senha_kinghost"]]).strip(),
                }
            )
        if args.email:
            filtro = {e.strip().lower() for e in args.email}
            rows = [r for r in rows if r["email"] in filtro]
        if not rows:
            raise SystemExit("Nenhuma conta no CSV (após filtro).")
        return rows

    # planilha alinhamento
    alinhamento = Path(args.planilha or DEFAULT_ALINHAMENTO)
    if alinhamento.exists() and args.email:
        df = carregar_alinhamento(alinhamento)
        contas = []
        for e in args.email:
            filtradas = filtrar_contas(
                df,
                apenas_nos_dois=False,
                somente_com_senhas=True,
                email=e.strip().lower(),
            )
            contas.extend(filtradas)
        if contas:
            return [
                {
                    "email": c["email"],
                    "senha_locaweb": c["senha_locaweb"],
                    "senha_kinghost": c["senha_kinghost"],
                }
                for c in contas
            ]

    # .env unica
    load_dotenv(PASTA / ".env")
    email_addr = os.getenv("LOCALWEB_EMAIL", "").strip().lower()
    if email_addr and os.getenv("LOCALWEB_PASSWORD") and os.getenv("KINGHOST_PASSWORD"):
        return [
            {
                "email": email_addr,
                "senha_locaweb": os.getenv("LOCALWEB_PASSWORD", ""),
                "senha_kinghost": os.getenv("KINGHOST_PASSWORD", ""),
            }
        ]

    raise SystemExit(
        "Informe contas via --contas contas_teste.csv ou "
        "--email X --senha-locaweb Y --senha-king Z"
    )


def salvar_saidas(resultados: List[Dict[str, Any]]) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx = PASTA_SAIDAS / f"migracao_teste_{stamp}.xlsx"
    resumo_rows = []
    detalhe_rows = []
    for res in resultados:
        resumo_rows.append(
            {
                "email": res["email"],
                "ok": "SIM" if res["ok"] else "NÃO",
                "dry_run": "SIM" if res["dry_run"] else "NÃO",
                "limite": res["limite"],
                "inbox_antes_locaweb": res["inbox_antes_lw"],
                "inbox_antes_kinghost": res["inbox_antes_kh"],
                "inbox_depois_kinghost": res["inbox_depois_kh"],
                "copiadas": res["copiadas"],
                "ja_existiam": res["ja_existiam"],
                "erros": res["erros"],
                "erro": res.get("erro", ""),
            }
        )
        detalhe_rows.extend(res.get("detalhes") or [])

    with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
        pd.DataFrame(resumo_rows).to_excel(w, sheet_name="resumo", index=False)
        pd.DataFrame(detalhe_rows if detalhe_rows else [{"info": "sem detalhes"}]).to_excel(
            w, sheet_name="detalhes", index=False
        )

    csv_path = PASTA_SAIDAS / f"migracao_teste_resumo_{stamp}.csv"
    pd.DataFrame(resumo_rows).to_csv(csv_path, index=False)

    if USB_OUTPUT_DIR:
        usb = Path(USB_OUTPUT_DIR)
        if usb.exists():
            try:
                usb_xlsx = usb / xlsx.name
                with pd.ExcelWriter(usb_xlsx, engine="openpyxl") as w:
                    pd.DataFrame(resumo_rows).to_excel(w, sheet_name="resumo", index=False)
                    pd.DataFrame(
                        detalhe_rows if detalhe_rows else [{"info": "sem detalhes"}]
                    ).to_excel(w, sheet_name="detalhes", index=False)
                logging.info("Cópia USB: %s", usb_xlsx)
            except Exception as exc:
                logging.warning("Não gravou no pendrive: %s", exc)

    return xlsx


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Migração TESTE IMAP Localweb → KingHost (só INBOX, poucas msgs)."
    )
    p.add_argument(
        "--contas",
        default=str(DEFAULT_CONTAS_TESTE) if DEFAULT_CONTAS_TESTE.exists() else None,
        help="CSV com email,senha_locaweb,senha_kinghost",
    )
    p.add_argument(
        "--email",
        action="append",
        default=None,
        help="Filtra e-mail (pode repetir). Ou usa com --senha-*",
    )
    p.add_argument("--senha-locaweb", default=None)
    p.add_argument("--senha-king", default=None)
    p.add_argument("--planilha", default=None, help="alinhamento_contas_*.xlsx")
    p.add_argument(
        "--limite",
        type=int,
        default=int(os.getenv("MIGRATE_LIMIT", "10")),
        help="Qtd máx. de mensagens (mais recentes). Default 10",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Só simula: não grava no destino",
    )
    p.add_argument("--pasta-origem", default="INBOX")
    p.add_argument("--pasta-destino", default="INBOX")
    return p


def main() -> None:
    load_dotenv(PASTA / ".env")
    args = build_parser().parse_args()
    log_path = configurar_logging()
    hosts = hosts_from_env()

    contas = carregar_contas_args(args)
    logging.info(
        "Migração TESTE | contas=%d | limite=%d | dry_run=%s",
        len(contas),
        args.limite,
        args.dry_run,
    )

    resultados = []
    for i, conta in enumerate(contas, 1):
        logging.info("==== [%d/%d] %s ====", i, len(contas), conta["email"])
        cfg = cfg_par(
            conta["email"], conta["senha_locaweb"], conta["senha_kinghost"], hosts
        )
        res = migrar_inbox_conta(
            cfg,
            limite=args.limite,
            dry_run=args.dry_run,
            pasta_origem=args.pasta_origem,
            pasta_destino=args.pasta_destino,
        )
        resultados.append(res)

    xlsx = salvar_saidas(resultados)
    print("\n=== Resumo migração teste ===")
    for res in resultados:
        print(
            f"  {res['email']}: "
            f"LW {res['inbox_antes_lw']} → KH {res['inbox_antes_kh']}..{res['inbox_depois_kh']} | "
            f"copiadas={res['copiadas']} ja_existiam={res['ja_existiam']} erros={res['erros']}"
            + (f" | ERRO: {res['erro']}" if res.get("erro") else "")
        )
    print(f"\nPlanilha: {xlsx}")
    print(f"Log:      {log_path}")


if __name__ == "__main__":
    main()
