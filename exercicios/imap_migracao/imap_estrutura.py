#!/usr/bin/env python3
"""Núcleo IMAP: estrutura, contagens e mapeamento Localweb × KingHost/Titan."""

from __future__ import annotations

import base64
import imaplib
import logging
import re
import ssl
from typing import Any, Dict, List, Optional, Tuple

STATUS_OK = "OK_PROVAVELMENTE_MIGRADA"
STATUS_FALTA = "FALTA_MIGRAR_OU_PARCIAL"
STATUS_SO_ORIGEM = "SO_ORIGEM"
STATUS_SO_DESTINO = "SO_DESTINO"
STATUS_AMBOS_VAZIOS = "AMBOS_VAZIOS"
STATUS_ERRO = "ERRO"
STATUS_SEM_PAR = "SEM_PAR_NO_DESTINO"
STATUS_SO_DESTINO_EXTRA = "SO_DESTINO_SEM_ORIGEM"
STATUS_LOGIN_FALHOU = "LOGIN_FALHOU"

ALIASES_PAPEL = {
    "INBOX": {"inbox"},
    "Sent": {
        "sent",
        "sent messages",
        "sent items",
        "enviadas",
        "itens enviados",
        "inbox.enviadas",
        "inbox.sent",
    },
    "Drafts": {
        "drafts",
        "draft",
        "rascunho",
        "rascunhos",
        "inbox.rascunho",
        "inbox.rascunhos",
        "inbox.drafts",
    },
    "Trash": {
        "trash",
        "deleted",
        "deleted messages",
        "deleted items",
        "lixo",
        "lixeira",
        "inbox.lixo",
        "inbox.lixeira",
        "inbox.trash",
    },
    "Spam": {
        "spam",
        "junk",
        "junk e-mail",
        "junk email",
        "bulk mail",
        "mala_direta",
        "mala direta",
        "inbox.mala_direta",
        "inbox.spam",
        "inbox.junk",
    },
    "Archive": {"archive", "arquivados", "arquivo", "inbox.archive", "inbox.arquivo"},
    "Scheduled": {"scheduled", "agendados", "inbox.scheduled"},
}

FLAG_PARA_PAPEL = {
    "\\inbox": "INBOX",
    "\\sent": "Sent",
    "\\drafts": "Drafts",
    "\\trash": "Trash",
    "\\junk": "Spam",
    "\\archive": "Archive",
}


def conectar_imap(nome: str, cfg: dict) -> imaplib.IMAP4_SSL:
    logging.info(
        "Conectando %s → %s:%s como %s",
        nome,
        cfg["host"],
        cfg["port"],
        cfg["email"],
    )
    ctx = ssl.create_default_context()
    client = imaplib.IMAP4_SSL(cfg["host"], cfg["port"], ssl_context=ctx)
    client.login(cfg["email"], cfg["password"])
    logging.info("%s: login OK", nome)
    return client


def _decode_imap_utf7(name: str) -> str:
    if "&" not in name:
        return name
    try:

        def repl(match: re.Match) -> str:
            chunk = match.group(1)
            if chunk == "":
                return "&"
            pad = (-len(chunk)) % 4
            raw = base64.b64decode(chunk.replace(",", "/") + ("=" * pad))
            return raw.decode("utf-16-be")

        return re.sub(r"&([^-]*)-", repl, name)
    except Exception:
        return name


def _encode_imap_utf7(name: str) -> str:
    """Codifica nome de pasta para Modified UTF-7 (ASCII seguro no imaplib)."""
    if not name:
        return name
    # Já parece Modified UTF-7 e é ASCII puro
    try:
        name.encode("ascii")
        return name
    except UnicodeEncodeError:
        pass

    out: List[str] = []
    i = 0
    n = len(name)
    while i < n:
        ch = name[i]
        if ch == "&":
            out.append("&-")
            i += 1
            continue
        if ord(ch) < 0x80:
            j = i
            while j < n and ord(name[j]) < 0x80 and name[j] != "&":
                j += 1
            out.append(name[i:j])
            i = j
            continue
        j = i
        while j < n and ord(name[j]) >= 0x80:
            j += 1
        raw = name[i:j].encode("utf-16-be")
        enc = base64.b64encode(raw).decode("ascii").rstrip("=").replace("/", ",")
        out.append(f"&{enc}-")
        i = j
    return "".join(out)


def _refs_mailbox(pasta: str, pasta_imap: Optional[str] = None) -> List[str]:
    """Candidatos ASCII-safe para STATUS/SELECT."""
    wire = pasta_imap or _encode_imap_utf7(pasta)
    try:
        wire.encode("ascii")
    except UnicodeEncodeError:
        wire = _encode_imap_utf7(pasta)
    refs = []
    for cand in (wire, pasta if pasta != wire else None):
        if not cand:
            continue
        try:
            cand.encode("ascii")
        except UnicodeEncodeError:
            continue
        refs.append(f'"{cand}"')
        refs.append(cand)
    # remove duplicados preservando ordem
    vistos = set()
    unicos = []
    for r in refs:
        if r not in vistos:
            vistos.add(r)
            unicos.append(r)
    return unicos


def parse_list_line(raw: Any) -> Optional[Dict[str, Any]]:
    if raw is None:
        return None
    if isinstance(raw, tuple):
        raw = raw[0]
    if isinstance(raw, bytes):
        text = raw.decode("utf-8", errors="replace")
    else:
        text = str(raw)

    text = text.strip()
    if text.upper().startswith("LIST "):
        text = text[5:].strip()
    if text.upper().startswith("LSUBSCRIBE "):
        text = text[11:].strip()

    m = re.match(
        r'^\((?P<flags>.*)\)\s+(?P<delim>"."|NIL)\s+(?P<name>.+)$',
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if not m:
        logging.debug("Linha LIST não reconhecida: %r", text)
        return None

    flags_raw = m.group("flags").strip()
    flags = [f for f in re.findall(r'\\[A-Za-z0-9_-]+|"[^"]+"|\S+', flags_raw) if f]
    flags = [f.strip('"') for f in flags]

    delim_token = m.group("delim")
    delimiter = None if delim_token.upper() == "NIL" else delim_token.strip('"')

    name_token = m.group("name").strip()
    if name_token.startswith('"') and name_token.endswith('"'):
        nome_wire = name_token[1:-1].replace('\\"', '"')
    else:
        nome_wire = name_token

    nome = _decode_imap_utf7(nome_wire)
    # nome_imap = forma para comandos (ASCII / Modified UTF-7)
    try:
        nome_wire.encode("ascii")
        nome_imap = nome_wire
    except UnicodeEncodeError:
        nome_imap = _encode_imap_utf7(nome)

    flags_lower = {f.lower() for f in flags}
    return {
        "nome": nome,
        "nome_imap": nome_imap,
        "delimiter": delimiter,
        "flags": flags,
        "selecionavel": "\\noselect" not in flags_lower,
        "tem_filhos": "\\haschildren" in flags_lower,
        "folha": "\\hasnochildren" in flags_lower,
    }


def papel_por_flags(flags: List[str]) -> Optional[str]:
    for flag in flags:
        papel = FLAG_PARA_PAPEL.get(flag.lower())
        if papel:
            return papel
    return None


def papel_por_nome(nome: str, delimiter: Optional[str]) -> Optional[str]:
    n = nome.strip().lower().replace("\\", "/")
    if delimiter and delimiter != "/":
        n = n.replace(delimiter, ".")
    n = re.sub(r"\s+", " ", n)
    if n == "inbox":
        return "INBOX"
    for papel, aliases in ALIASES_PAPEL.items():
        if n in aliases:
            return papel
        base = n.split(".")[-1].split("/")[-1]
        if base in aliases:
            return papel
    return None


def inferir_papel(pasta: Dict[str, Any]) -> str:
    return (
        papel_por_flags(pasta["flags"])
        or papel_por_nome(pasta["nome"], pasta.get("delimiter"))
        or "CUSTOM"
    )


def caminho_partes(nome: str, delimiter: Optional[str]) -> List[str]:
    if not delimiter:
        return [nome]
    return [p for p in nome.split(delimiter) if p != ""]


def listar_estrutura(
    client: imaplib.IMAP4_SSL, provedor: str, verbose: bool = True
) -> List[Dict[str, Any]]:
    pastas: List[Dict[str, Any]] = []
    vistos = set()

    for comando in ("list", "lsub"):
        try:
            status, data = getattr(client, comando)()
        except imaplib.IMAP4.error as exc:
            logging.warning("%s: %s falhou: %s", provedor, comando.upper(), exc)
            continue
        if status != "OK" or not data:
            continue
        for item in data:
            parsed = parse_list_line(item)
            if not parsed:
                continue
            chave = parsed["nome"]
            if chave in vistos:
                for p in pastas:
                    if p["nome"] == chave:
                        p["flags"] = list(dict.fromkeys(p["flags"] + parsed["flags"]))
                        break
                continue
            vistos.add(chave)
            parsed["provedor"] = provedor
            parsed["papel"] = inferir_papel(parsed)
            parsed["nivel"] = len(caminho_partes(parsed["nome"], parsed["delimiter"]))
            parsed["origem_lista"] = comando.upper()
            pastas.append(parsed)

    delims = [p["delimiter"] for p in pastas if p.get("delimiter")]
    delim_comum = max(set(delims), key=delims.count) if delims else None
    for p in pastas:
        if not p.get("delimiter"):
            p["delimiter"] = delim_comum

    pastas.sort(key=lambda p: (p["papel"] != "INBOX", p["nome"].lower()))
    logging.info("%s: %d pastas | delimitador=%r", provedor, len(pastas), delim_comum)
    if verbose:
        for p in pastas:
            logging.info(
                "  [%s] %-28s | papel=%-8s | flags=%s | sel=%s",
                provedor,
                p["nome"],
                p["papel"],
                " ".join(p["flags"]) or "-",
                "sim" if p["selecionavel"] else "não",
            )
    return pastas


def status_mensagens(
    client: imaplib.IMAP4_SSL,
    pasta: str,
    pasta_imap: Optional[str] = None,
) -> Tuple[Optional[int], Optional[int]]:
    candidatos = _refs_mailbox(pasta, pasta_imap)
    for ref in candidatos:
        try:
            st, data = client.status(ref, "(MESSAGES UNSEEN)")
            if st != "OK" or not data or data[0] is None:
                continue
            raw = data[0]
            text = raw.decode("utf-8", errors="replace") if isinstance(raw, bytes) else str(raw)
            msgs = re.search(r"MESSAGES\s+(\d+)", text, re.I)
            unseen = re.search(r"UNSEEN\s+(\d+)", text, re.I)
            if msgs:
                return int(msgs.group(1)), int(unseen.group(1)) if unseen else 0
        except (imaplib.IMAP4.error, UnicodeEncodeError, UnicodeDecodeError):
            continue

    for ref in candidatos:
        try:
            st, _ = client.select(ref, readonly=True)
            if st != "OK":
                continue
            st, data = client.search(None, "ALL")
            if st != "OK" or data[0] is None:
                return 0, 0
            total = len(data[0].split())
            st, data_u = client.search(None, "UNSEEN")
            nao_lidas = len(data_u[0].split()) if st == "OK" and data_u[0] else 0
            return total, nao_lidas
        except (imaplib.IMAP4.error, UnicodeEncodeError, UnicodeDecodeError) as exc:
            logging.debug("SELECT falhou '%s' ref=%r: %s", pasta, ref, exc)
            continue

    logging.warning("Falha ao ler contagem de '%s'", pasta)
    return None, None


def enriquecer_contagens(
    client: imaplib.IMAP4_SSL,
    pastas: List[Dict[str, Any]],
    provedor: str,
    verbose: bool = True,
) -> None:
    for p in pastas:
        if not p["selecionavel"]:
            p["mensagens"] = None
            p["nao_lidas"] = None
            p["erro_leitura"] = "Noselect"
            continue
        try:
            total, unseen = status_mensagens(
                client, p["nome"], p.get("nome_imap")
            )
            p["mensagens"] = total
            p["nao_lidas"] = unseen
            p["erro_leitura"] = "" if total is not None else "STATUS/SELECT falhou"
        except Exception as exc:
            p["mensagens"] = None
            p["nao_lidas"] = None
            p["erro_leitura"] = str(exc)
            logging.warning(
                "[%s] erro ao contar '%s': %s", provedor, p["nome"], exc
            )
        if verbose:
            logging.info(
                "  [%s] %-28s | msgs=%s | unseen=%s",
                provedor,
                p["nome"],
                p["mensagens"] if p.get("mensagens") is not None else "?",
                p["nao_lidas"] if p.get("nao_lidas") is not None else "?",
            )


def sugerir_destino_titan(
    origem: Dict[str, Any],
    destino_por_papel: Dict[str, Dict[str, Any]],
    destinos: List[Dict[str, Any]],
) -> Tuple[Optional[str], str]:
    papel = origem["papel"]
    if papel != "CUSTOM" and papel in destino_por_papel:
        return destino_por_papel[papel]["nome"], f"papel={papel}"

    for d in destinos:
        if d["nome"].lower() == origem["nome"].lower():
            return d["nome"], "nome_igual"

    folha = origem["nome"].split(origem.get("delimiter") or ".")[-1].lower()
    for d in destinos:
        folha_d = d["nome"].split(d.get("delimiter") or "/")[-1].lower()
        if folha == folha_d:
            return d["nome"], "folha_igual"

    if papel == "CUSTOM":
        delim_src = origem.get("delimiter") or "."
        partes = [p for p in origem["nome"].split(delim_src) if p.upper() != "INBOX"]
        sugerido = "/".join(partes) if partes else origem["nome"]
        return sugerido, "criar_no_destino"

    return None, "sem_mapeamento"


def classificar(qtd_origem: Optional[int], qtd_destino: Optional[int]) -> str:
    if qtd_origem is None or qtd_destino is None:
        return STATUS_ERRO
    if qtd_origem == 0 and qtd_destino == 0:
        return STATUS_AMBOS_VAZIOS
    if qtd_origem > 0 and qtd_destino == 0:
        return STATUS_SO_ORIGEM
    if qtd_origem == 0 and qtd_destino > 0:
        return STATUS_SO_DESTINO
    diff = abs(qtd_origem - qtd_destino)
    limite = max(2, int(max(qtd_origem, qtd_destino) * 0.05))
    if diff <= limite:
        return STATUS_OK
    return STATUS_FALTA


def montar_mapeamento(
    origem: List[Dict[str, Any]], destino: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    dest_por_papel: Dict[str, Dict[str, Any]] = {}
    for d in destino:
        if d["papel"] != "CUSTOM" and d["papel"] not in dest_por_papel:
            dest_por_papel[d["papel"]] = d

    usados_destino = set()
    linhas: List[Dict[str, Any]] = []

    for o in origem:
        if not o["selecionavel"]:
            continue
        dest_nome, motivo = sugerir_destino_titan(o, dest_por_papel, destino)
        dest_obj = next((d for d in destino if d["nome"] == dest_nome), None)
        if dest_obj:
            usados_destino.add(dest_obj["nome"])

        q_o = o.get("mensagens")
        q_d = dest_obj.get("mensagens") if dest_obj else None
        if dest_obj is None:
            status = STATUS_SO_ORIGEM if motivo == "criar_no_destino" else STATUS_SEM_PAR
        else:
            status = classificar(q_o, q_d)

        linhas.append(
            {
                "papel": o["papel"],
                "pasta_locaweb": o["nome"],
                "delim_locaweb": o.get("delimiter") or "",
                "msgs_locaweb": q_o if q_o is not None else "",
                "unseen_locaweb": o.get("nao_lidas") if o.get("nao_lidas") is not None else "",
                "pasta_kinghost": dest_nome or "",
                "delim_kinghost": (dest_obj or {}).get("delimiter") or "",
                "msgs_kinghost": q_d if q_d is not None else "",
                "unseen_kinghost": (
                    dest_obj.get("nao_lidas")
                    if dest_obj and dest_obj.get("nao_lidas") is not None
                    else ""
                ),
                "destino_existe": "SIM" if dest_obj else "NÃO",
                "regra_mapeamento": motivo,
                "diferenca": "" if q_o is None or q_d is None else (q_d - q_o),
                "status": status,
            }
        )

    for d in destino:
        if not d["selecionavel"] or d["nome"] in usados_destino:
            continue
        linhas.append(
            {
                "papel": d["papel"],
                "pasta_locaweb": "",
                "delim_locaweb": "",
                "msgs_locaweb": "",
                "unseen_locaweb": "",
                "pasta_kinghost": d["nome"],
                "delim_kinghost": d.get("delimiter") or "",
                "msgs_kinghost": d["mensagens"] if d.get("mensagens") is not None else "",
                "unseen_kinghost": d["nao_lidas"] if d.get("nao_lidas") is not None else "",
                "destino_existe": "SIM",
                "regra_mapeamento": "somente_destino",
                "diferenca": "",
                "status": STATUS_SO_DESTINO_EXTRA,
            }
        )

    return linhas


def filtrar_por_modo(mapeamento: List[Dict[str, Any]], modo: str) -> List[Dict[str, Any]]:
    if modo == "ALL":
        return mapeamento
    return [
        r
        for r in mapeamento
        if r["papel"] == "INBOX"
        or (r["pasta_locaweb"] or "").upper() == "INBOX"
        or (r["pasta_kinghost"] or "").upper() == "INBOX"
    ]


def estrutura_para_rows(pastas: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rows = []
    for p in pastas:
        rows.append(
            {
                "provedor": p["provedor"],
                "nome": p["nome"],
                "papel": p["papel"],
                "delimiter": p.get("delimiter") or "",
                "nivel": p.get("nivel") or "",
                "selecionavel": "SIM" if p["selecionavel"] else "NÃO",
                "flags": " ".join(p.get("flags") or []),
                "mensagens": p["mensagens"] if p.get("mensagens") is not None else "",
                "nao_lidas": p["nao_lidas"] if p.get("nao_lidas") is not None else "",
                "erro_leitura": p.get("erro_leitura") or "",
            }
        )
    return rows


def analisar_conta(cfg: dict, verbose: bool = True) -> Dict[str, Any]:
    """
    Analisa uma conta (Localweb × KingHost).
    cfg precisa de localweb/kinghost (email,password,host,port) e modo.
    """
    resultado: Dict[str, Any] = {
        "email": cfg["localweb"]["email"],
        "ok": False,
        "erro": "",
        "login_locaweb": "NÃO",
        "login_kinghost": "NÃO",
        "origem": [],
        "destino": [],
        "mapeamento": [],
        "inbox_locaweb": "",
        "inbox_kinghost": "",
        "status_inbox": STATUS_ERRO,
        "status_geral": STATUS_ERRO,
    }

    lw = kh = None
    try:
        try:
            lw = conectar_imap("Localweb", cfg["localweb"])
            resultado["login_locaweb"] = "SIM"
        except Exception as exc:
            resultado["erro"] = f"Localweb: {exc}"
            resultado["status_geral"] = STATUS_LOGIN_FALHOU
            logging.error("Login Localweb falhou (%s): %s", cfg["localweb"]["email"], exc)
            return resultado

        try:
            kh = conectar_imap("KingHost", cfg["kinghost"])
            resultado["login_kinghost"] = "SIM"
        except Exception as exc:
            resultado["erro"] = f"KingHost: {exc}"
            resultado["status_geral"] = STATUS_LOGIN_FALHOU
            logging.error("Login KingHost falhou (%s): %s", cfg["kinghost"]["email"], exc)
            return resultado

        try:
            origem = listar_estrutura(lw, "Localweb", verbose=verbose)
            destino = listar_estrutura(kh, "KingHost", verbose=verbose)
            if verbose:
                logging.info("Contando mensagens (STATUS)…")
            enriquecer_contagens(lw, origem, "Localweb", verbose=verbose)
            enriquecer_contagens(kh, destino, "KingHost", verbose=verbose)

            mapeamento = filtrar_por_modo(
                montar_mapeamento(origem, destino), cfg["modo"]
            )
            resultado["origem"] = origem
            resultado["destino"] = destino
            resultado["mapeamento"] = mapeamento
            resultado["ok"] = True

            inbox = next((r for r in mapeamento if r["papel"] == "INBOX"), None)
            if inbox:
                resultado["inbox_locaweb"] = inbox["msgs_locaweb"]
                resultado["inbox_kinghost"] = inbox["msgs_kinghost"]
                resultado["status_inbox"] = inbox["status"]

            ordem = [
                STATUS_LOGIN_FALHOU,
                STATUS_ERRO,
                STATUS_SEM_PAR,
                STATUS_SO_ORIGEM,
                STATUS_FALTA,
                STATUS_SO_DESTINO,
                STATUS_SO_DESTINO_EXTRA,
                STATUS_AMBOS_VAZIOS,
                STATUS_OK,
            ]
            relevantes = [
                r["status"]
                for r in mapeamento
                if r.get("pasta_locaweb")
                and r.get("msgs_locaweb") != ""
                and r.get("msgs_locaweb") != 0
            ] or [r["status"] for r in mapeamento if r.get("pasta_locaweb")]
            if not relevantes and inbox:
                relevantes = [inbox["status"]]
            resultado["status_geral"] = sorted(
                relevantes or [STATUS_AMBOS_VAZIOS],
                key=lambda s: ordem.index(s) if s in ordem else 99,
            )[0]
        except Exception as exc:
            resultado["ok"] = False
            resultado["erro"] = str(exc)
            resultado["status_geral"] = STATUS_ERRO
            logging.exception(
                "Erro ao analisar %s: %s", cfg["localweb"]["email"], exc
            )
        return resultado
    finally:
        for client in (lw, kh):
            if client is None:
                continue
            try:
                client.logout()
            except Exception:
                pass
