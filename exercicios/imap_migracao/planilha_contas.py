#!/usr/bin/env python3
"""Carrega contas a partir das planilhas de alinhamento / senhas."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

# Domínio padrão genérico — sobrescreva com EMAIL_DOMAIN no .env
DOMAIN = (os.getenv("EMAIL_DOMAIN") or "exemplo.com.br").strip() or "exemplo.com.br"


def _txt(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def _sim(v: Any) -> bool:
    return _txt(v).upper() in {"SIM", "S", "YES", "TRUE", "1", "X"}


def _norm_local(s: str) -> str:
    s = s.strip().lower().replace(" ", "")
    if "@" in s:
        return s.split("@")[0]
    return s


def _to_email(local_or_email: str) -> str:
    s = local_or_email.strip().lower()
    if "@" in s:
        return s
    return f"{s}@{DOMAIN}"


def carregar_alinhamento(path: Path) -> pd.DataFrame:
    """Lê alinhamento_contas_king_locaweb.xlsx (ou CSV equivalente)."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {path}")

    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)

    colmap = {c.lower().strip(): c for c in df.columns}

    def col(*names: str) -> Optional[str]:
        for n in names:
            if n in colmap:
                return colmap[n]
        return None

    c_email = col("email")
    c_sl = col("senha_locaweb")
    c_sk = col("senha_kinghost")
    if not c_email or not c_sl or not c_sk:
        raise ValueError(
            "Planilha precisa das colunas: email, senha_locaweb, senha_kinghost. "
            f"Encontradas: {list(df.columns)}"
        )

    c_nl = col("na_locaweb")
    c_nk = col("na_kinghost")
    c_mig = col("migrado")
    c_sit = col("situacao")

    rows = []
    for _, r in df.iterrows():
        email = _txt(r[c_email]).lower()
        if not email or "@" not in email:
            continue
        rows.append(
            {
                "email": email,
                "senha_locaweb": _txt(r[c_sl]),
                "senha_kinghost": _txt(r[c_sk]),
                "na_locaweb": _sim(r[c_nl]) if c_nl else bool(_txt(r[c_sl])),
                "na_kinghost": _sim(r[c_nk]) if c_nk else bool(_txt(r[c_sk])),
                "migrado": _sim(r[c_mig]) if c_mig else False,
                "situacao": _txt(r[c_sit]) if c_sit else "",
            }
        )
    return pd.DataFrame(rows)


def reconstruir_de_fontes(
    path_king: Path,
    path_senhas: Path,
) -> pd.DataFrame:
    """
    Monta alinhamento a partir de:
      - listagemsatoKingHOST.xlsx (senhas KingHost)
      - NOVASSENHAS.xlsx (senhas Localweb)
    """
    kh = pd.read_excel(path_king, sheet_name="caixas_postais(1).csv - Cop (2", header=None)
    king = {}
    for _, row in kh.iterrows():
        email = _txt(row[5]).lower()
        if "@" not in email or email.startswith("emai"):
            continue
        pwd = _txt(row[6])
        king[email] = pwd

    ns = pd.read_excel(path_senhas, sheet_name="2023", header=None)
    loc = {}
    for _, row in ns.iterrows():
        raw = _txt(row[5])
        if not raw or raw.lower() == "email" or "senha:" in raw.lower():
            continue
        if raw.lower().startswith("webmail") or raw.lower().startswith("locaweb"):
            continue
        pwd = _txt(row[8])
        if not pwd:
            continue
        local = _norm_local(raw)
        if not local or local == "registro.br":
            continue
        loc[_to_email(local)] = pwd

    emails = sorted(set(king) | set(loc))
    rows = []
    for email in emails:
        sl = loc.get(email, "")
        sk = king.get(email, "")
        rows.append(
            {
                "email": email,
                "senha_locaweb": sl,
                "senha_kinghost": sk,
                "na_locaweb": bool(sl),
                "na_kinghost": bool(sk),
                "migrado": False,
                "situacao": (
                    "NOS_DOIS"
                    if sl and sk
                    else ("SO_LOCAWEB" if sl else "SO_KINGHOST")
                ),
            }
        )
    return pd.DataFrame(rows)


def filtrar_contas(
    df: pd.DataFrame,
    *,
    apenas_nos_dois: bool = True,
    somente_com_senhas: bool = True,
    pular_migrados: bool = False,
    email: Optional[str] = None,
    limite: Optional[int] = None,
) -> List[Dict[str, Any]]:
    out = df.copy()

    if email:
        alvo = email.strip().lower()
        out = out[out["email"] == alvo]
        if out.empty:
            raise SystemExit(f"E-mail não encontrado na planilha: {alvo}")

    if apenas_nos_dois:
        out = out[out["na_locaweb"] & out["na_kinghost"]]

    if somente_com_senhas:
        out = out[
            (out["senha_locaweb"].astype(str).str.len() > 0)
            & (out["senha_kinghost"].astype(str).str.len() > 0)
        ]

    if pular_migrados:
        out = out[~out["migrado"].astype(bool)]

    out = out.sort_values("email").reset_index(drop=True)
    if limite is not None and limite > 0:
        out = out.head(limite)

    return out.to_dict("records")
