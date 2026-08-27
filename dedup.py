"""Deduplica delle aziende contro lo storico persistente (CSV).

Lo storico e' un semplice CSV con due colonne: ``nome_normalizzato`` e
``dominio``. Il confronto si fa su una versione normalizzata del nome
(lowercase, senza suffissi societari, senza punteggiatura) e sul dominio
principale del sito web (senza ``www.`` e senza path).
"""

from __future__ import annotations

import csv
import logging
import re
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlparse


logger = logging.getLogger(__name__)


# Suffissi societari italiani/svizzeri comuni da rimuovere nella normalizzazione.
_SUFFIX_PATTERN = re.compile(
    r"\b("
    r"s\.?r\.?l\.?s?|"
    r"s\.?p\.?a\.?|"
    r"s\.?n\.?c\.?|"
    r"s\.?a\.?s\.?|"
    r"s\.?s\.?|"
    r"s\.?a\.?|"
    r"srl|srls|spa|snc|sas|"
    r"soc\.?\s*coop\.?|"
    r"societa'?\s*cooperativa|"
    r"cooperativa|"
    r"gmbh|ag|"
    r"& c\.?|"
    r"e\s+c\.?"
    r")\b",
    re.IGNORECASE,
)

_PUNCT_PATTERN = re.compile(r"[^\w\s]", re.UNICODE)
_WS_PATTERN = re.compile(r"\s+")


def normalize_name(name: str) -> str:
    """Normalizza il nome azienda per il confronto.

    - lowercase
    - rimuove suffissi societari (S.r.l., S.p.A., SNC, GmbH, AG, ecc.)
    - rimuove punteggiatura
    - collassa spazi multipli
    """

    if not name:
        return ""
    n = name.lower().strip()
    n = _SUFFIX_PATTERN.sub(" ", n)
    n = _PUNCT_PATTERN.sub(" ", n)
    n = _WS_PATTERN.sub(" ", n).strip()
    return n


def normalize_domain(url: str) -> str:
    """Estrae il dominio principale da un URL, in lowercase, senza ``www.``.

    Ritorna stringa vuota se l'URL e' vuoto o non parsabile.
    """

    if not url:
        return ""
    candidate = url.strip()
    if not candidate:
        return ""
    if "://" not in candidate:
        candidate = "http://" + candidate
    try:
        parsed = urlparse(candidate)
    except ValueError:
        return ""
    host = (parsed.hostname or "").lower().strip()
    if host.startswith("www."):
        host = host[4:]
    return host


def load_history(path: Path) -> tuple[set[str], set[str]]:
    """Carica lo storico e ritorna (nomi_normalizzati, domini)."""

    names: set[str] = set()
    domains: set[str] = set()

    if not path.exists():
        logger.info("Storico non trovato (%s), parto da zero", path)
        return names, domains

    try:
        with path.open("r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                nome_norm = (row.get("nome_normalizzato") or "").strip()
                dominio = (row.get("dominio") or "").strip().lower()
                if nome_norm:
                    names.add(nome_norm)
                if dominio:
                    domains.add(dominio)
    except OSError as exc:
        logger.warning("Impossibile leggere lo storico %s: %s", path, exc)

    logger.info(
        "Storico caricato: %d nomi, %d domini", len(names), len(domains)
    )
    return names, domains


def filter_new(
    rows: Iterable[dict[str, Any]],
    known_names: set[str],
    known_domains: set[str],
) -> tuple[list[dict[str, Any]], int]:
    """Filtra le righe gia' note.

    Deduplica anche INTRA-batch: se la stessa azienda compare due volte nella
    stessa risposta di Claude, tiene solo la prima occorrenza.

    Ritorna (nuove_righe, numero_duplicati_scartati).
    """

    new_rows: list[dict[str, Any]] = []
    seen_names: set[str] = set(known_names)
    seen_domains: set[str] = set(known_domains)
    duplicates = 0

    for row in rows:
        nome_norm = normalize_name(row.get("nome_azienda", ""))
        dominio = normalize_domain(row.get("sito_web", ""))

        is_dup = False
        if nome_norm and nome_norm in seen_names:
            is_dup = True
        if dominio and dominio in seen_domains:
            is_dup = True

        if is_dup:
            duplicates += 1
            logger.debug(
                "Duplicato scartato: %s (%s)",
                row.get("nome_azienda", ""),
                dominio or "no-domain",
            )
            continue

        if nome_norm:
            seen_names.add(nome_norm)
        if dominio:
            seen_domains.add(dominio)
        new_rows.append(row)

    return new_rows, duplicates


def append_history(path: Path, rows: Iterable[dict[str, Any]]) -> None:
    """Aggiunge in append allo storico le nuove aziende.

    Crea il file con header se non esiste.
    """

    rows = list(rows)
    if not rows:
        return

    file_exists = path.exists()
    try:
        with path.open("a", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["nome_normalizzato", "dominio", "nome_originale"])
            for row in rows:
                writer.writerow(
                    [
                        normalize_name(row.get("nome_azienda", "")),
                        normalize_domain(row.get("sito_web", "")),
                        row.get("nome_azienda", ""),
                    ]
                )
        logger.info("Storico aggiornato con %d nuove aziende (%s)", len(rows), path)
    except OSError as exc:
        logger.error("Impossibile aggiornare lo storico %s: %s", path, exc)
