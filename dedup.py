"""Deduplica delle aziende contro lo storico persistente (CSV).

Lo storico e' un CSV con 5 colonne:
``nome_normalizzato``, ``dominio``, ``nome_originale``,
``email_normalizzata``, ``telefono_normalizzato``.

Chiavi di confronto (OR logico): un'azienda e' considerata duplicata se
COINCIDE su almeno una tra: nome normalizzato, dominio, email
normalizzata, telefono normalizzato.

Se lo storico esistente ha il vecchio schema a 3 colonne (v1), viene
migrato automaticamente al primo caricamento: il file originale viene
copiato in ``<path>.bak`` e riscritto con i nuovi header, email/telefono
vuoti per le righe pre-esistenti.
"""

from __future__ import annotations

import csv
import logging
import re
import shutil
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
_DIGITS_PATTERN = re.compile(r"\D+")


_HISTORY_HEADER_V2: list[str] = [
    "nome_normalizzato",
    "dominio",
    "nome_originale",
    "email_normalizzata",
    "telefono_normalizzato",
]
_HISTORY_HEADER_V1: list[str] = [
    "nome_normalizzato",
    "dominio",
    "nome_originale",
]

_MIN_PHONE_DIGITS = 6


# ---------------------------------------------------------------------------
# Normalizzazione
# ---------------------------------------------------------------------------


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


def normalize_email(email: str) -> str:
    """Normalizza un'email per il confronto.

    - lowercase + strip
    - stringa vuota se manca la ``@`` (evita di trattare "info" come chiave)
    """

    if not email:
        return ""
    e = email.strip().lower()
    if "@" not in e:
        return ""
    return e


def normalize_phone(phone: str) -> str:
    """Normalizza un numero di telefono per il confronto.

    - tiene solo le cifre
    - rimuove prefissi internazionali comuni per l'Italia/Svizzera
      (``+39``, ``0039``, ``+41``, ``0041``) all'inizio, per confrontare
      il numero locale
    - ritorna stringa vuota se il risultato ha meno di 6 cifre (per
      evitare falsi positivi tipo "0" o "N/D")
    """

    if not phone:
        return ""
    digits = _DIGITS_PATTERN.sub("", phone)
    if not digits:
        return ""

    # Rimuovi prefissi internazionali IT/CH ripetutamente (es. "0039" e poi ancora 0).
    for prefix in ("0039", "0041", "39", "41"):
        # "+39" → "39" dopo aver rimosso i non-digit, quindi qui basta un match del prefisso.
        if digits.startswith(prefix) and len(digits) > len(prefix):
            digits = digits[len(prefix):]
            break

    if len(digits) < _MIN_PHONE_DIGITS:
        return ""
    return digits


# ---------------------------------------------------------------------------
# I/O storico
# ---------------------------------------------------------------------------


def _peek_header(path: Path) -> list[str] | None:
    try:
        with path.open("r", encoding="utf-8", newline="") as f:
            reader = csv.reader(f)
            return next(reader, None)
    except OSError as exc:
        logger.warning("Impossibile leggere header di %s: %s", path, exc)
        return None


def _migrate_history_v1_to_v2(path: Path) -> None:
    """Migra un CSV v1 (3 col) a v2 (5 col).

    - copia il file originale in ``<path>.bak``
    - riscrive il file con i 5 header, riportando le righe pre-esistenti
      con email/telefono vuoti
    """

    backup = path.with_suffix(path.suffix + ".bak")
    try:
        with path.open("r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            old_rows = list(reader)
    except OSError as exc:
        logger.error("Migrazione storico: impossibile leggere %s: %s", path, exc)
        return

    try:
        shutil.copy2(path, backup)
        logger.info("Storico v1 rilevato: backup creato in %s", backup)
    except OSError as exc:
        logger.error(
            "Migrazione storico: impossibile creare backup %s: %s. Abort migrazione.",
            backup,
            exc,
        )
        return

    try:
        with path.open("w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(_HISTORY_HEADER_V2)
            for row in old_rows:
                writer.writerow(
                    [
                        row.get("nome_normalizzato", ""),
                        row.get("dominio", ""),
                        row.get("nome_originale", ""),
                        "",
                        "",
                    ]
                )
        logger.info(
            "Storico migrato a schema v2 (5 colonne): %d righe portate avanti",
            len(old_rows),
        )
    except OSError as exc:
        logger.error("Migrazione storico: scrittura fallita: %s", exc)


def load_history(
    path: Path,
) -> tuple[set[str], set[str], set[str], set[str]]:
    """Carica lo storico e ritorna 4 set: nomi, domini, email, telefoni.

    Se rileva schema v1, migra automaticamente prima di leggere.
    """

    names: set[str] = set()
    domains: set[str] = set()
    emails: set[str] = set()
    phones: set[str] = set()

    if not path.exists():
        logger.info("Storico non trovato (%s), parto da zero", path)
        return names, domains, emails, phones

    header = _peek_header(path)
    if header is None:
        logger.warning("Storico %s vuoto o illeggibile", path)
        return names, domains, emails, phones

    if header == _HISTORY_HEADER_V1:
        _migrate_history_v1_to_v2(path)
    elif header != _HISTORY_HEADER_V2:
        logger.warning(
            "Storico %s ha header inatteso: %s. Provo a leggere in modalita' tollerante.",
            path,
            header,
        )

    try:
        with path.open("r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                n = (row.get("nome_normalizzato") or "").strip()
                d = (row.get("dominio") or "").strip().lower()
                e = (row.get("email_normalizzata") or "").strip().lower()
                p = (row.get("telefono_normalizzato") or "").strip()
                if n:
                    names.add(n)
                if d:
                    domains.add(d)
                if e:
                    emails.add(e)
                if p:
                    phones.add(p)
    except OSError as exc:
        logger.warning("Impossibile leggere lo storico %s: %s", path, exc)

    logger.info(
        "Storico caricato: %d nomi, %d domini, %d email, %d telefoni",
        len(names),
        len(domains),
        len(emails),
        len(phones),
    )
    return names, domains, emails, phones


# ---------------------------------------------------------------------------
# Filtro e append
# ---------------------------------------------------------------------------


def filter_new(
    rows: Iterable[dict[str, Any]],
    known_names: set[str],
    known_domains: set[str],
    known_emails: set[str],
    known_phones: set[str],
) -> tuple[list[dict[str, Any]], int]:
    """Filtra le righe gia' note.

    Confronto duplicati (OR): nome normalizzato, dominio, email, telefono.
    Deduplica anche INTRA-batch: se la stessa azienda compare due volte
    nella stessa risposta di Claude, tiene solo la prima occorrenza.

    NOTA: i set passati vengono **mutati** in-place con gli identificativi
    delle righe nuove accettate. Cosi' il chiamante puo' invocare piu' volte
    ``filter_new`` in sequenza (loop multi-coppia) senza dover ricaricare
    lo storico da disco tra un'iterazione e l'altra.

    Ritorna (nuove_righe, numero_duplicati_scartati).
    """

    new_rows: list[dict[str, Any]] = []
    duplicates = 0

    for row in rows:
        nome_norm = normalize_name(row.get("nome_azienda", ""))
        dominio = normalize_domain(row.get("sito_web", ""))
        email = normalize_email(row.get("email_contatto", ""))
        phone = normalize_phone(row.get("telefono", ""))

        is_dup = False
        matched_on = ""
        if nome_norm and nome_norm in known_names:
            is_dup = True
            matched_on = "nome"
        elif dominio and dominio in known_domains:
            is_dup = True
            matched_on = "dominio"
        elif email and email in known_emails:
            is_dup = True
            matched_on = "email"
        elif phone and phone in known_phones:
            is_dup = True
            matched_on = "telefono"

        if is_dup:
            duplicates += 1
            logger.debug(
                "Duplicato scartato (match=%s): %s",
                matched_on,
                row.get("nome_azienda", ""),
            )
            continue

        if nome_norm:
            known_names.add(nome_norm)
        if dominio:
            known_domains.add(dominio)
        if email:
            known_emails.add(email)
        if phone:
            known_phones.add(phone)
        new_rows.append(row)

    return new_rows, duplicates


def append_history(path: Path, rows: Iterable[dict[str, Any]]) -> None:
    """Aggiunge in append allo storico le nuove aziende (5 colonne).

    Crea il file con header v2 se non esiste.
    """

    rows = list(rows)
    if not rows:
        return

    file_exists = path.exists()
    try:
        with path.open("a", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(_HISTORY_HEADER_V2)
            for row in rows:
                writer.writerow(
                    [
                        normalize_name(row.get("nome_azienda", "")),
                        normalize_domain(row.get("sito_web", "")),
                        row.get("nome_azienda", ""),
                        normalize_email(row.get("email_contatto", "")),
                        normalize_phone(row.get("telefono", "")),
                    ]
                )
        logger.info(
            "Storico aggiornato con %d nuove aziende (%s)", len(rows), path
        )
    except OSError as exc:
        logger.error("Impossibile aggiornare lo storico %s: %s", path, exc)
