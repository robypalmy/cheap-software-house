"""Entry point per la ricerca automatica di lead prospecting.

Esempi:

    # Test manuale (batch piccolo, singola coppia, NON aggiorna lo stato)
    python main.py --regione "Ticino" --settore "ristorazione" --numero 5

    # Esecuzione automatica: processa COPPIE_PER_RUN coppie in un solo run
    python main.py

    # Automatica con numero di coppie custom
    python main.py --coppie 5

Rotazione automatica:
- Si RESTA sulla stessa coppia (regione, settore) finche' produce nuovi lead.
- La coppia viene considerata "esaurita" dopo ``SATURATION_THRESHOLD`` run
  consecutivi con 0 nuove aziende dopo la deduplica.
- A quel punto si passa alla coppia successiva in ordine "regione esterna,
  settore interno".
- Quando tutte le coppie sono esaurite, i flag vengono resettati e si
  ricomincia (siti nuovi/chiusi/rifatti nel frattempo).

In modalita' AUTOMATICA lo script processa fino a ``--coppie`` coppie in
sequenza, aggiornando Excel/storico/stato in modo incrementale dopo ogni
coppia (un'interruzione a meta' non perde il lavoro gia' fatto). Se una
coppia fallisce con ``ClaudeAPIError``, si passa alla successiva.

Stato persistente in ``.rotation_state.json`` per ogni coppia:
``runs``, ``zero_streak``, ``exhausted``, ``total_new``, ``last_run``.
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

import config
from claude_client import ClaudeAPIError, search_companies
from dedup import append_history, filter_new, load_history
from excel_writer import ExcelLockedError, append_companies


logger = logging.getLogger("prospecting")


# ---------------------------------------------------------------------------
# Logging + housekeeping
# ---------------------------------------------------------------------------


def _setup_logging(logs_dir: Path) -> None:
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_path = logs_dir / f"{date.today().isoformat()}.log"

    root = logging.getLogger()
    root.setLevel(logging.INFO)

    for handler in list(root.handlers):
        root.removeHandler(handler)

    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(fmt)
    root.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(fmt)
    root.addHandler(stream_handler)


def _cleanup_old_logs(logs_dir: Path, retention_days: int) -> None:
    """Cancella i file ``logs/*.log`` piu' vecchi di ``retention_days`` giorni.

    La data del file viene dedotta prima dal nome ``YYYY-MM-DD.log``; se il
    nome non e' parsabile, si usa l'mtime del file come fallback.
    Errori di cancellazione non bloccano il run (solo warning).
    """

    if retention_days is None or retention_days <= 0:
        return
    if not logs_dir.exists():
        return

    cutoff = date.today() - timedelta(days=retention_days)
    deleted = 0
    kept = 0
    today_log_name = f"{date.today().isoformat()}.log"

    for log_file in logs_dir.glob("*.log"):
        # Non cancellare il log del giorno stesso, anche se retention_days=0
        if log_file.name == today_log_name:
            kept += 1
            continue

        try:
            file_date = date.fromisoformat(log_file.stem)
        except ValueError:
            try:
                file_date = datetime.fromtimestamp(
                    log_file.stat().st_mtime
                ).date()
            except OSError as exc:
                logger.warning(
                    "Impossibile determinare l'eta' di %s: %s", log_file, exc
                )
                continue

        if file_date < cutoff:
            try:
                log_file.unlink()
                deleted += 1
            except OSError as exc:
                logger.warning("Impossibile cancellare log %s: %s", log_file, exc)
        else:
            kept += 1

    if deleted:
        logger.info(
            "Log cleanup: cancellati %d file piu' vecchi di %d giorni (mantenuti %d)",
            deleted,
            retention_days,
            kept,
        )


# ---------------------------------------------------------------------------
# Stato rotazione
# ---------------------------------------------------------------------------

_PAIR_SEP = "||"


def _pair_key(regione: str, settore: str) -> str:
    return f"{regione}{_PAIR_SEP}{settore}"


def _default_state() -> dict[str, Any]:
    return {"pairs": {}, "cursor": {"regione_idx": 0, "settore_idx": 0}}


def _load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return _default_state()
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict) or "pairs" not in data:
            logger.warning("Stato rotazione con schema legacy o invalido: reset")
            return _default_state()
        data.setdefault("cursor", {"regione_idx": 0, "settore_idx": 0})
        data.setdefault("pairs", {})
        return data
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        logger.warning("Stato rotazione illeggibile (%s), reset: %s", path, exc)
        return _default_state()


def _save_state(path: Path, state: dict[str, Any]) -> None:
    try:
        with path.open("w", encoding="utf-8") as f:
            json.dump(state, f, indent=2, ensure_ascii=False)
    except OSError as exc:
        logger.warning("Impossibile salvare stato rotazione: %s", exc)


def _reset_exhausted_flags(state: dict[str, Any]) -> None:
    for pair in state.get("pairs", {}).values():
        pair["exhausted"] = False
        pair["zero_streak"] = 0


def _pick_pair(state: dict[str, Any]) -> tuple[str, str]:
    """Sceglie la prossima coppia (regione, settore) NON esaurita.

    Ordine: regione esterna, settore interno. Il cursore viene posizionato
    sulla coppia scelta. Se tutte sono esaurite, i flag vengono resettati e
    si riparte da (0, 0).
    """

    pairs_state: dict[str, Any] = state["pairs"]
    cursor: dict[str, int] = state["cursor"]

    n_reg = len(config.REGIONI)
    n_set = len(config.SETTORI)
    total = n_reg * n_set

    r_idx = int(cursor.get("regione_idx", 0)) % n_reg
    s_idx = int(cursor.get("settore_idx", 0)) % n_set

    for _ in range(total):
        regione = config.REGIONI[r_idx]
        settore = config.SETTORI[s_idx]
        key = _pair_key(regione, settore)
        pair = pairs_state.get(key, {})
        if not pair.get("exhausted", False):
            cursor["regione_idx"] = r_idx
            cursor["settore_idx"] = s_idx
            return regione, settore

        s_idx += 1
        if s_idx >= n_set:
            s_idx = 0
            r_idx = (r_idx + 1) % n_reg

    logger.warning(
        "Tutte le %d coppie regione+settore risultano esaurite: reset flag e riparto",
        total,
    )
    _reset_exhausted_flags(state)
    cursor["regione_idx"] = 0
    cursor["settore_idx"] = 0
    return config.REGIONI[0], config.SETTORI[0]


def _advance_cursor(state: dict[str, Any]) -> None:
    """Avanza il cursore di una posizione (regione esterna, settore interno)."""

    cursor = state["cursor"]
    n_reg = len(config.REGIONI)
    n_set = len(config.SETTORI)

    s_idx = int(cursor.get("settore_idx", 0)) + 1
    r_idx = int(cursor.get("regione_idx", 0))
    if s_idx >= n_set:
        s_idx = 0
        r_idx = (r_idx + 1) % n_reg
    cursor["settore_idx"] = s_idx
    cursor["regione_idx"] = r_idx


def _update_pair_after_run(
    state: dict[str, Any],
    regione: str,
    settore: str,
    new_count: int,
) -> None:
    """Aggiorna la coppia dopo un run riuscito.

    - Se ``new_count > 0``: resetta ``zero_streak``, non tocca il cursore.
    - Se ``new_count == 0``: incrementa ``zero_streak``. Se raggiunge la
      soglia, marca la coppia esaurita e avanza il cursore.
    """

    key = _pair_key(regione, settore)
    pairs_state = state.setdefault("pairs", {})
    pair = pairs_state.setdefault(
        key,
        {
            "runs": 0,
            "zero_streak": 0,
            "exhausted": False,
            "total_new": 0,
            "last_run": "",
        },
    )
    pair["runs"] = int(pair.get("runs", 0)) + 1
    pair["total_new"] = int(pair.get("total_new", 0)) + int(new_count)
    pair["last_run"] = date.today().isoformat()

    if new_count > 0:
        pair["zero_streak"] = 0
        logger.info(
            "Coppia %r: +%d lead (totale storico %d). Resto qui al prossimo run.",
            key,
            new_count,
            pair["total_new"],
        )
        return

    pair["zero_streak"] = int(pair.get("zero_streak", 0)) + 1
    logger.info(
        "Coppia %r: 0 nuovi lead (streak zero=%d/%d).",
        key,
        pair["zero_streak"],
        config.SATURATION_THRESHOLD,
    )
    if pair["zero_streak"] >= config.SATURATION_THRESHOLD:
        pair["exhausted"] = True
        logger.info(
            "Coppia %r marcata come ESAURITA dopo %d run consecutivi a zero. Avanzo cursore.",
            key,
            pair["zero_streak"],
        )
        _advance_cursor(state)


# ---------------------------------------------------------------------------
# CLI + orchestrazione
# ---------------------------------------------------------------------------


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Cerca PMI in Ticino/Nord Italia con siti web outdated e le "
            "salva in un file Excel per prospecting."
        )
    )
    parser.add_argument(
        "--regione",
        type=str,
        default=None,
        help="Regione/provincia da cercare (default: rotazione automatica)",
    )
    parser.add_argument(
        "--settore",
        type=str,
        default=None,
        help='Settore da cercare, oppure "qualsiasi" (default: rotazione)',
    )
    parser.add_argument(
        "--numero",
        type=int,
        default=config.NUMERO_AZIENDE_PER_RUN,
        help=f"Numero di aziende per singola coppia (default: {config.NUMERO_AZIENDE_PER_RUN})",
    )
    parser.add_argument(
        "--coppie",
        type=int,
        default=config.COPPIE_PER_RUN,
        help=(
            "Quante coppie (regione, settore) processare per run in modalita' "
            f"automatica (default: {config.COPPIE_PER_RUN}). Ignorato se "
            "--regione/--settore sono passati (override manuale = 1 coppia)."
        ),
    )
    parser.add_argument(
        "--modello",
        type=str,
        default=config.MODELLO,
        help=f"Modello Claude (default: {config.MODELLO})",
    )
    parser.add_argument(
        "--excel",
        type=Path,
        default=config.EXCEL_FILE,
        help=f"Path del file Excel di output (default: {config.EXCEL_FILE.name})",
    )
    parser.add_argument(
        "--storico",
        type=Path,
        default=config.STORICO_FILE,
        help=f"Path del CSV di storico (default: {config.STORICO_FILE.name})",
    )
    return parser.parse_args(argv)


def _process_single_pair(
    regione: str,
    settore: str,
    numero: int,
    modello: str,
    excel_path: Path,
    storico_path: Path,
    known_sets: tuple[set[str], set[str], set[str], set[str]],
) -> tuple[int, int, int | None]:
    """Esegue una singola coppia: API + dedup + Excel + storico.

    ``known_sets`` viene mutato in-place da ``filter_new`` con i nuovi
    identificativi, cosi' il caller puo' passarlo alle iterazioni successive
    senza ricaricare lo storico da disco.

    Ritorna ``(new_count, duplicates, error_exit_code)``:
    - ``error_exit_code`` e' ``None`` in caso di successo,
      ``2`` per ClaudeAPIError, ``3`` per ExcelLockedError.
    - Su ExcelLockedError il caller dovrebbe interrompere il loop (il
      lock non si sblocca da solo entro un run).
    """

    known_names, known_domains, known_emails, known_phones = known_sets

    try:
        rows = search_companies(
            regione=regione, settore=settore, numero=numero, model=modello
        )
    except ClaudeAPIError as exc:
        logger.error("Chiamata API fallita per (%s, %s): %s", regione, settore, exc)
        return 0, 0, 2

    logger.info(
        "Aziende ricevute e valide per (%s, %s): %d", regione, settore, len(rows)
    )

    new_rows, duplicates = filter_new(
        rows, known_names, known_domains, known_emails, known_phones
    )
    logger.info(
        "Deduplica (%s, %s): %d nuove, %d duplicati scartati",
        regione,
        settore,
        len(new_rows),
        duplicates,
    )

    if not new_rows:
        return 0, duplicates, None

    try:
        append_companies(excel_path, new_rows)
    except ExcelLockedError as exc:
        logger.error("Excel non scrivibile: %s", exc)
        return 0, duplicates, 3

    append_history(storico_path, new_rows)
    return len(new_rows), duplicates, None


def run(args: argparse.Namespace) -> int:
    _setup_logging(config.LOGS_DIR)
    _cleanup_old_logs(config.LOGS_DIR, config.LOG_RETENTION_DAYS)
    load_dotenv()

    started_at = datetime.now()
    logger.info("=== START run @ %s ===", started_at.isoformat(timespec="seconds"))

    state = _load_state(config.STATE_FILE)
    manual_override = bool(args.regione) or bool(args.settore)

    if manual_override:
        coppie_da_fare = 1
        logger.info(
            "Modalita' MANUALE (override CLI): singola coppia, stato rotazione NON aggiornato"
        )
    else:
        coppie_da_fare = max(1, int(args.coppie))
        logger.info(
            "Modalita' AUTOMATICA: processero' fino a %d coppie in questo run",
            coppie_da_fare,
        )

    known_sets = load_history(args.storico)

    total_new = 0
    total_duplicates = 0
    successes = 0
    failures = 0
    first_error_code: int | None = None
    stop_reason = "completato"

    for i in range(coppie_da_fare):
        if manual_override:
            regione = args.regione or config.REGIONI[0]
            settore = args.settore or config.SETTORI[0]
        else:
            regione, settore = _pick_pair(state)

        logger.info(
            "--- Coppia %d/%d: (%s, %s) ---",
            i + 1,
            coppie_da_fare,
            regione,
            settore,
        )

        new_count, dups, err = _process_single_pair(
            regione=regione,
            settore=settore,
            numero=args.numero,
            modello=args.modello,
            excel_path=args.excel,
            storico_path=args.storico,
            known_sets=known_sets,
        )

        total_new += new_count
        total_duplicates += dups

        if err is None:
            successes += 1
            if not manual_override:
                _update_pair_after_run(state, regione, settore, new_count)
                _save_state(config.STATE_FILE, state)
        else:
            failures += 1
            if first_error_code is None:
                first_error_code = err

            if err == 3:
                # Excel lockato: le coppie successive non riuscirebbero comunque a scrivere.
                logger.error(
                    "Interrompo il loop: file Excel bloccato, riprova dopo aver chiuso il file."
                )
                stop_reason = "excel-locked"
                break
            # ClaudeAPIError: passo alla coppia successiva senza toccare lo stato.

    duration = (datetime.now() - started_at).total_seconds()
    logger.info(
        "RIEPILOGO run: coppie ok=%d, coppie fallite=%d, lead nuovi totali=%d, "
        "duplicati totali=%d, durata=%.1fs, stop=%s",
        successes,
        failures,
        total_new,
        total_duplicates,
        duration,
        stop_reason,
    )

    if first_error_code is not None:
        logger.info("=== END run (FAIL parziale, exit=%d) ===", first_error_code)
        return first_error_code

    logger.info("=== END run (OK) ===")
    return 0


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        return run(args)
    except Exception:
        logger.exception("Errore non gestito, termino")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
