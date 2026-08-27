"""Entry point per la ricerca automatica di lead prospecting.

Esempi:

    # Test manuale (batch piccolo)
    python main.py --regione "Ticino" --settore "ristorazione" --numero 5

    # Esecuzione automatica (rotazione stay-until-exhausted)
    python main.py

Rotazione automatica:
- Si RESTA sulla stessa coppia (regione, settore) finche' produce nuovi lead.
- La coppia viene considerata "esaurita" dopo ``SATURATION_THRESHOLD`` run
  consecutivi con 0 nuove aziende dopo la deduplica.
- A quel punto si passa alla coppia successiva in ordine "regione esterna,
  settore interno": esauriti tutti i settori della regione corrente, si
  passa alla regione successiva.
- Quando tutte le coppie sono esaurite, i flag vengono resettati e si
  ricomincia (le aziende cambiano nel tempo, siti nuovi/chiusi/rifatti).

Stato persistente in ``.rotation_state.json`` per ogni coppia:
``runs``, ``zero_streak``, ``exhausted``, ``total_new``, ``last_run``.

Lo script e' pensato per essere lanciato da cron/Task Scheduler senza
intervento manuale. In caso di errore, logga e termina con exit code != 0.
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

import config
from claude_client import ClaudeAPIError, search_companies
from dedup import append_history, filter_new, load_history
from excel_writer import ExcelLockedError, append_companies


logger = logging.getLogger("prospecting")


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
        help=f"Numero di aziende da cercare (default: {config.NUMERO_AZIENDE_PER_RUN})",
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


def run(args: argparse.Namespace) -> int:
    _setup_logging(config.LOGS_DIR)
    load_dotenv()

    started_at = datetime.now()
    logger.info("=== START run @ %s ===", started_at.isoformat(timespec="seconds"))

    state = _load_state(config.STATE_FILE)

    if args.regione and args.settore:
        regione, settore = args.regione, args.settore
        manual_override = True
    else:
        auto_regione, auto_settore = _pick_pair(state)
        regione = args.regione or auto_regione
        settore = args.settore or auto_settore
        manual_override = bool(args.regione) or bool(args.settore)

    logger.info(
        "Parametri: regione=%r settore=%r numero=%d modello=%s (rotazione %s)",
        regione,
        settore,
        args.numero,
        args.modello,
        "MANUALE" if manual_override else "AUTO",
    )

    try:
        rows = search_companies(
            regione=regione,
            settore=settore,
            numero=args.numero,
            model=args.modello,
        )
    except ClaudeAPIError as exc:
        logger.error("Chiamata API fallita: %s", exc)
        logger.info("=== END run (FAIL API, stato NON aggiornato) ===")
        return 2

    logger.info("Aziende ricevute e valide: %d", len(rows))

    known_names, known_domains = load_history(args.storico)
    new_rows, duplicates = filter_new(rows, known_names, known_domains)
    logger.info(
        "Deduplica: %d nuove, %d duplicati scartati", len(new_rows), duplicates
    )

    if new_rows:
        try:
            written = append_companies(args.excel, new_rows)
        except ExcelLockedError as exc:
            logger.error("Excel non scrivibile: %s", exc)
            logger.info("=== END run (FAIL Excel locked, stato NON aggiornato) ===")
            return 3
        append_history(args.storico, new_rows)
    else:
        written = 0
        logger.info("Nessuna azienda nuova, non tocco Excel ne' storico")

    if not manual_override:
        _update_pair_after_run(state, regione, settore, len(new_rows))
        _save_state(config.STATE_FILE, state)
    else:
        logger.info("Override manuale via CLI: stato rotazione NON aggiornato")

    duration = (datetime.now() - started_at).total_seconds()
    logger.info(
        "Run OK: %d lead scritti su %s in %.1fs",
        written,
        args.excel,
        duration,
    )
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
