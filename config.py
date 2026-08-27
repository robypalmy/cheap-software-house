"""Configurazione centralizzata per lo script di prospecting.

I valori qui definiti fungono da default. Possono essere sovrascritti da CLI
(vedi ``main.py``) o da variabili d'ambiente per i segreti.
"""

from __future__ import annotations

from pathlib import Path


PROJECT_ROOT: Path = Path(__file__).resolve().parent


LOGS_DIR: Path = PROJECT_ROOT / "logs"
EXCEL_FILE: Path = PROJECT_ROOT / "lead_prospecting.xlsx"
STORICO_FILE: Path = PROJECT_ROOT / "storico_aziende.csv"
STATE_FILE: Path = PROJECT_ROOT / ".rotation_state.json"


# Rotazione geografica: ad ogni run viene scelta una regione/provincia diversa
# per garantire copertura nel tempo senza ripetersi.
REGIONI: list[str] = [
    "Ticino",
    "Provincia di Milano",
    "Provincia di Monza e Brianza",
    "Provincia di Varese",
    "Provincia di Como",
    "Provincia di Lecco",
    "Provincia di Bergamo",
    "Provincia di Brescia",
    "Provincia di Pavia",
    "Provincia di Lodi",
    "Provincia di Cremona",
    "Provincia di Mantova",
    "Provincia di Sondrio",
    "Provincia di Torino",
    "Provincia di Novara",
    "Provincia di Vercelli",
    "Provincia di Biella",
    "Provincia di Cuneo",
    "Provincia di Asti",
    "Provincia di Alessandria",
    "Provincia del Verbano-Cusio-Ossola",
]


SETTORI: list[str] = [
    "manifatturiero",
    "commercio al dettaglio",
    "ristorazione",
    "artigianato",
    "edilizia e serramenti",
    "meccanica di precisione",
    "servizi alla persona",
    "autofficine e carrozzerie",
    "agricoltura e agroalimentare",
    "impiantistica",
    "qualsiasi",
]


NUMERO_AZIENDE_PER_RUN: int = 18

MODELLO: str = "claude-sonnet-5"

WEB_SEARCH_MAX_USES: int = 15

# Alzato a 16000 per accomodare risposte piu' lunghe con molti risultati:
# con 8000 capitava di venire troncati (stop_reason="max_tokens") e perdere
# righe. Il client rileva comunque il troncamento e fa retry.
MAX_TOKENS: int = 16000

API_MAX_RETRIES: int = 3
API_BACKOFF_BASE_SECONDS: float = 4.0


# Rotazione "stay-until-exhausted": una coppia (regione, settore) viene
# considerata esaurita dopo N run consecutivi in cui, dopo la deduplica,
# non emergono nuove aziende. A quel punto lo script passa alla coppia
# successiva. Con N=2 si evita di dichiarare esaurita una coppia solo
# per un run sfortunato (rate limit, risposta scarsa, ecc.).
SATURATION_THRESHOLD: int = 2


# Quante coppie (regione, settore) processare per ogni singolo run in
# modalita' automatica. Con 21 regioni x 11 settori = 231 coppie e
# SATURATION_THRESHOLD=2, processare piu' coppie per notte accelera la
# copertura. Modalita' manuale (--regione/--settore) resta 1 coppia.
COPPIE_PER_RUN: int = 3


# Ritenzione log giornalieri in giorni. All'inizio di ogni run, i file
# ``logs/*.log`` piu' vecchi di questa soglia vengono cancellati.
# Impostare a 0 o negativo per disabilitare la pulizia.
LOG_RETENTION_DAYS: int = 30


STATI_LEAD: list[str] = [
    "Da contattare",
    "Email inviata",
    "Contattato telefonicamente",
    "Appuntamento fissato",
    "Non interessato",
]


EXCEL_HEADERS: list[str] = [
    "Data_Trovato",
    "Azienda",
    "Settore",
    "Comune",
    "Provincia_Cantone",
    "Sito_Web",
    "Problemi_Rilevati",
    "Email",
    "Telefono",
    "Fonte",
    "Stato",
    "Note",
]
