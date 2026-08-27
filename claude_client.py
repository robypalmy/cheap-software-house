"""Wrapper attorno al client Anthropic con web_search e retry esponenziale.

Espone due funzioni principali:
- ``build_prompt``: costruisce il prompt per la ricerca di aziende in una
  regione/settore.
- ``search_companies``: esegue la chiamata all'API con retry e ritorna una
  lista di dizionari già validati (parsing JSON incluso).
"""

from __future__ import annotations

import json
import logging
import os
import re
import time
from typing import Any

import anthropic

from config import (
    API_BACKOFF_BASE_SECONDS,
    API_MAX_RETRIES,
    MAX_TOKENS,
    MODELLO,
    WEB_SEARCH_MAX_USES,
)


logger = logging.getLogger(__name__)


REQUIRED_FIELDS: tuple[str, ...] = (
    "nome_azienda",
    "settore",
    "comune",
    "provincia_o_cantone",
    "sito_web",
    "problemi_rilevati",
    "email_contatto",
    "telefono",
    "fonte",
    "note",
)


class ClaudeAPIError(RuntimeError):
    """Sollevata quando la chiamata a Claude fallisce definitivamente."""


def build_prompt(regione: str, settore: str, numero: int) -> str:
    """Costruisce il prompt per Claude.

    Il prompt richiede esplicitamente output JSON puro, per semplificare il
    parsing lato Python. Il ``settore`` "qualsiasi" viene tradotto in una
    ricerca generica.
    """

    settore_clause = (
        "qualsiasi settore (privilegia PMI tradizionali, no tech/digital)"
        if settore.strip().lower() in {"qualsiasi", "any", ""}
        else f"settore: {settore}"
    )

    return f"""Sei un assistente di ricerca commerciale B2B. Devi trovare {numero} PMI reali e verificabili nella zona "{regione}" ({settore_clause}) che abbiano un sito web OUTDATED, mal fatto, oppure siano completamente assenti online.

Criteri di target (importante):
- PMI locali, NON multinazionali né grandi gruppi.
- Priorità ad aziende con bassa maturita' digitale (non gia' clienti di grosse agenzie web), per evitare concorrenza al ribasso sul prezzo.
- Cerca segnali di sito outdated: design anni 2000, non responsive/mobile, no HTTPS, contenuti fermi da anni, tecnologie deprecate (Flash, tabelle di layout, Frontpage), oppure semplicemente sito assente ma azienda attiva (P.IVA valida, presenza su registro imprese o directory ufficiali).
- Fonti ammesse: siti aziendali pubblici, registroimprese.it, camere di commercio, pagine LinkedIn aziendali pubbliche, directory locali ufficiali, PagineGialle. NON usare data broker, NON estrarre dati personali privati.
- I contatti (email, telefono) devono essere quelli GENERALI e PUBBLICI dell'azienda, pubblicati sul sito o su fonti ufficiali. Se non li trovi verificati, lascia la stringa vuota. NON inventare mai.

REGOLE DI OUTPUT (critiche):
- Rispondi ESCLUSIVAMENTE con un array JSON valido. Nessun testo prima o dopo. Nessun blocco ```json. Nessun commento.
- L'array deve contenere esattamente {numero} oggetti (o meno se non ne trovi abbastanza di validi — meglio pochi ma veri).
- Ogni oggetto deve avere ESATTAMENTE queste chiavi (stringhe, tranne problemi_rilevati che e' una lista di stringhe):
  - "nome_azienda": ragione sociale come da fonte ufficiale
  - "settore": settore o attivita' principale
  - "comune": comune / citta'
  - "provincia_o_cantone": sigla provincia italiana o nome cantone svizzero
  - "sito_web": URL completo del sito (o stringa vuota se l'azienda non ha sito)
  - "problemi_rilevati": lista di stringhe, es. ["non responsive", "no HTTPS", "design datato", "contenuti fermi al 2015", "sito assente"]
  - "email_contatto": email generale pubblica (info@, contatti@...) oppure stringa vuota
  - "telefono": numero fisso/generale pubblico oppure stringa vuota
  - "fonte": URL o descrizione della fonte dove hai trovato l'azienda
  - "note": eventuali note utili al commerciale (max 1-2 frasi)

Non aggiungere altri campi. Non inserire null: usa "" per stringhe mancanti e [] per liste vuote.
"""


def _extract_text_blocks(response: Any) -> str:
    """Concatena i blocchi 'text' della risposta Claude in una singola stringa."""

    parts: list[str] = []
    for block in getattr(response, "content", []) or []:
        block_type = getattr(block, "type", None)
        if block_type == "text":
            text = getattr(block, "text", "") or ""
            if text:
                parts.append(text)
    return "\n".join(parts).strip()


def _extract_json_array(raw: str) -> list[dict[str, Any]]:
    """Estrae un array JSON da un testo, tollerando eventuali code fence.

    Supporta:
    - JSON puro (caso normale con il prompt corrente)
    - Blocchi ```json ... ```
    - JSON annegato dentro altro testo (ricerca del primo ``[`` e dell'ultimo ``]``)
    """

    if not raw:
        raise ValueError("Risposta vuota da Claude")

    fenced = re.search(r"```(?:json)?\s*(\[.*?\])\s*```", raw, re.DOTALL)
    if fenced:
        candidate = fenced.group(1)
    else:
        start = raw.find("[")
        end = raw.rfind("]")
        if start == -1 or end == -1 or end <= start:
            raise ValueError("Nessun array JSON trovato nella risposta")
        candidate = raw[start : end + 1]

    data = json.loads(candidate)
    if not isinstance(data, list):
        raise ValueError("Il JSON ricevuto non e' un array")
    return data


def _validate_row(row: Any) -> dict[str, Any] | None:
    """Valida e normalizza una singola riga.

    Ritorna il dict normalizzato oppure ``None`` se la riga e' malformata
    (non ha almeno ``nome_azienda`` non vuoto — il sito puo' essere vuoto se
    l'azienda e' proprio assente online, che e' un caso interessante).
    """

    if not isinstance(row, dict):
        logger.warning("Riga scartata (non e' un oggetto): %r", row)
        return None

    normalized: dict[str, Any] = {}
    for key in REQUIRED_FIELDS:
        value = row.get(key, "")
        if key == "problemi_rilevati":
            if value is None:
                value = []
            elif isinstance(value, str):
                value = [value] if value.strip() else []
            elif not isinstance(value, list):
                value = []
            value = [str(v).strip() for v in value if str(v).strip()]
        else:
            if value is None:
                value = ""
            value = str(value).strip()
        normalized[key] = value

    nome = normalized["nome_azienda"]
    sito = normalized["sito_web"]
    problemi = normalized["problemi_rilevati"]

    if not nome:
        logger.warning("Riga scartata: nome_azienda vuoto (%r)", row)
        return None

    if not sito and not problemi:
        logger.warning(
            "Riga scartata: ne' sito_web ne' problemi_rilevati presenti per %s",
            nome,
        )
        return None

    return normalized


def _sleep_backoff(attempt: int) -> None:
    delay = API_BACKOFF_BASE_SECONDS * (2 ** (attempt - 1))
    logger.info("Backoff: attendo %.1fs prima del retry", delay)
    time.sleep(delay)


def search_companies(
    regione: str,
    settore: str,
    numero: int,
    *,
    client: anthropic.Anthropic | None = None,
    model: str = MODELLO,
) -> list[dict[str, Any]]:
    """Interroga Claude con web_search e ritorna le righe validate.

    Solleva ``ClaudeAPIError`` se tutti i tentativi falliscono.
    """

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ClaudeAPIError(
            "ANTHROPIC_API_KEY non impostata (vedi .env.example)"
        )

    if client is None:
        client = anthropic.Anthropic(api_key=api_key)

    prompt = build_prompt(regione=regione, settore=settore, numero=numero)

    last_error: Exception | None = None
    for attempt in range(1, API_MAX_RETRIES + 1):
        try:
            logger.info(
                "Chiamata API Claude (tentativo %d/%d) modello=%s",
                attempt,
                API_MAX_RETRIES,
                model,
            )
            response = client.messages.create(
                model=model,
                max_tokens=MAX_TOKENS,
                tools=[
                    {
                        "type": "web_search_20250305",
                        "name": "web_search",
                        "max_uses": WEB_SEARCH_MAX_USES,
                    }
                ],
                messages=[{"role": "user", "content": prompt}],
            )
            raw_text = _extract_text_blocks(response)
            logger.debug("Risposta grezza Claude (len=%d)", len(raw_text))
            data = _extract_json_array(raw_text)

            rows: list[dict[str, Any]] = []
            for item in data:
                normalized = _validate_row(item)
                if normalized is not None:
                    rows.append(normalized)

            logger.info(
                "Ricevute %d righe dall'API, %d valide dopo validazione",
                len(data),
                len(rows),
            )
            return rows

        except (
            anthropic.APIStatusError,
            anthropic.APIConnectionError,
            anthropic.RateLimitError,
        ) as exc:
            last_error = exc
            logger.warning("Errore API Claude (tentativo %d): %s", attempt, exc)
            if attempt < API_MAX_RETRIES:
                _sleep_backoff(attempt)
        except (ValueError, json.JSONDecodeError) as exc:
            last_error = exc
            logger.warning(
                "Errore di parsing JSON (tentativo %d): %s", attempt, exc
            )
            if attempt < API_MAX_RETRIES:
                _sleep_backoff(attempt)

    raise ClaudeAPIError(
        f"Tutti i {API_MAX_RETRIES} tentativi falliti. Ultimo errore: {last_error}"
    )
