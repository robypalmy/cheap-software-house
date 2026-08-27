# Prospecting siti outdated — Ticino & Nord Italia

Script Python che ogni notte cerca **PMI con siti web datati o assenti** in
Ticino, Lombardia e Piemonte, e produce un file Excel pronto da leggere la
mattina dopo per il prospecting commerciale (vendita di siti nuovi).

Nessun invio email, nessuna azione automatica: **solo ricerca + output Excel**.
L'invio email sara' un modulo separato che leggera' lo stesso file.

---

## Struttura del progetto

```
.
├── main.py                # entry point / orchestratore
├── config.py              # parametri (regioni, settori, path, modello…)
├── claude_client.py       # wrapper Anthropic + web_search + retry
├── dedup.py               # deduplica contro storico (nome + dominio)
├── excel_writer.py        # append + formattazione + data validation
├── requirements.txt
├── .env.example           # ANTHROPIC_API_KEY=
├── README.md
└── logs/                  # log giornalieri (YYYY-MM-DD.log)
```

File generati a runtime:
- `lead_prospecting.xlsx` — output principale (append giornaliero)
- `storico_aziende.csv` — storico aziende gia' viste (per deduplica)
- `.rotation_state.json` — indice corrente della rotazione regione/settore
- `logs/YYYY-MM-DD.log` — log del giorno

---

## Setup

Serve **Python 3.11+**.

```bash
git clone <questo repo>
cd cheap-software-house

python3 -m venv .venv
source .venv/bin/activate            # Windows: .venv\Scripts\activate

pip install -r requirements.txt

cp .env.example .env
# apri .env e incolla la chiave API Anthropic:
# ANTHROPIC_API_KEY=sk-ant-...
```

---

## Test manuale (batch piccolo)

Prima di metterlo in cron, verifica la qualita' dei risultati con un batch
piccolo (5 aziende):

```bash
python main.py --regione "Ticino" --settore "manifatturiero" --numero 5
```

Controlla:
1. `logs/YYYY-MM-DD.log` — deve terminare con `=== END run (OK) ===`.
2. `lead_prospecting.xlsx` — deve avere le colonne corrette, l'header colorato,
   e la lista a discesa nella colonna **Stato** (default `Da contattare`).
3. `storico_aziende.csv` — deve contenere le righe appena scritte.

Rilancia lo stesso comando: nel log dovresti vedere **duplicati scartati**
e nessuna nuova riga aggiunta all'Excel (conferma che la deduplica funziona).

---

## Parametri CLI

| Flag | Default | Descrizione |
|---|---|---|
| `--regione` | rotazione automatica su `config.REGIONI` | Zona da cercare |
| `--settore` | rotazione automatica su `config.SETTORI` | Settore, o `qualsiasi` |
| `--numero` | `18` (`NUMERO_AZIENDE_PER_RUN`) | Quante aziende per run |
| `--modello` | `claude-sonnet-5` | Modello Anthropic |
| `--excel` | `lead_prospecting.xlsx` | Path Excel di output |
| `--storico` | `storico_aziende.csv` | Path CSV storico |

Senza flag lo script usa la **rotazione automatica intelligente** (vedi
sotto): resta sulla stessa coppia `(regione, settore)` finche' produce
nuovi lead, poi passa alla successiva.

Per personalizzare regioni/settori/numero di default, modifica `config.py`.

---

## Rotazione automatica "stay-until-exhausted"

Lo script **NON** cicla regioni/settori a ogni run in modo cieco. Invece:

1. Sceglie una coppia `(regione, settore)` — es. `(Ticino, ristorazione)`.
2. **Ci resta** finche' produce nuovi lead (post-deduplica).
3. Quando due run consecutivi restituiscono **0 nuove aziende** dopo la
   deduplica, la coppia viene marcata come `exhausted` e lo script passa
   alla coppia successiva.
4. Ordine di avanzamento: **regione fissa, cicla tutti i settori**, poi
   cambia regione. Esempio:
   `(Ticino, manifatturiero) → (Ticino, commercio) → ... → (Ticino, qualsiasi) → (Provincia di Milano, manifatturiero) → ...`
5. Quando **tutte** le coppie sono esaurite, i flag vengono resettati e si
   ricomincia dall'inizio (utile: nel frattempo aziende nuove aprono, siti
   vecchi vengono rifatti, ecc.).

### Stato persistente

Il file `.rotation_state.json` tiene lo stato di ogni coppia:

```json
{
  "pairs": {
    "Ticino||ristorazione": {
      "runs": 3,
      "zero_streak": 0,
      "exhausted": false,
      "total_new": 42,
      "last_run": "2026-08-27"
    },
    "Ticino||manifatturiero": {
      "runs": 5,
      "zero_streak": 2,
      "exhausted": true,
      "total_new": 61,
      "last_run": "2026-08-24"
    }
  },
  "cursor": { "regione_idx": 0, "settore_idx": 1 }
}
```

- `runs`: numero di run eseguiti su questa coppia.
- `zero_streak`: run consecutivi con 0 nuovi lead. Si azzera appena arriva
  almeno un lead nuovo.
- `exhausted`: quando `zero_streak >= SATURATION_THRESHOLD` (default `2`).
- `total_new`: totale storico di lead trovati su questa coppia.
- `last_run`: data dell'ultimo run su questa coppia.

### Configurazione

In `config.py`:

- `REGIONI`: elenco regioni/province da coprire (Ticino + Lombardia + Piemonte).
- `SETTORI`: elenco settori (include anche `"qualsiasi"` come sweep finale).
- `SATURATION_THRESHOLD`: default `2` — quanti run consecutivi a zero servono
  per marcare esaurita una coppia. Alza a `3` se hai risposte API rumorose,
  abbassa a `1` per rotazione piu' aggressiva.

### Override manuale

Se lanci con `--regione` e/o `--settore` da CLI, **lo stato NON viene
aggiornato**: e' un run "fuori rotazione" per test o per un focus mirato.
Cosi' non sporchi la logica automatica quando fai esperimenti.

### Reset

Vuoi far ripartire tutto da zero (es. dopo aver cambiato la lista di
regioni/settori)? Basta cancellare `.rotation_state.json`:

```bash
rm .rotation_state.json
```

Se invece vuoi solo rimuovere il flag `exhausted` da alcune coppie (senza
resettare i contatori), puoi editare a mano il JSON.

---

## Esecuzione automatica notturna

### macOS / Linux (cron)

Apri il crontab:

```bash
crontab -e
```

Aggiungi una riga per lanciare lo script ogni notte alle 04:30
(**usa path assoluti** — cron non conosce il tuo `$PATH`):

```cron
30 4 * * * cd /Users/robypalmy/Desktop/cheap-software-house && /Users/robypalmy/Desktop/cheap-software-house/.venv/bin/python main.py >> logs/cron.log 2>&1
```

Note:
- Sostituisci il path con quello reale della tua cartella.
- Su macOS potrebbe essere necessario dare a `cron` (o meglio a `Terminal`
  / `iTerm`) i permessi di *Full Disk Access* in *Impostazioni di sistema →
  Privacy e sicurezza → Accesso completo al disco*, altrimenti la scrittura
  su alcune cartelle utente puo' essere bloccata.
- In alternativa a cron su macOS puoi usare `launchd` (piu' robusto per
  laptop che possono essere spenti la notte — vedi `man launchd.plist`).

### Windows (Task Scheduler)

1. Apri **Utilita' di pianificazione** (Task Scheduler).
2. **Crea attivita' di base** → nome: `Prospecting siti outdated`.
3. Trigger: **Ogni giorno** alle 04:30.
4. Azione: **Avvia programma**.
   - Programma/script:
     ```
     C:\path\al\progetto\.venv\Scripts\python.exe
     ```
   - Argomenti:
     ```
     main.py
     ```
   - Inizia in:
     ```
     C:\path\al\progetto
     ```
5. In *Condizioni*, spunta **Riattiva il computer per eseguire l'attivita'**
   se vuoi che parta anche con PC in sospensione.

---

## Output Excel

Il file `lead_prospecting.xlsx` ha una sola scheda `Lead` con queste colonne:

| Data_Trovato | Azienda | Settore | Comune | Provincia_Cantone | Sito_Web | Problemi_Rilevati | Email | Telefono | Fonte | Stato | Note |

- **Header**: grassetto, sfondo blu.
- **Data_Trovato**: data del run (ISO `YYYY-MM-DD`).
- **Problemi_Rilevati**: lista di problemi separati da `; ` (es. `non responsive; no HTTPS; design datato`).
- **Stato**: lista a discesa con `Da contattare`, `Email inviata`,
  `Contattato telefonicamente`, `Appuntamento fissato`, `Non interessato`.
  Stessa logica del file `Chrono_Tracker_pipeline_commerciale.xlsx`, cosi' in
  futuro i due file si possono unire facilmente.

Se apri il file in Excel mentre parte un run, lo script **non crasha**: logga
un errore chiaro (`Il file Excel ... sembra aperto in un'altra applicazione`)
e termina con exit code `3`. Chiudi Excel e rilancia manualmente.

---

## Come funziona in breve

1. **Rotazione stay-until-exhausted**: sceglie una coppia `(regione, settore)`
   dallo stato persistente e ci resta finche' produce nuovi lead (vedi
   sezione sopra). Con `--regione`/`--settore` da CLI si fa override senza
   toccare lo stato.
2. **Chiamata API**: manda un prompt a Claude con il tool `web_search`
   abilitato, chiedendo `N` PMI reali in quella zona con siti outdated.
3. **Parsing**: estrae l'array JSON dalla risposta (tollerando code fence),
   valida i campi obbligatori, scarta le righe malformate con warning nei log.
4. **Retry**: max 3 tentativi con backoff esponenziale su errori di rete,
   rate limit o JSON non parsabile.
5. **Deduplica**: confronta contro `storico_aziende.csv` su nome normalizzato
   (senza `S.r.l.`, `S.p.A.`, punteggiatura) e dominio del sito.
6. **Excel**: append delle sole aziende nuove, con formattazione header e
   data validation sulla colonna Stato.
7. **Storico**: aggiorna `storico_aziende.csv` con le nuove righe.
8. **Stato rotazione**: aggiorna `.rotation_state.json` — se ci sono nuovi
   lead, resta sulla coppia corrente; se sono zero, incrementa lo streak
   e, alla soglia, marca la coppia esaurita e avanza al prossimo settore
   (o regione, se erano tutti fatti).
9. **Log**: `logs/YYYY-MM-DD.log` con parametri usati, conteggi, errori.

---

## Note GDPR / etica

- Le fonti target sono **pubbliche**: siti aziendali, registroimprese.it,
  camere di commercio, pagine LinkedIn aziendali pubbliche, PagineGialle.
- I contatti sono **generali aziendali** (`info@…`, centralino), mai
  personali.
- Il prompt istruisce esplicitamente il modello a **non inventare** contatti:
  se una email/telefono non e' verificabile, il campo resta vuoto.
- Nessun data broker, nessun scraping aggressivo.

---

## Troubleshooting

- **`ANTHROPIC_API_KEY non impostata`** → controlla che `.env` esista nella
  root del progetto e contenga la chiave.
- **`Tutti i 3 tentativi falliti`** → probabile rate limit o problema di rete.
  Il file Excel non viene toccato. Riprova piu' tardi o abbassa `--numero`.
- **`Il file Excel ... sembra aperto`** → chiudi il file in Excel/Numbers e
  rilancia.
- **Cron non parte** → verifica che il path sia assoluto, che il `.venv`
  esista, e che su macOS `cron` abbia *Full Disk Access*.
- **Risultati di bassa qualita'** → prova ad abbassare `--numero` (batch piu'
  piccoli hanno hit rate migliore) o a specificare un `--settore` diverso.
