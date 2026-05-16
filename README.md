# Gestione appartamenti

Mini app locale per caricare il file Excel delle prenotazioni e lavorare su moduli separati:

- pricing
- biancheria
- pulizie
- rendiconto

## Avvio locale

Serve Python 3, senza librerie esterne.

```bash
python3 server.py
```

Poi apri:

```text
http://localhost:8765
```

## Dati salvati

La app salva nel browser:

- appartamenti registrati
- parametri per appartamento
- consegne inserite
- utenti accesso configurati

Se apri da un altro dispositivo, quei dati non vengono trasferiti automaticamente perche sono salvati nel browser locale.

## Nota deploy online

Questa versione usa un backend Python (`server.py`). Va bene per uso locale o per hosting Python.

Per andare online con utenti pulizie serve una versione con backend persistente:

- login utenti
- ruoli: admin, staff pulizie e proprietario
- login staff con email e codice di verifica via email
- salvataggio lato server dell'ultimo XLS valido
- salvataggio lato server delle configurazioni appartamenti
- permessi per email su appartamenti e moduli visibili
- vista staff limitata al planning pulizie

La strada consigliata e convertire questa app in una piccola webapp online con backend e storage. Cloudflare Pages statico da solo non basta se vogliamo login, XLS salvato e dati condivisi.

## File deploy inclusi

- `Procfile`: avvio su hosting Python compatibili
- `render.yaml`: configurazione pronta per Render
- `requirements.txt`: nessuna dipendenza esterna per ora
- `.env.example`: variabili ambiente di esempio
- `DEPLOY_NOTES.md`: promemoria tecnico per la messa online
