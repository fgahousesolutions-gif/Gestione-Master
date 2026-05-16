# Note per messa online

Questa cartella contiene la versione pulita da caricare su GitHub per il primo deploy.

## Avvio

In locale:

```bash
python3 server.py
```

Online:

```bash
HOST=0.0.0.0 PORT=8765 python3 server.py
```

Molti servizi impostano automaticamente `PORT`; in quel caso basta avviare `python3 server.py`.

Per Render e gia presente `render.yaml`, con runtime Python, avvio `python3 server.py`, `HOST=0.0.0.0` e versione Python fissata.

## Cosa funziona ora

- Caricamento file Excel `.xlsx`.
- Lettura soli fogli appartamento con nome che inizia per `ID00`.
- Moduli separati: Pricing, Calendario, Biancheria, Pulizie, Rendiconto.
- Configurazione appartamenti salvata nel browser.
- Scheda utenti accesso con email/codice, telefono, ruolo, appartamenti visibili, moduli visibili e note.
- Vista staff pulizie tramite `?staff=1` o `?view=pulizie`.
- Calendario prenotazioni con barre soggiorno tra check-in e check-out.

## Da fare prima dell'accesso reale staff

- Spostare XLS, configurazioni e utenti da `localStorage` a salvataggio server/database.
- Aggiungere login reale con email e codice di verifica.
- Definire ruoli: amministratore, pulizie, proprietario.
- Applicare lato server i permessi legati all'email su appartamenti e moduli visibili.
- Decidere dove salvare lo storico file caricati.
- Proteggere la vista pulizie: oggi il parametro `?staff=1` cambia solo la visualizzazione.

## Prima bozza deploy

Per andare online velocemente domani possiamo usare un servizio Python semplice come Render/Railway/Fly.
La versione attuale non richiede pacchetti esterni: `requirements.txt` e vuoto a parte una nota.
