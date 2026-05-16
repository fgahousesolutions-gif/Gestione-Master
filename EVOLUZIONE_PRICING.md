# Evoluzione pricing appartamenti

Nota di continuita per riprendere il lavoro.

## Obiettivo principale

Costruire il modulo piu importante: un cruscotto quotidiano che, partendo dal file XLS delle prenotazioni, propone per ogni appartamento il prezzo da impostare data per data.

## Decisioni prese

- Si carica un XLS come nella webapp biancheria.
- Sono appartamenti solo i fogli il cui nome inizia con `ID00` seguito dal numero progressivo.
- I fogli che non iniziano con `ID00` sono fogli di servizio e devono essere ignorati.
- La app calcola ADR, occupazione, notti vendute, notti libere e buchi tra prenotazioni.
- Ogni appartamento ha parametri pricing salvati nel browser:
  - prezzo base
  - prezzo weekend
  - costo pulizie
  - extra
  - prezzo minimo
  - prezzo massimo
  - note
- La app presenta una tabella data per data con prezzo suggerito e motivo.
- Per ora non modifica automaticamente i portali: presenta azioni manuali.
- In locale l'XLS viene caricato a mano per ogni test; online andra salvato lato server come ultimo database valido.
- Online serviranno utenti/ruoli: admin completo, staff pulizie e proprietari con accesso limitato per email, appartamenti e moduli.

## Eventi locali

Prossimo passo complesso: integrare fonti serie per eventi locali. Idee citate:

- siti ufficiali eventi locali e turismo
- TicketOne o equivalenti per concerti/spettacoli
- venue ufficiali
- fiere, congressi, palazzetti, teatri, stadi

Da evitare fonti casuali o non verificabili. La prima versione puo mostrare eventi rilevanti e impatto stimato; la decisione finale resta manuale.

## Prezzo attualmente impostato

Tema aperto. Possibili strade:

- Se il file XLS contiene una colonna prezzo/tariffa, la app la legge.
- iCal di solito non contiene prezzi, solo occupazione.
- Per leggere prezzi reali dai calendari OTA serve esportazione prezzi, channel manager, oppure API/integrazione dedicata.
- In MVP si puo aggiungere una colonna nel database XLS con prezzo attuale per notte.

## Stato implementazione

Prima versione avviata in `server.py`:

- parsing valori economici
- output `pricing`
- card pricing per appartamento
- salvataggio parametri pricing in `localStorage`
- tabella 45 giorni con prezzo suggerito

## Online e utenti

Prima versione online da costruire:

- admin carica XLS e configura appartamenti
- XLS viene salvato lato server come ultimo schema valido
- utenti entrano con email e codice di verifica
- modello preferito login: email come username + codice a 6 cifre via email
- telefono resta opzionale come dato di contatto
- scheda admin utenti accesso: codice/nome, email, telefono, ruolo, attivo, appartamenti visibili, moduli visibili
- calendario visibile solo sugli appartamenti assegnati all'email autenticata
- staff pulizie vede solo lista interventi, senza upload, configurazione, pricing, biancheria o rendiconto
- utenti iniziali previsti: 4 persone staff pulizie
- la vista mobile pulizie deve essere la priorita, perche verra usata principalmente da smartphone
