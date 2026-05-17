# Gestione appartamenti - Cloudflare

Versione destinata al deploy definitivo su Cloudflare Workers.

## Architettura

- Python Worker: backend e API
- Assets statici: interfaccia web
- D1: utenti, permessi, impostazioni
- R2: ultimo XLS e archivio caricamenti
- Cloudflare Access: login email con OTP

## Comandi previsti

```bash
npm install
npx wrangler login
uv run pywrangler dev
```

Quando la configurazione Cloudflare e pronta:

```bash
npx wrangler login
npx wrangler d1 create gestione-appartamenti
npx wrangler r2 bucket create gestione-appartamenti-uploads
npm run db:migrate:remote
npm run deploy
```

## Stato

Questa cartella contiene gia.:

- UI completa
- parser XLS migrato dentro un Python Worker
- salvataggio dell'ultimo XLS in R2
- salvataggio dell'ultimo stato elaborato in D1
- recupero automatico dell'ultimo stato all'apertura
- filtraggio lato server per appartamenti visibili
- visibilita moduli lato interfaccia in base ai permessi
- upload XLS riservato agli admin

L'autenticazione email va configurata in Cloudflare Access con One-time PIN e policy sugli indirizzi autorizzati.

## Admin iniziale

- `fgahousesolutions@gmail.com`
