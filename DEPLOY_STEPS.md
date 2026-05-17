# Deploy Cloudflare

## 1. Repository GitHub

Carica nel repository GitHub il contenuto di questa cartella, non la cartella contenitore.

## 2. Risorse Cloudflare

Da terminale, dentro la cartella progetto:

```bash
npx wrangler login
npx wrangler d1 create gestione-appartamenti
npx wrangler r2 bucket create gestione-appartamenti-uploads
```

Dopo il comando D1, copia `database_id` dentro `wrangler.jsonc` al posto di `replace-after-create`.

## 3. Database

```bash
npm run db:migrate:remote
```

Questo crea tabelle e admin iniziale:

- `fgahousesolutions@gmail.com`

## 4. Deploy

```bash
npm run deploy
```

Cloudflare restituira un URL `workers.dev`.

## 5. Login via email

Nel pannello Cloudflare:

1. Apri `Zero Trust`.
2. Vai in `Integrations` > `Identity providers`.
3. Aggiungi `One-time PIN`.
4. Crea una applicazione `Access` per l'URL del Worker.
5. Consenti solo le email autorizzate, iniziando da `fgahousesolutions@gmail.com`.

## 6. Dopo il primo accesso

Carica il primo XLS come admin. Da quel momento:

- l'ultimo XLS resta salvato in R2
- l'ultimo stato elaborato resta salvato in D1
- gli altri utenti vedranno solo gli appartamenti e i moduli autorizzati
