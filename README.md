# OUTLOOK VAULT
Private Outlook address organizer. React + TS + Vite + Tailwind 4 + Dexie (IndexedDB) + SheetJS. Spreadsheets are parsed in-browser and never stored or uploaded; only extracted emails + import metadata are saved.

## Run locally
`npm install && npm run dev`

## Deploy to Cloudflare Pages
1. Create a GitHub repo and push this folder.
2. Cloudflare Dashboard → Workers & Pages → Create → Pages → Connect to Git.
3. Pick the repo and `main` branch.
4. Build command: `npm run build` · Output directory: `dist`.
5. Deploy. (`public/_redirects` provides SPA routing.)

## Make it private (required)
The in-app password is only a convenience lock; data sits in each visitor's own browser. For real access control:
Cloudflare Zero Trust → Access → Applications → Add → Self-hosted → enter your `*.pages.dev` (or custom) domain → add a policy allowing only your email (one-time PIN). Zero Trust has a free tier for small teams (up to 50 users); custom domains/extra features may be paid.
Also protect preview deployments (`*.project.pages.dev`) with the same Access app.

## Notes (v1.1.0)
- Data is stored in IndexedDB (`OutlookVaultDB`) of the current browser profile and device only; it is never sent anywhere. Back up regularly (Settings → Backup JSON).
- Apply the Cloudflare Access policy to both the production domain and `*.pages.dev` preview URLs.

## Tests
`npm test` runs `test/run.ts` (Node + fake-indexeddb): import, duplicate-only import, backup/restore, sorting by import date, persistence. UI behaviour (search, filters, pagination, layout) is not covered by automated tests.
