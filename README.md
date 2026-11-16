# Imdath · Outlook Drafting Assistant

An Outlook add-in for drafting replies with conversation context, CRM information, follow-up support and per-user writing-style features. A Node.js/Express backend connects the add-in to Microsoft Graph, HubSpot and an Anthropic model.

**Status: development prototype.** This portfolio version uses placeholder Azure registration IDs and excludes organization deployment instructions, internal security notes and private writing-style examples. It is intended for source review and development with your own test tenant.

## Architecture

- `public/`: served Office add-in interface, authentication dialog and client-side modules.
- `server/`: Express API, model requests, Microsoft Graph/CRM adapters, token storage and subscription handling.
- `manifest.xml`: Office add-in registration and development URLs.
- `tests/`: Jest service, API, frontend and security regression tests.
- Root client files are retained as source counterparts; the server serves `public/` rather than the repository root.

## Development setup

Use Node.js 22, npm, an Outlook development account and your own Azure development application. Real email/CRM accounts and model requests may involve external data and charges.

```sh
npm ci
cp .env.example .env
```

Fill in only your own development service settings. Replace the zero GUIDs in `config.js` and `public/config.js` with your development Azure client/tenant IDs, and configure that app's redirects and delegated Microsoft Graph permissions. Review the existing scope list before granting access.

Set the manifest and development application URLs consistently for your environment; `example.invalid` addresses are intentionally nonfunctional placeholders. Do not use a real company's deployment or tenant during review.

Office add-ins require local HTTPS. `npm run install-certs` installs a local development certificate and can prompt for operating-system trust approval. Run it deliberately when setting up the add-in:

```sh
npm run install-certs
npm start
```

The default development URL is <https://localhost:3000>. Sideload the reviewed development manifest into your test Outlook environment.

## Validation

```sh
npm test -- --runInBand
```

Tests and local code checks do not establish live Outlook/CRM compatibility or production readiness. No real email is sent and no production tenant is connected as part of portfolio preparation. Keep credentials, message content and token stores outside Git.
