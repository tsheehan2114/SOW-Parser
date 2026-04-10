# SOW Parser

Extracts financial line items from SOW PDFs for NetSuite entry. Uses a two-pass Claude AI extraction and exports a structured Excel file with 4 tabs: SOW Extract, Billing Periods, SKU Reference, and Audit Trail.

## Tech Stack

- React + Vite (frontend)
- Vercel (hosting + serverless API route)
- Anthropic Claude API (extraction)
- SheetJS (Excel export)

## Local Development

1. Install [Vercel CLI](https://vercel.com/docs/cli): `npm i -g vercel`
2. Clone the repo and install deps: `npm install`
3. Copy env file: `cp .env.example .env.local` and add your API key
4. Run locally with Vercel CLI (required for `/api` routes): `vercel dev`
5. Open http://localhost:3000

## Deployment

1. Push code to GitHub
2. Connect repo in [Vercel Dashboard](https://vercel.com/dashboard)
3. Add `ANTHROPIC_API_KEY` in Vercel → Project Settings → Environment Variables
4. Vercel auto-deploys on every push to `main`

## How the API Proxy Works

`/api/claude.js` is a Vercel serverless function. The browser calls `/api/claude` instead of Anthropic directly. Vercel injects the `ANTHROPIC_API_KEY` server-side, so it's never exposed in the browser bundle.

## SOWParser.jsx — One Required Change

In `SOWParser.jsx`, find any line calling the Anthropic API directly:
```
https://api.anthropic.com/v1/messages
```
Replace it with:
```
/api/claude
```
(There should be two calls — one for extract, one for verify.)
