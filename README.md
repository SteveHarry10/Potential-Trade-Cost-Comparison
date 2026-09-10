# Trade Cost Dashboard

A responsive trade-bid comparison dashboard modeled on the supplied Builder Bid Desk snapshot. It matches potential bids to the latest Floorplans workbook row by **plan + cost code**.

## How data works

- **Current costs:** generated from the SharePoint workbook by a GitHub Action, using the same `FLOORPLAN_XLSX_URL` secret and parsing script as the existing Floorplan Cost Comparison dashboard.
- **Shared bids:** stored in Cloudflare D1 instead of browser storage, so coworkers see the same bids.
- **Updated comparisons:** each displayed bid is re-matched to the newest plan + cost code after the cost data is refreshed and deployed.

Power Automate and a SharePoint Cloudflare secret are not required.

## Included

- Fast bid entry using the requested Plan → Trade → Bidder → Amount/Date layout
- All populated cost lines from the Floorplans worksheet
- Plan, trade, and keyword filters
- Bidder-level average variance rollups
- Current cost, bid, variance, and percentage comparison
- Shared bid persistence through Cloudflare D1
- Hourly and manual SharePoint cost refresh through GitHub Actions
- Responsive dark interface based on the supplied reference image

## Setup

Follow `CLOUDFLARE-SETUP.md`. In short:

1. Create and bind the D1 database, then create its bids table.
2. Add the GitHub repository secret `FLOORPLAN_XLSX_URL` with the SharePoint Excel sharing link.
3. Push to `main` and allow Cloudflare to deploy.
4. Run **Actions → Refresh dashboard from SharePoint Excel** whenever an immediate cost update is needed.

## Run locally

Requirements: Node.js 22.13+ and pnpm.

```bash
pnpm install
pnpm dev
```

Open `http://localhost:3000`.

## Production build

```bash
pnpm build
```
