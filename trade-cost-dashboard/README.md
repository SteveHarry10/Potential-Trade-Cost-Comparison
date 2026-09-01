# Trade Cost Dashboard

A responsive bid-comparison dashboard modeled on the supplied Builder Bid Desk snapshot. It matches potential trade bids to the latest Floorplans workbook row by **plan + cost code**, then recalculates the dollar and percentage variance whenever costs are refreshed.

## Included

- Fast bid entry matching the supplied Plan → Trade → Bidder → Amount/Date layout
- Complete Floorplans snapshot: 47 plans and 3,383 current plan/cost-code combinations
- Plan, trade, and keyword filtering
- Bidder-level average variance rollups
- Current cost / bid / variance / percentage comparison table
- Device-local bid persistence
- Server-side SharePoint/Excel connector so the workbook URL or token never ships to the browser
- Responsive dark UI matching the reference image

## Run locally

Requirements: Node.js 22.13+ and pnpm.

```bash
pnpm install
cp .env.example .env.local
pnpm dev
```

Open `http://localhost:3000`.

## Connect the SharePoint Floorplans workbook

The repository includes the complete current `Floorplans` sheet from **Master Costing Sheet Live** as its initial dataset. The supplied SharePoint view link is session-based, so a deployed server cannot safely reuse the browser session behind that link. For automatic refreshes, expose the existing `FloorplanCosts` table through a secure Microsoft 365 endpoint. A simple approach is a Power Automate flow:

1. Use the workbook's existing `FloorplanCosts` table stored in SharePoint.
2. Create a flow with an HTTP request trigger (or another authenticated HTTP front door).
3. Add **Excel Online (Business) → List rows present in a table** and choose the SharePoint site, workbook, and Floorplans table.
4. Return the `value` array in the response.
5. Put the generated endpoint in `SHAREPOINT_EXCEL_ENDPOINT` in `.env.local` or the host's secret settings. If the endpoint expects a bearer token, also set `SHAREPOINT_EXCEL_TOKEN`.

The connector accepts either a JSON array or `{ "value": [...] }`. It understands both the workbook's native wide layout (`Code`, `Item`, then one column per plan) and a normalized row layout with these names:

| Dashboard field | Accepted workbook headers |
| --- | --- |
| Plan | `plan`, `Plan`, `Floorplan` |
| Cost code | `code`, `Code`, `Cost Code` |
| Trade | `trade`, `Trade`, `Category` |
| Description | `description`, `Description`, `Item` |
| Current cost | `current`, `Current`, `Cost`, `Amount` |

After changing costs in Excel, click **Refresh costs**. Every existing bid is re-matched to its latest plan + cost code and the comparison updates immediately.

## Production

```bash
pnpm build
pnpm start
```

Keep `.env.local` out of Git. Only `.env.example` is included in the repository.
