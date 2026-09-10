# Setup: GitHub-refreshed costs and shared Cloudflare bids

The two kinds of data use different services:

- **Floorplan costs:** the same GitHub Action and `FLOORPLAN_XLSX_URL` secret used by the Floorplan Cost Comparison dashboard.
- **Bids:** Cloudflare D1, so every coworker sees the same bids.

Power Automate and a SharePoint Cloudflare secret are not required.

## 1. Create the shared bids database

1. Open the Cloudflare dashboard.
2. Go to **Storage & databases → D1 SQL database**.
3. Select **Create database**.
4. Name it `trade-cost-dashboard` and create it.
5. Copy the database ID shown on the database page.
6. Open `wrangler.jsonc` in this repository and replace `PASTE_YOUR_D1_DATABASE_ID_HERE` with that ID. Keep the quotation marks.
7. On the database page, open **Console**.
8. Copy all of `migrations/0001_create_bids.sql` into the console and select **Execute**.

The D1 binding name is `DB`. Do not rename it.

## 2. Add the SharePoint workbook secret to this GitHub repository

GitHub secrets are stored per repository, so add the same secret to `Potential-Trade-Cost-Comparison` even if it already exists in `Floorplan-Cost-Comparison`.

1. In GitHub, open **Potential-Trade-Cost-Comparison → Settings**.
2. Select **Secrets and variables → Actions**.
3. Select **New repository secret**.
4. Name it exactly `FLOORPLAN_XLSX_URL`.
5. Paste the same SharePoint Excel sharing link used by the Floorplan Cost Comparison dashboard and save it.

GitHub does not reveal an existing secret's value. If necessary, copy the original SharePoint workbook link again.

## 3. Commit and deploy

Commit these files in GitHub Desktop and push to `main`. Suggested summary:

`Use GitHub refresh for costs and D1 for shared bids`

Cloudflare should build and deploy the commit automatically. Keep these build settings:

- Build command: `pnpm build`
- Deploy command: `npx wrangler deploy --config dist/server/wrangler.json`
- Root directory: `/`

## 4. Refresh costs

The workflow runs automatically once per hour. To refresh immediately:

1. Open **Actions → Refresh dashboard from SharePoint Excel** in the GitHub repository.
2. Select **Run workflow → Run workflow**.
3. Wait for the workflow to finish successfully.
4. Wait for Cloudflare to deploy the resulting commit, then reload the dashboard.

The dashboard's **Refresh via GitHub** button opens that workflow page. Like the existing Floorplan Cost Comparison dashboard, it does not directly run the workflow.

## 5. Verify shared bids

1. Save a distinctive test bid.
2. Open the dashboard in a private/incognito window or on a coworker's computer.
3. Confirm the same bid appears.
4. Delete the test bid and reload the other browser to confirm it is gone.

For company use, protect the Worker URL with Cloudflare Access so only approved coworkers can view or change bids.
