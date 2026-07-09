# SKY REI · Portfolio Reports

Live P&L dashboard for management. Served via Cloudflare Pages; access gated by Cloudflare Access.

## Files
- `index.html` — redirects to the P&L dashboard
- `PL Story - Main.dc.html` — the dashboard
- `pl-data.json` — the data (REPLACED on every publish; generated from PL.xlsx by publish_pl.py — never hand-edited)
- `support.js` — dashboard runtime

## Monthly update
1. Update PL.xlsx locally, save.
2. Ask Claude to "publish" — validation gates run; you receive the new pl-data.json.
3. In this repo: Add file → Upload files → drop pl-data.json → Commit.
4. Cloudflare redeploys automatically (~1 minute).

## Rollback
Repo → Commits → open the bad commit → Revert. Cloudflare redeploys the previous data.
