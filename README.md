# Meteora DLMM Liquidity Pool Operations Analysis

## Project framing

This is a portfolio BI project built from Meteora’s public, read-only DLMM Data API on Solana. The goal is to treat decentralised liquidity pools as an operating product: compare liquidity capacity (TVL), demand (volume), fee monetisation, transaction activity and volatility instead of presenting price speculation.

## Deliverables

- `01_Raw_Pool_Snapshot.xlsx` — auditable pool-level snapshot and field notes.
- `02_Meteora_DLMM_Operations_Analysis.xlsx` — dashboard, pool benchmark, trend view and metric guide.
- `source_pool_snapshot.json` / `source_protocol_daily_volume.json` — exact API responses used by the workbooks.
- `build_project.py` — refreshes the above from the public API.

## Analysis questions

1. Which pools demonstrate high trading demand relative to liquidity supplied?
2. Does fee monetisation track liquidity utilisation, or are there pools requiring a deeper pricing / incentive review?
3. Is daily protocol volume stable enough for operating decisions, or does it need event-aware monitoring?

## Method notes

- **Volume / TVL** is a liquidity-utilisation proxy, not a return measure.
- **24h Fee Yield** is gross daily fees divided by point-in-time TVL. It is deliberately not annualised.
- Metrics are extracted as a dated snapshot. Re-run the script before presenting current numbers.
- This project is analytical and educational; it is not investment advice.

## Visualisation

Use the free Power BI Desktop workflow in `PowerBI_Setup_Guide.md` to create the two image-ready portfolio visuals: a pool-efficiency scatter plot and a daily-volume trend line.

## Resume-ready positioning

“Built a reproducible Web3 BI workflow using public Solana liquidity-pool data, translating TVL, volume, fee and transaction signals into operational questions for liquidity allocation and demand monitoring.”
