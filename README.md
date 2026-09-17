# Meteora DLMM Liquidity Pool Operations Analysis

An independent Web3 business-intelligence portfolio project analysing Solana liquidity-pool operations using public Meteora DLMM API data, Excel and Power BI Desktop.

---

## 1. Project Overview

**Role:** Data Analyst | **Tools:** Power BI Desktop, Python, Excel, Public API  
**Objective:** Analysed liquidity-pool TVL, 24-hour trading volume, fees, transaction activity and protocol-level trends to identify liquidity-utilisation and fee-monetisation questions for product and operations teams.

> **Note:** This project uses point-in-time public API data. Metrics are operational diagnostics, not investment advice or performance forecasts.

### 📂 Full Report and Supporting Files

- **English Report:** [Meteora DLMM Operations Analysis Report](03_Meteora_DLMM_Operations_Analysis_Report.docx)
- **Pool Snapshot:** [01_Raw_Pool_Snapshot.xlsx](01_Raw_Pool_Snapshot.xlsx)
- **Operations Analysis Workbook:** [02_Meteora_DLMM_Operations_Analysis.xlsx](02_Meteora_DLMM_Operations_Analysis.xlsx)
- **Power BI Trend Dataset:** [PowerBI_Trend_Ready.xlsx](PowerBI_Trend_Ready.xlsx)
- **Power BI Build Guide:** [PowerBI_Setup_Guide.md](PowerBI_Setup_Guide.md)

## 🔗 Key Insights

- **Liquidity utilisation:** Compared 24-hour trading volume against TVL to distinguish pools with strong demand from pools where liquidity may be under-utilised.
- **Fee monetisation:** Used 24-hour fee yield alongside volume and transaction activity to identify pools requiring a deeper review of fee tiers, liquidity incentives or trading composition.
- **Volatility-aware monitoring:** Used daily protocol volume and a 7-day moving average to separate sustained demand patterns from single-day volume spikes.

## 📊 Dashboard Preview

### Pool Efficiency Segmentation

TVL and 24-hour volume position each pool by liquidity capacity and trading demand. Colour segments indicate operating-review priorities rather than return recommendations.

<img width="1956" height="1191" alt="pool-efficiency-scatter" src="https://github.com/user-attachments/assets/54743521-35b4-4da6-b29b-8b54c8127a5b" />


### Protocol Activity Trend

Daily protocol volume is shown with a 7-day moving average to make demand volatility visible without over-interpreting individual spikes.

![Uploading protocol-volume-trend.png…]()


## Methodology

| Metric | Definition | Business use |
| --- | --- | --- |
| TVL | Current USD value of liquidity displayed by the API | Liquidity capacity |
| Volume / TVL | 24-hour trading volume divided by current TVL | Liquidity-utilisation proxy |
| Fee Yield (24h) | 24-hour gross fees divided by current TVL | Fee-monetisation intensity |
| 7-day moving average | Rolling average of daily protocol volume | Trend monitoring |

## Data and Reproducibility

- **Source:** Meteora DLMM public read-only API endpoints: `/pools` and `/stats/daily/volume`.
- `source_pool_snapshot.json` and `source_protocol_daily_volume.json` retain the API responses used by the analysis.
- Run `build_project.py` to refresh the source snapshot and Excel workbooks.
- `build_powerbi_trend_source.mjs` creates the Power BI-ready trend file with literal dates and pre-calculated 7-day averages.

## Repository Structure

```text
├── assets/                                      # README dashboard images
├── 01_Raw_Pool_Snapshot.xlsx                    # pool-level source snapshot
├── 02_Meteora_DLMM_Operations_Analysis.xlsx     # analysis workbook
├── 03_Meteora_DLMM_Operations_Analysis_Report.docx
├── PowerBI_Trend_Ready.xlsx                     # clean import file for Power BI
├── PowerBI_Setup_Guide.md
├── build_project.py
├── build_powerbi_trend_source.mjs
└── source_*.json
```
