# CMC Invest Screener (Daily Stock Ratings)

Rates the stocks on CMC Invest from **Strong Buy** to **Strong Sell** using 20 technical signals (trend, momentum, volume) and up to 22 fundamental signals (P/E and other valuation ratios vs sector peers, quality, growth, financial health, analyst sentiment), then publishes an interactive screener page.

## The screener page

The page is rebuilt automatically every weekday morning at about **9am Melbourne time** (8am in winter), after the US close and before the ASX opens, and published to this repo's GitHub Pages site (see the link under **About** on the repo page, or **Settings → Pages**).

## Run it now or change the markets

1. Open the **Actions** tab.
2. Pick **Update CMC screener** on the left.
3. Click **Run workflow**. Leave markets as `au us` (ASX + US), or enter codes such as `au us gb ca jp`, or `all` for all 16 CMC Invest markets (a much bigger, slower page).

Each run also keeps the colour-coded Excel file and CSV for 14 days: open the run in the **Actions** tab and download **cmc-ratings** at the bottom.

To change the daily schedule or the default markets, edit `.github/workflows/screener.yml`.

## Run it on your own computer

On a Mac, double-click **Run Screener.command** (it sets itself up the first time). Or from a terminal:

```
pip install -r requirements.txt
python cmc_screener.py --markets au us
```

It writes `screener_output/site/index.html` (the screener page, opened in your browser when done), plus the Excel file and CSV. Run `python cmc_screener.py --help` for all options.

## Notes

- Data comes from Yahoo Finance via the free `yfinance` library, filtered to the exchanges CMC Invest offers. It is unofficial and can have gaps.
- Company financials are cached for 7 days between runs, so most daily runs only refresh prices.
- Not financial advice. The rating is a mechanical summary of indicators: use it to shortlist, then do your own research.
