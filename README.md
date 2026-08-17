# EGX AI Scanner

An automated technical-analysis scanner for the Egyptian Exchange (EGX). It pulls daily price data, evaluates each stock against a rule-based **SMA200 + Breakout/Pullback** strategy, scores the candidates, and pushes the results to Telegram (with an Excel export) on a daily schedule after market close.

## Features

- **Trend filter** — SMA200 to confirm the primary trend direction
- **Momentum filter** — RSI(14) with configurable overbought/oversold bounds
- **Volatility-based risk management** — ATR-derived stop-loss and take-profit levels
- **Volume confirmation** — flags breakouts backed by above-average volume
- **Trend strength filter** — ADX threshold to avoid choppy/sideways markets
- **Pullback mode** — optional detection of pullback-to-SMA200 setups
- **Automated scheduling** — runs daily after the EGX closing bell (Cairo time)
- **Telegram delivery** — sends a summary message plus a full Excel report

## Tech stack

Python · pandas · numpy · yfinance · python-telegram-bot · openpyxl · schedule

## Setup

1. Clone the repo and install dependencies:
   ```bash
   git clone https://github.com/<your-username>/egx-ai-scanner.git
   cd egx-ai-scanner
   pip install -r requirements.txt
   ```

2. Create your local environment file:
   ```bash
   cp .env.example .env
   ```
   Then open `.env` and fill in your own Telegram bot token and chat ID.
   (Create a bot via [@BotFather](https://t.me/BotFather) to get a token.)

3. Run a one-off scan:
   ```bash
   python egx_scanner.py
   ```

4. Or run the daily scheduler (fires at 15:20 Cairo time on trading days):
   ```bash
   python scheduler.py
   ```

## Configuration

All strategy parameters (SMA length, RSI bounds, ATR multipliers, volume threshold, ADX threshold, etc.) live in `config.py` and can be tuned without touching the scanning logic.

## Disclaimer

This project is for educational and informational purposes only and does not constitute financial advice. Trading involves risk; always do your own research before making investment decisions.

## License

MIT
