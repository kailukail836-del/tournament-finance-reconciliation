# Tournament Finance Reconciliation Tool

A demo Streamlit finance reconciliation dashboard for an online tournament platform.

The app reconciles platform registration transactions with PayU settlements, organizer payouts, and a bank statement. It also highlights duplicate records, missing settlements, PayU fee losses, missing pass-through fee data, and organizer payout mismatches.

## Business Rules

- Platform fee = `gross_amount * platform_fee_rate`
- Pass-through fees = `tnsca_fee + district_fee`
- Expected organizer payout = `gross_amount - platform_fee - tnsca_fee - district_fee`
- Transactions are matched to PayU settlements by `order_id`
- PayU fee impact = `platform_fee - payu_fee_amount`
- Transactions are loss-making when `payu_fee_amount > platform_fee`
- Net platform profit = total platform fees - total PayU fees
- Karthikeyan and partner share net platform profit 50/50

## Project Files

- `app.py`
- `requirements.txt`
- `sample_data/platform_transactions.csv`
- `sample_data/payu_settlements.csv`
- `sample_data/organizer_payouts.csv`
- `sample_data/bank_statement.csv`

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Using the App

The app loads the bundled sample data by default. Use the sidebar upload controls to replace any sample file with your own CSV.

Required CSV columns are shown in the sample files. The sample data includes multiple tournaments, duplicate order IDs, duplicate transaction IDs, PayU loss-making transactions, TNSCA and district fees, missing fee data, payout mismatches, and a missing payout.

## Dashboard Tabs

1. Overview Dashboard
2. Transaction Reconciliation
3. PayU Fee Impact
4. Organizer Payout Reconciliation
5. Partner Profit Summary
6. Exception Report

Each report can be downloaded as an Excel workbook from its relevant tab.
