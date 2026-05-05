from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


BASE_DIR = Path(__file__).parent
SAMPLE_DIR = BASE_DIR / "sample_data"

st.set_page_config(
    page_title="Tournament Finance Reconciliation Tool",
    page_icon=":bar_chart:",
    layout="wide",
)


def money(value: float) -> str:
    """Format Indian rupee values for KPI cards and tables."""
    if pd.isna(value):
        value = 0
    sign = "-" if value < 0 else ""
    return f"{sign}Rs. {abs(value):,.2f}"


def load_csv(uploaded_file, sample_file: str) -> pd.DataFrame:
    """Read an uploaded CSV, falling back to bundled sample data."""
    if uploaded_file is not None:
        return pd.read_csv(uploaded_file)
    return pd.read_csv(SAMPLE_DIR / sample_file)


def to_excel_download(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Create an in-memory Excel workbook from one or more dataframes."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
    return output.getvalue()


def clean_money_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Convert amount columns to numeric values while preserving blanks as NaN."""
    for column in columns:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce")
    return df


def build_reconciliation(
    platform_df: pd.DataFrame,
    payu_df: pd.DataFrame,
    organizer_df: pd.DataFrame,
    bank_df: pd.DataFrame,
):
    platform_df = platform_df.copy()
    payu_df = payu_df.copy()
    organizer_df = organizer_df.copy()
    bank_df = bank_df.copy()

    platform_df.columns = platform_df.columns.str.strip()
    payu_df.columns = payu_df.columns.str.strip()
    organizer_df.columns = organizer_df.columns.str.strip()
    bank_df.columns = bank_df.columns.str.strip()

    platform_df = clean_money_columns(
        platform_df,
        ["gross_amount", "platform_fee_rate", "tnsca_fee", "district_fee"],
    )
    payu_df = clean_money_columns(
        payu_df,
        ["gross_amount", "payu_fee_rate", "payu_fee_amount", "net_settled_amount"],
    )
    organizer_df = clean_money_columns(organizer_df, ["actual_payout_amount"])
    bank_df = clean_money_columns(bank_df, ["amount"])

    platform_df["missing_pass_through_fee_data"] = (
        platform_df[["tnsca_fee", "district_fee"]].isna().any(axis=1)
    )
    platform_df["tnsca_fee_calc"] = platform_df["tnsca_fee"].fillna(0)
    platform_df["district_fee_calc"] = platform_df["district_fee"].fillna(0)
    platform_df["platform_fee"] = (
        platform_df["gross_amount"] * platform_df["platform_fee_rate"]
    ).round(2)
    platform_df["pass_through_fees"] = (
        platform_df["tnsca_fee_calc"] + platform_df["district_fee_calc"]
    ).round(2)
    platform_df["expected_organizer_payout"] = (
        platform_df["gross_amount"]
        - platform_df["platform_fee"]
        - platform_df["pass_through_fees"]
    ).round(2)
    platform_df["duplicate_order_id"] = platform_df.duplicated("order_id", keep=False)
    platform_df["duplicate_transaction_id"] = platform_df.duplicated(
        "transaction_id", keep=False
    )

    payu_for_merge = payu_df.rename(
        columns={
            "gross_amount": "payu_gross_amount",
            "settlement_date": "payu_settlement_date",
        }
    )
    reconciled = platform_df.merge(payu_for_merge, on="order_id", how="left")
    reconciled["missing_payu_settlement"] = reconciled["settlement_id"].isna()
    reconciled["payu_fee_amount"] = reconciled["payu_fee_amount"].fillna(0)
    reconciled["payu_fee_impact"] = (
        reconciled["platform_fee"] - reconciled["payu_fee_amount"]
    ).round(2)
    reconciled["loss_making_transaction"] = (
        reconciled["payu_fee_amount"] > reconciled["platform_fee"]
    )
    reconciled["transaction_status"] = "Matched"
    reconciled.loc[
        reconciled["missing_payu_settlement"], "transaction_status"
    ] = "Missing PayU Settlement"
    reconciled.loc[
        reconciled["loss_making_transaction"], "transaction_status"
    ] = "Loss Making"

    expected_by_tournament = (
        platform_df.groupby(
            ["tournament_id", "tournament_name", "organizer_name"], dropna=False
        )
        .agg(
            transactions=("transaction_id", "count"),
            total_gross_amount=("gross_amount", "sum"),
            expected_organizer_payout=("expected_organizer_payout", "sum"),
            platform_fee=("platform_fee", "sum"),
            pass_through_fees=("pass_through_fees", "sum"),
        )
        .reset_index()
    )

    organizer_summary = (
        organizer_df.groupby(["tournament_id"], dropna=False)
        .agg(
            actual_payout_amount=("actual_payout_amount", "sum"),
            payout_count=("payout_id", "count"),
            payout_references=("bank_reference", lambda x: ", ".join(x.dropna())),
        )
        .reset_index()
    )

    payout_recon = expected_by_tournament.merge(
        organizer_summary, on="tournament_id", how="left"
    )
    payout_recon["actual_payout_amount"] = payout_recon[
        "actual_payout_amount"
    ].fillna(0)
    payout_recon["payout_variance"] = (
        payout_recon["actual_payout_amount"]
        - payout_recon["expected_organizer_payout"]
    ).round(2)

    def payout_status(row):
        if row["payout_count"] != row["payout_count"]:
            return "Missing Payout"
        if abs(row["payout_variance"]) <= 1:
            return "Matched"
        if row["payout_variance"] < 0:
            return "Short Paid"
        return "Overpaid"

    payout_recon["payout_status"] = payout_recon.apply(payout_status, axis=1)

    if not bank_df.empty and "reference" in bank_df.columns:
        payout_recon["bank_reference_found"] = payout_recon[
            "payout_references"
        ].fillna("").apply(
            lambda refs: any(
                ref.strip() in set(bank_df["reference"].astype(str))
                for ref in refs.split(",")
                if ref.strip()
            )
        )
    else:
        payout_recon["bank_reference_found"] = False

    total_platform_fee = reconciled["platform_fee"].sum()
    total_payu_fee = reconciled["payu_fee_amount"].sum()
    net_platform_profit = total_platform_fee - total_payu_fee
    partner_summary = pd.DataFrame(
        [
            {"metric": "Platform Fee Revenue", "amount": total_platform_fee},
            {"metric": "PayU Fees", "amount": total_payu_fee},
            {"metric": "Net Platform Profit / Loss", "amount": net_platform_profit},
            {"metric": "Karthikeyan 50% Share", "amount": net_platform_profit * 0.5},
            {"metric": "Partner 50% Share", "amount": net_platform_profit * 0.5},
        ]
    )

    exception_frames = []

    duplicate_orders = reconciled[reconciled["duplicate_order_id"]].copy()
    duplicate_orders["exception_type"] = "Duplicate order_id"
    exception_frames.append(duplicate_orders)

    duplicate_txns = reconciled[reconciled["duplicate_transaction_id"]].copy()
    duplicate_txns["exception_type"] = "Duplicate transaction_id"
    exception_frames.append(duplicate_txns)

    missing_payu = reconciled[reconciled["missing_payu_settlement"]].copy()
    missing_payu["exception_type"] = "Missing PayU settlement"
    exception_frames.append(missing_payu)

    loss_making = reconciled[reconciled["loss_making_transaction"]].copy()
    loss_making["exception_type"] = "Loss-making transaction"
    exception_frames.append(loss_making)

    missing_fees = reconciled[reconciled["missing_pass_through_fee_data"]].copy()
    missing_fees["exception_type"] = "Missing TNSCA/district fee data"
    exception_frames.append(missing_fees)

    payout_exceptions = payout_recon[
        payout_recon["payout_status"].isin(["Short Paid", "Overpaid", "Missing Payout"])
    ].copy()
    payout_exceptions["exception_type"] = payout_exceptions["payout_status"].map(
        {
            "Short Paid": "Organizer payout mismatch",
            "Overpaid": "Organizer payout mismatch",
            "Missing Payout": "Missing payout",
        }
    )

    exception_report = (
        pd.concat(exception_frames, ignore_index=True, sort=False)
        if exception_frames
        else pd.DataFrame()
    )

    return {
        "platform": platform_df,
        "payu": payu_df,
        "organizer": organizer_df,
        "bank": bank_df,
        "reconciled": reconciled,
        "payout_recon": payout_recon,
        "payout_exceptions": payout_exceptions,
        "exception_report": exception_report,
        "partner_summary": partner_summary,
    }


def kpi_card(label: str, value: str, help_text: str | None = None):
    st.metric(label, value, help=help_text)


def dataframe_with_money(df: pd.DataFrame):
    st.dataframe(df, use_container_width=True, hide_index=True)


st.title("Tournament Finance Reconciliation Tool")
st.caption(
    "Demo dashboard for reconciling tournament registrations, PayU settlements, "
    "organizer payouts, and partner profit sharing."
)

with st.sidebar:
    st.header("Data Source")
    st.write(
        "Upload CSV files to replace the built-in sample data. Leave uploads blank "
        "to run the demo with bundled sample files."
    )
    platform_upload = st.file_uploader("Platform transactions CSV", type=["csv"])
    payu_upload = st.file_uploader("PayU settlements CSV", type=["csv"])
    organizer_upload = st.file_uploader("Organizer payouts CSV", type=["csv"])
    bank_upload = st.file_uploader("Bank statement CSV", type=["csv"])

try:
    data = build_reconciliation(
        load_csv(platform_upload, "platform_transactions.csv"),
        load_csv(payu_upload, "payu_settlements.csv"),
        load_csv(organizer_upload, "organizer_payouts.csv"),
        load_csv(bank_upload, "bank_statement.csv"),
    )
except Exception as exc:
    st.error(f"Could not load or reconcile the files: {exc}")
    st.stop()

reconciled = data["reconciled"]
payout_recon = data["payout_recon"]
partner_summary = data["partner_summary"]
exception_report = data["exception_report"]
payout_exceptions = data["payout_exceptions"]

total_gross = reconciled["gross_amount"].sum()
total_platform_fee = reconciled["platform_fee"].sum()
total_payu_fee = reconciled["payu_fee_amount"].sum()
net_profit = total_platform_fee - total_payu_fee
loss_count = int(reconciled["loss_making_transaction"].sum())
duplicate_count = int(
    reconciled["duplicate_order_id"].sum()
    + reconciled["duplicate_transaction_id"].sum()
)
organizer_variance = payout_recon["payout_variance"].sum()
karthikeyan_share = net_profit * 0.5
partner_share = net_profit * 0.5

tabs = st.tabs(
    [
        "Overview Dashboard",
        "Transaction Reconciliation",
        "PayU Fee Impact",
        "Organizer Payout Reconciliation",
        "Partner Profit Summary",
        "Exception Report",
    ]
)

with tabs[0]:
    st.subheader("Overview Dashboard")
    st.write(
        "This page summarizes collections, platform revenue, PayU fees, payout "
        "variance, and partner profit sharing."
    )

    row1 = st.columns(4)
    with row1[0]:
        kpi_card("Total Gross Collections", money(total_gross))
    with row1[1]:
        kpi_card("Platform Fee Revenue", money(total_platform_fee))
    with row1[2]:
        kpi_card("PayU Fees", money(total_payu_fee))
    with row1[3]:
        kpi_card("Net Platform Profit / Loss", money(net_profit))

    row2 = st.columns(4)
    with row2[0]:
        kpi_card("Loss-Making Transactions Count", f"{loss_count:,}")
    with row2[1]:
        kpi_card("Duplicate Records Count", f"{duplicate_count:,}")
    with row2[2]:
        kpi_card("Organizer Payout Variance", money(organizer_variance))
    with row2[3]:
        kpi_card("Karthikeyan 50% Share", money(karthikeyan_share))

    row3 = st.columns(4)
    with row3[0]:
        kpi_card("Partner 50% Share", money(partner_share))

    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        fee_chart = partner_summary[
            partner_summary["metric"].isin(["Platform Fee Revenue", "PayU Fees"])
        ]
        fig = px.bar(
            fee_chart,
            x="metric",
            y="amount",
            text="amount",
            title="Platform Fee vs PayU Fees",
            color="metric",
        )
        fig.update_traces(texttemplate="Rs. %{text:,.0f}", textposition="outside")
        fig.update_layout(showlegend=False, yaxis_title="Amount")
        st.plotly_chart(fig, use_container_width=True)

    with chart_col2:
        status_counts = payout_recon["payout_status"].value_counts().reset_index()
        status_counts.columns = ["payout_status", "count"]
        fig = px.pie(
            status_counts,
            names="payout_status",
            values="count",
            title="Organizer Payout Status",
            hole=0.35,
        )
        st.plotly_chart(fig, use_container_width=True)

with tabs[1]:
    st.subheader("Transaction Reconciliation")
    st.write(
        "Each platform transaction is matched to PayU using order_id. The table "
        "shows expected organizer payout, PayU settlement status, and duplicate flags."
    )

    status_filter = st.multiselect(
        "Filter transaction status",
        sorted(reconciled["transaction_status"].dropna().unique()),
        default=sorted(reconciled["transaction_status"].dropna().unique()),
    )
    filtered = reconciled[reconciled["transaction_status"].isin(status_filter)]
    dataframe_with_money(filtered)

    st.download_button(
        "Download reconciled transaction report as Excel",
        data=to_excel_download({"Reconciled Transactions": reconciled}),
        file_name="reconciled_transaction_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tabs[2]:
    st.subheader("PayU Fee Impact")
    st.write(
        "PayU fees are deducted on the full collected amount. When PayU fees are "
        "higher than the 3% platform fee, the transaction becomes loss-making."
    )

    fee_impact = reconciled[
        [
            "transaction_id",
            "order_id",
            "tournament_id",
            "gross_amount",
            "platform_fee",
            "payu_fee_amount",
            "payu_fee_impact",
            "loss_making_transaction",
        ]
    ].copy()
    dataframe_with_money(fee_impact)

    fig = px.bar(
        fee_impact,
        x="order_id",
        y="payu_fee_impact",
        color="loss_making_transaction",
        title="PayU Fee Impact by Order",
        labels={
            "payu_fee_impact": "Platform Fee - PayU Fee",
            "loss_making_transaction": "Loss Making",
        },
    )
    fig.add_hline(y=0, line_dash="dash", line_color="gray")
    st.plotly_chart(fig, use_container_width=True)

with tabs[3]:
    st.subheader("Organizer Payout Reconciliation")
    st.write(
        "Expected organizer payout is grouped by tournament and compared with the "
        "actual payout file. Status is Matched, Short Paid, Overpaid, or Missing Payout."
    )

    dataframe_with_money(payout_recon)
    fig = px.bar(
        payout_recon,
        x="tournament_id",
        y="payout_variance",
        color="payout_status",
        title="Organizer Payout Variance by Tournament",
        hover_data=["tournament_name", "organizer_name"],
    )
    fig.add_hline(y=0, line_dash="dash", line_color="gray")
    st.plotly_chart(fig, use_container_width=True)

    st.download_button(
        "Download organizer payout reconciliation as Excel",
        data=to_excel_download({"Organizer Payout Reconciliation": payout_recon}),
        file_name="organizer_payout_reconciliation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tabs[4]:
    st.subheader("Partner Profit Summary")
    st.write(
        "Net platform profit is platform fee revenue minus PayU fees. The result is "
        "split equally between Karthikeyan and the partner."
    )

    summary_display = partner_summary.copy()
    summary_display["formatted_amount"] = summary_display["amount"].apply(money)
    dataframe_with_money(summary_display)

    fig = px.bar(
        partner_summary,
        x="metric",
        y="amount",
        text="amount",
        title="Profit Sharing Summary",
    )
    fig.update_traces(texttemplate="Rs. %{text:,.0f}", textposition="outside")
    st.plotly_chart(fig, use_container_width=True)

    st.download_button(
        "Download partner profit summary as Excel",
        data=to_excel_download({"Partner Profit Summary": partner_summary}),
        file_name="partner_profit_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tabs[5]:
    st.subheader("Exception Report")
    st.write(
        "This report highlights records requiring review, including duplicates, "
        "missing PayU settlement, PayU losses, fee data gaps, and payout exceptions."
    )

    st.markdown("**Transaction Exceptions**")
    if exception_report.empty:
        st.success("No transaction exceptions found.")
    else:
        exception_types = sorted(exception_report["exception_type"].dropna().unique())
        selected_exceptions = st.multiselect(
            "Filter exception type",
            exception_types,
            default=exception_types,
        )
        dataframe_with_money(
            exception_report[
                exception_report["exception_type"].isin(selected_exceptions)
            ]
        )

    st.markdown("**Organizer Payout Exceptions**")
    if payout_exceptions.empty:
        st.success("No organizer payout exceptions found.")
    else:
        dataframe_with_money(payout_exceptions)

    st.download_button(
        "Download exception report as Excel",
        data=to_excel_download(
            {
                "Transaction Exceptions": exception_report,
                "Payout Exceptions": payout_exceptions,
            }
        ),
        file_name="exception_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
