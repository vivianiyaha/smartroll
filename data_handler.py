import pandas as pd
import numpy as np

REQUIRED_COLUMNS = [
    "STAFF NAME", "BASIC", "HOUSING", "TRANSPORT",
    "TAX", "PENSION", "LOAN", "SAL. ADV.", "PENALTY",
    "TOTAL DED.", "NET SALARY",
]

EARNING_COLS   = ["BASIC", "HOUSING", "TRANSPORT"]
DEDUCTION_COLS = ["TAX", "PENSION", "LOAN", "SAL. ADV.", "PENALTY"]


def load_payroll_data(file) -> pd.DataFrame:
    """Read a single Excel file, clean it, and return a DataFrame."""
    df = pd.read_excel(file, engine="openpyxl")
    df.columns = [str(c).strip().upper() for c in df.columns]
    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)

    if "STAFF NAME" in df.columns:
        df["STAFF NAME"] = df["STAFF NAME"].astype(str).str.strip()
        df = df[df["STAFF NAME"].notna()
                & (df["STAFF NAME"] != "")
                & (df["STAFF NAME"].str.lower() != "nan")]

    numeric_cols = [c for c in REQUIRED_COLUMNS if c != "STAFF NAME"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["_GROSS"] = (
        df.get("BASIC", pd.Series(0, index=df.index))
        + df.get("HOUSING", pd.Series(0, index=df.index))
        + df.get("TRANSPORT", pd.Series(0, index=df.index))
    )

    if "TOTAL DED." not in df.columns or df["TOTAL DED."].sum() == 0:
        ded_cols = ["TAX", "PENSION", "LOAN", "SAL. ADV.", "PENALTY"]
        df["TOTAL DED."] = sum(
            df.get(c, pd.Series(0, index=df.index)) for c in ded_cols
        )

    if "NET SALARY" not in df.columns or df["NET SALARY"].sum() == 0:
        df["NET SALARY"] = df["_GROSS"] - df["TOTAL DED."]

    df.reset_index(drop=True, inplace=True)
    return df


def validate_columns(df: pd.DataFrame) -> list:
    """Return list of missing required columns (only STAFF NAME is truly required now)."""
    return [c for c in ["STAFF NAME"] if c not in df.columns]


def compute_summaries(df: pd.DataFrame) -> dict:
    """Compute aggregate payroll metrics."""
    if "GROSS SALARY" in df.columns:
        gross = df["GROSS SALARY"].sum()
    else:
        gross = (
            df.get("BASIC", pd.Series(0)).fillna(0)
            + df.get("HOUSING", pd.Series(0)).fillna(0)
            + df.get("TRANSPORT", pd.Series(0)).fillna(0)
        ).sum()

    return {
        "total_employees":  len(df),
        "total_gross":      float(gross),
        "total_deductions": float(df.get("TOTAL DED.", pd.Series(0)).sum()),
        "total_net":        float(df.get("NET SALARY", pd.Series(0)).sum()),
    }
