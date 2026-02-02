import pandas as pd
from datetime import datetime, date


def _is_absent_cell(value) -> bool:
    return "abs" in str(value).strip().lower()


def _norm(s: str) -> str:
    return str(s).strip().lower().replace("\n", " ")


def _parse_col_date(col):
    """Try parsing a column header into a datetime. Returns datetime-like or NaT."""
    if isinstance(col, (pd.Timestamp, datetime, date)):
        return pd.to_datetime(col)
    s = str(col).strip()
    return pd.to_datetime(s, errors="coerce")


def _keep_col_for_semester(col, semester: str) -> bool:
    """
    S1: Sept (9) -> Jan (1)
    S2: Feb (2) -> May (5)
    """
    dt = _parse_col_date(col)
    if dt is None or pd.isna(dt):
        return False

    m = dt.month
    if semester == "S1":
        return (m >= 9) or (m == 1)
    if semester == "S2":
        return 2 <= m <= 5
    return True  # no filter


def _date_label(col) -> str:
    """Display column date as dd/mm (no hour). Fallback to raw string."""
    dt = _parse_col_date(col)
    if dt is not None and not pd.isna(dt):
        return dt.strftime("%d/%m")
    return str(col)


def process_sheet(df: pd.DataFrame, id_cols=("Nom", "Prénom", "Email"), semester=None) -> pd.DataFrame:
    df = df.copy()
    cols_to_check = [c for c in df.columns if c not in id_cols]

    if semester in ("S1", "S2"):
        cols_to_check = [c for c in cols_to_check if _keep_col_for_semester(c, semester)]

    absent_counts = []
    absent_dates_list = []

    for _, row in df.iterrows():
        absent_dates = [
            _date_label(col)
            for col in cols_to_check
            if _is_absent_cell(row[col])
        ]
        absent_dates_list.append(absent_dates)
        absent_counts.append(len(absent_dates))

    df["Absent_Count"] = absent_counts
    df["Absent_Dates"] = absent_dates_list
    return df


def find_students_with_absences(excel_path: str, min_absences=2, semester=None):
    sheets = pd.read_excel(excel_path, sheet_name=None)
    results = []

    for sheet_name, df in sheets.items():
        # Normalize column names (Email/email, Prénom/prenom)
        rename_map = {}
        for c in df.columns:
            if isinstance(c, str):
                nc = _norm(c)
                if nc == "email":
                    rename_map[c] = "Email"
                elif nc == "nom":
                    rename_map[c] = "Nom"
                elif nc in ("prénom", "prenom"):
                    rename_map[c] = "Prénom"

        if rename_map:
            df = df.rename(columns=rename_map)

        if not {"Nom", "Prénom", "Email"}.issubset(df.columns):
            continue

        df = df.dropna(subset=["Nom", "Prénom", "Email"], how="any")

        dfp = process_sheet(df, semester=semester)
        filtered = dfp[dfp["Absent_Count"] >= int(min_absences)]

        for _, row in filtered.iterrows():
            results.append(
                {
                    "Sport": sheet_name,
                    "Nom": str(row["Nom"]).strip(),
                    "Prénom": str(row["Prénom"]).strip(),
                    "Email": str(row["Email"]).strip(),
                    "Absent_Count": int(row["Absent_Count"]),
                    "Absent_Dates": list(row["Absent_Dates"]) if isinstance(row["Absent_Dates"], list) else [],
                }
            )

    return results
