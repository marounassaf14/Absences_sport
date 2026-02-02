import pandas as pd
from datetime import datetime, date


def _is_absent_cell(value) -> bool:
    # détecte "Abs.", "abs", "ABS", et aussi "absent" (car contient "abs")
    return "abs" in str(value).strip().lower()


def _norm(s: str) -> str:
    return str(s).strip().lower().replace("\n", " ")


def _date_label_as_excel(col) -> str:
    """
    Rend le nom de colonne (date) sans l'heure.
    - Si c'est un Timestamp/datetime/date -> "dd/mm"
    - Si c'est une string type "2025-10-06 00:00:00" -> "dd/mm"
    - Sinon -> string telle quelle
    """
    # 1) vrai objet date
    if isinstance(col, (pd.Timestamp, datetime, date)):
        return col.strftime("%d/%m")

    # 2) string potentiellement parsable
    s = str(col).strip()
    dt = pd.to_datetime(s, errors="coerce")
    if not pd.isna(dt):
        return dt.strftime("%d/%m")

    # 3) fallback
    return s


def process_sheet(df: pd.DataFrame, id_cols=("Nom", "Prénom", "Email")) -> pd.DataFrame:
    df = df.copy()
    cols_to_check = [c for c in df.columns if c not in id_cols]

    absent_counts = []
    absent_dates_list = []

    for _, row in df.iterrows():
        absent_dates = [
            _date_label_as_excel(col)
            for col in cols_to_check
            if _is_absent_cell(row[col])
        ]
        absent_dates_list.append(absent_dates)
        absent_counts.append(len(absent_dates))

    df["Absent_Count"] = absent_counts
    df["Absent_Dates"] = absent_dates_list
    return df


def find_students_with_absences(excel_path: str, min_absences=2):
    sheets = pd.read_excel(excel_path, sheet_name=None)

    results = []
    for sheet_name, df in sheets.items():
        # --- normalize/rename columns ---
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

        # required columns
        if not {"Nom", "Prénom", "Email"}.issubset(df.columns):
            continue

        # enlever lignes sans identité (souvent ligne “test” ou vide)
        df = df.dropna(subset=["Nom", "Prénom", "Email"], how="any")

        dfp = process_sheet(df)
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
