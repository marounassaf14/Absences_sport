import pandas as pd


def _is_absent_cell(value) -> bool:
    return "absent" in str(value).strip().lower()


def process_sheet(
    df: pd.DataFrame,
    id_cols=("Nom", "Prénom", "Email"),
) -> pd.DataFrame:

    df = df.copy()
    cols_to_check = [c for c in df.columns if c not in id_cols]

    absent_counts = []
    absent_dates_list = []

    for _, row in df.iterrows():
        absent_dates = [str(col) for col in cols_to_check if _is_absent_cell(row[col])]
        absent_dates_list.append(absent_dates)
        absent_counts.append(len(absent_dates))

    df["Absent_Count"] = absent_counts
    df["Absent_Dates"] = absent_dates_list
    return df


def find_students_with_absences(excel_path: str, min_absences=2):
    sheets = pd.read_excel(excel_path, sheet_name=None)

    results = []
    for sheet_name, df in sheets.items():
        if not {"Nom", "Prénom", "Email"}.issubset(df.columns):
            continue

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
