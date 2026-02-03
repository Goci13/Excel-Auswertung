import pandas as pd


def analyze_excel(path: str):
    try:
        df = pd.read_excel(path)

        required_columns = {"Datum", "Kategorie", "Betrag"}
        if not required_columns.issubset(df.columns):
            return None, "Fehlende Spalten: Datum, Kategorie, Betrag"

        df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")

        total = df["Betrag"].sum()
        by_category = df.groupby("Kategorie")["Betrag"].sum()
        by_month = df.groupby(df["Datum"].dt.to_period("M"))["Betrag"].sum()

        return {
            "total": total,
            "by_category": by_category,
            "by_month": by_month
        }, None

    except Exception as e:
        return None, str(e)
    
def export_to_excel(result: dict, save_path: str):
    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
        # Übersicht
        overview_df = pd.DataFrame({
            "Kennzahl": ["Gesamtsumme"],
            "Wert": [result["total"]]
        })
        overview_df.to_excel(writer, sheet_name="Übersicht", index=False)

        # Kategorien
        result["by_category"].reset_index().rename(
            columns={"Kategorie": "Kategorie", "Betrag": "Summe"}
        ).to_excel(writer, sheet_name="Kategorien", index=False)

        # Monate
        result["by_month"].reset_index().rename(
            columns={"Datum": "Monat", "Betrag": "Summe"}
        ).to_excel(writer, sheet_name="Monate", index=False)

