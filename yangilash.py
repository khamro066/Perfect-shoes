import pandas as pd
import datetime
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

OUTPUT_FILE = "Ombor_2026_PowerBI.xlsx"
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))

SHEETS = {
    "Тери":    {"valyuta": "UZS"},
    "Хом ашё": {"valyuta": "UZS"},
    "Подош":   {"valyuta": "USD"},
}

def process_file(filepath):
    all_data = []

    for sheet_name, meta in SHEETS.items():
        try:
            raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        except Exception:
            print(f"  '{sheet_name}' sheet topilmadi, o'tkazildi.")
            continue

        row0 = raw.iloc[0]

        date_cols = []
        for col_idx, val in enumerate(row0):
            if isinstance(val, datetime.datetime):
                date_cols.append((col_idx, val.date()))

        if sheet_name == "Подош":
            num_col, name_col, narx_col = 0, 2, 4
        else:
            num_col, name_col, narx_col = 0, 2, 4

        data_rows = raw.iloc[2:].copy().reset_index(drop=True)

        for _, row in data_rows.iterrows():
            num_val = row.iloc[num_col]
            if pd.isna(num_val) or str(num_val).strip() in ["", "ЖАМИ", "nan"]:
                continue
            try:
                float(num_val)
            except (ValueError, TypeError):
                continue

            nomi = str(row.iloc[name_col]).strip() if pd.notna(row.iloc[name_col]) else ""
            if not nomi or nomi in ["nan", "ЖАМИ"]:
                continue

            try:
                narxi = float(row.iloc[narx_col]) if pd.notna(row.iloc[narx_col]) else 0
            except (ValueError, TypeError):
                narxi = 0

            for col_idx, sana in date_cols:
                def to_float(v):
                    try:
                        return float(v) if pd.notna(v) and str(v).strip() != "" else 0.0
                    except (ValueError, TypeError):
                        return 0.0

                kirim  = to_float(row.iloc[col_idx]     if col_idx     < len(row) else None)
                chikim = to_float(row.iloc[col_idx + 1] if col_idx + 1 < len(row) else None)
                qoldiq = to_float(row.iloc[col_idx + 2] if col_idx + 2 < len(row) else None)

                if kirim == 0 and chikim == 0:
                    continue

                all_data.append({
                    "Tur":      sheet_name,
                    "Mahsulot": nomi,
                    "Valyuta":  meta["valyuta"],
                    "Narxi":    narxi,
                    "Sana":     sana,
                    "Kirim":    kirim,
                    "Chikim":   chikim,
                    "Qoldiq":   qoldiq,
                })

    return pd.DataFrame(all_data)


def find_new_source_file():
    skip = {OUTPUT_FILE.lower(), "clean_ombor.py", "yangilash.py"}
    xlsx_files = [
        f for f in os.listdir(SCRIPT_DIR)
        if f.lower().endswith(".xlsx")
        and f.lower() not in skip
        and not f.startswith("~$")
    ]
    if not xlsx_files:
        return None
    # Eng yangi faylni tanlash
    xlsx_files.sort(key=lambda f: os.path.getmtime(os.path.join(SCRIPT_DIR, f)), reverse=True)
    return os.path.join(SCRIPT_DIR, xlsx_files[0])


if __name__ == "__main__":
    source = find_new_source_file()
    if not source:
        print("Xato: Perfect papkasida manba Excel fayl topilmadi.")
        sys.exit(1)

    print(f"Fayl topildi: {os.path.basename(source)}")
    print("Ishlanmoqda...")

    df = process_file(source)

    if df.empty:
        print("Xato: Ma'lumot topilmadi.")
        sys.exit(1)

    df["Sana"] = pd.to_datetime(df["Sana"])
    df = df.sort_values(["Tur", "Mahsulot", "Sana"]).reset_index(drop=True)

    output_path = os.path.join(SCRIPT_DIR, OUTPUT_FILE)
    df.to_excel(output_path, index=False)

    print(f"Tayyor! {len(df)} qator -> {OUTPUT_FILE}")
    print(f"Turlar: {df.groupby('Tur').size().to_dict()}")
