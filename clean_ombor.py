import pandas as pd
import datetime

INPUT_FILE = "Омбор 2026 (3).xlsx"
OUTPUT_FILE = "Ombor_2026_PowerBI.xlsx"
SHEETS = ["Тери", "Хом ашё", "Подош"]

all_data = []

for sheet_name in SHEETS:
    raw = pd.read_excel(INPUT_FILE, sheet_name=sheet_name, header=None)

    row0 = raw.iloc[0]
    row1 = raw.iloc[1]

    # Mahsulot ma'lumotlari uchun ustunlar
    # №, Kategoriya, Nomi, Narxi
    if sheet_name == "Подош":
        num_col, cat_col, name_col, narx_col = 0, 1, 2, 4
    else:
        num_col, cat_col, name_col, narx_col = 0, 1, 2, 4

    # Sana ustunlarini topish (row0 da datetime bo'lgan ustunlar)
    date_cols = []
    for col_idx, val in enumerate(row0):
        if isinstance(val, datetime.datetime):
            date_cols.append((col_idx, val.date()))

    # Ma'lumot qatorlari (row 2 dan boshlanadi, ЖАМИ qatorini o'tkazib yuborish)
    data_rows = raw.iloc[2:].copy()
    data_rows = data_rows[data_rows.iloc[:, num_col].notna()]
    data_rows = data_rows[data_rows.iloc[:, num_col] != 'ЖАМИ']

    for _, row in data_rows.iterrows():
        nom = str(row.iloc[num_col]).strip() if pd.notna(row.iloc[num_col]) else ""
        kategoriya = str(row.iloc[cat_col]).strip() if pd.notna(row.iloc[cat_col]) else sheet_name
        nomi = str(row.iloc[name_col]).strip() if pd.notna(row.iloc[name_col]) else ""
        narxi = row.iloc[narx_col] if pd.notna(row.iloc[narx_col]) else 0

        if not nomi or nomi in ["nan", "ЖАМИ"]:
            continue

        for col_idx, sana in date_cols:
            kirim = row.iloc[col_idx] if pd.notna(row.iloc[col_idx]) else 0
            chikim = row.iloc[col_idx + 1] if (col_idx + 1) < len(row) and pd.notna(row.iloc[col_idx + 1]) else 0
            qoldiq = row.iloc[col_idx + 2] if (col_idx + 2) < len(row) and pd.notna(row.iloc[col_idx + 2]) else 0

            try:
                kirim = float(kirim) if str(kirim).strip() != "" else 0
            except (ValueError, TypeError):
                kirim = 0
            try:
                chikim = float(chikim) if str(chikim).strip() != "" else 0
            except (ValueError, TypeError):
                chikim = 0
            try:
                qoldiq = float(qoldiq) if str(qoldiq).strip() != "" else 0
            except (ValueError, TypeError):
                qoldiq = 0

            # 0 dan tozalash - uchala ham 0 bo'lsa o'tkazib yuborish
            if kirim == 0 and chikim == 0 and qoldiq == 0:
                continue

            all_data.append({
                "Tur": sheet_name,
                "Kategoriya": kategoriya,
                "Mahsulot": nomi,
                "Narxi": narxi,
                "Sana": sana,
                "Kirim": kirim,
                "Chikim": chikim,
                "Qoldiq": qoldiq,
            })

df_result = pd.DataFrame(all_data)
df_result["Sana"] = pd.to_datetime(df_result["Sana"])
df_result = df_result.sort_values(["Tur", "Mahsulot", "Sana"]).reset_index(drop=True)

df_result.to_excel(OUTPUT_FILE, index=False)
print(f"Tayyor! Jami {len(df_result)} qator saqlandi: {OUTPUT_FILE}")
print(f"\nTurlar bo'yicha:")
print(df_result.groupby("Tur")[["Kirim","Chikim"]].sum())
