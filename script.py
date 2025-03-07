import pandas as pd

def merge_excels(input_files, output_file):
    dataframes = []
    
    for file in input_files:
        try:
            # Citim datele din Sheet1
            df = pd.read_excel(file, sheet_name="Sheet1", dtype=str)
            
            # Verificăm dacă fișierul conține date
            if df.empty:
                print(f"Avertizare: Fișierul {file} este gol și va fi ignorat.")
                continue

            # Verificăm dacă există coloane în fișier
            if df.columns.empty:
                print(f"Avertizare: Fișierul {file} nu are coloane valide și va fi ignorat.")
                continue

            # Afișăm coloanele disponibile pentru debugging
            print(f"Fișier: {file} -> Coloane disponibile: {df.columns.tolist()}")

            # Curățăm numele coloanelor (eliminăm spațiile inutile)
            df.columns = df.columns.astype(str).str.strip()

            # Verificăm dacă toate coloanele necesare există
            selected_columns = ["NUME", "PRENUME", "FACULTATE", "JUDET", "TARA"]
            missing_columns = [col for col in selected_columns if col not in df.columns]
            
            if missing_columns:
                print(f"Avertizare: Fișierul {file} nu conține coloanele: {missing_columns} și va fi ignorat.")
                continue  # Sărim peste fișierele incomplete

            # Selectăm doar coloanele necesare
            df = df[selected_columns]
            
            # Dacă JUDET este gol și TARA nu este "Romania", păstrăm valoarea din TARA
            df.loc[df["JUDET"].isna() | (df["JUDET"].str.strip() == ""), "TARA"] = df["TARA"]
            
            dataframes.append(df)

        except Exception as e:
            print(f"Eroare la procesarea fișierului {file}: {e}")
            continue  # Evităm oprirea programului în caz de eroare la un fișier
    
    if dataframes:
        final_df = pd.concat(dataframes, ignore_index=True)
        
        # Adăugăm coloana Nr. crt. ca prima coloană
        final_df.insert(0, "Nr. crt.", range(1, len(final_df) + 1))
        
        final_df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"Fișierul final a fost salvat ca {output_file}")
    else:
        print("Eroare: Niciun fișier valid nu a fost găsit.")

# Lista cu numele fișierelor de intrare
input_files = [
    "export_date_studenti1740574949840.xlsx",
    "export_date_studenti1740575392150.xlsx",
    "export_date_studenti1740575440654.xlsx",
    "export_date_studenti1740575484669.xlsx",
    "export_date_studenti1740575526963.xlsx"
]

# Numele fișierului de ieșire
output_file = "rezultat_final.xlsx"

# Apelăm funcția
merge_excels(input_files, output_file)
