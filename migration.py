import pandas as pd
import pyodbc

# --- KONFIGURACJA ---
# Uzupełnij poniższe zmienne swoimi danymi

# 1. Ustawienia pliku Excel
EXCEL_FILE_PATH = r'C:\Users\ceglarskik\OneDrive - Mubea\Pulpit\Book1.xlsx' # Ważne 'r' przed ścieżką!
EXCEL_SHEET_NAME = 0 # 0 oznacza pierwszy arkusz w pliku

# 2. Ustawienia Bazy Danych SQL Server
DB_SERVER = "UJAMSSQL02.group.mubea.net"
DB_DATABASE = "it"
DB_TABLE = "dbo.WnioskiIdeaManagement"
# Uwierzytelnianie Windows (jeśli używasz loginu i hasła, zmień connection string)
DB_CONNECTION_STRING = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={DB_SERVER};"
    f"DATABASE={DB_DATABASE};"
    f"Trusted_Connection=yes;"
)


# 3. Mapowanie kolumn
EXCEL_COLUMNS = [
    'Tytuł', 'Data zgłoszenia', 'Dział', 'Opis usprawnienia', 'Priorytet', 'Wniosek Nr', 'Data wdrożenia', 'Poprawa wpłynie na', 'Status wniosku', 'ProgressBar'
]

SQL_COLUMNS = [
    'Tytul', 'DataZgloszenia', 'ImieNazwisko', 'Dzial', 'Priorytet', 'WniosekNumer', 'DataWdrozenia', 'PoprawaWplynieNa', 'StatusWniosku', 'ProgressBar'
]


# --- GŁÓWNA LOGIKA SKRYPTU --
def import_excel_to_sql():
    """Wczytuje dane z Excela i importuje je do bazy danych SQL Server."""
    
    try:
        # --- Krok 1: Wczytanie danych z Excela ---
        print(f"Wczytywanie danych z pliku: {EXCEL_FILE_PATH}...")
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=EXCEL_SHEET_NAME)
        
        # Upewnij się, że wczytano tylko potrzebne kolumny z zachowaniem kolejności

        df['Tytuł'].fillna('Brak tytułu', inplace=True)
        
        df = df[EXCEL_COLUMNS]
        df = df.where(pd.notnull(df), None)


        df['ProgressBar'] = df['ProgressBar'].apply(lambda x: 0 if x is None or (isinstance(x, float) and pd.isna(x)) or x == '' else x)

       # Zamień NaN, '', None, pd.NaT na None (NULL w SQL) dla daty wdrożenia
        df['Data wdrożenia'] = df['Data wdrożenia'].apply(
            lambda x: None if x in [None, '', pd.NaT] or (isinstance(x, float) and pd.isna(x)) else x
        )
        
        # Konwersja pustych wartości (NaN) na None, które SQL rozumie jako NULL

        print(f"Wczytano {len(df)} wierszy z Excela do importu.")
        if df.empty:
            print("Plik Excel jest pusty lub nie zawiera wymaganych kolumn. Zakończono.")
            return

        # --- Krok 2: Połączenie z bazą danych ---
        print(f"Łączenie z bazą danych '{DB_DATABASE}' na serwerze '{DB_SERVER}'...")
        conn = pyodbc.connect(DB_CONNECTION_STRING)

       
        cursor = conn.cursor()

        print("✅ Połączenie z bazą danych udane.")

        # --- Krok 3: Przygotowanie i wykonanie zapytań INSERT ---
        print("Rozpoczynanie importu danych do tabeli...")
        
        # Zapytanie INSERT zostanie dynamicznie zbudowane dla wszystkich zdefiniowanych kolumn
        columns_str = ", ".join([f"[{col}]" for col in SQL_COLUMNS])
        placeholders_str = ", ".join(["?" for _ in SQL_COLUMNS])
        insert_query = f"INSERT INTO {DB_TABLE} ({columns_str}) VALUES ({placeholders_str})"
        
        for index, row in df.iterrows():
            cursor.execute(insert_query, tuple(row))

        conn.commit()
        print(f"✅ Pomyślnie zaimportowano {len(df)} wierszy do tabeli '{DB_TABLE}'.")

    except FileNotFoundError:
        print(f"!!! BŁĄD: Nie znaleziono pliku Excel pod ścieżką: {EXCEL_FILE_PATH}")
    except KeyError as e:
        print(f"!!! BŁĄD: W Twoim pliku Excel brakuje kolumny o nazwie: {e}. Sprawdź literówki w pliku lub w skrypcie.")
    except Exception as e:
        print(f"!!! Wystąpił krytyczny błąd: {e}")
        if 'conn' in locals() and conn:
            conn.rollback()
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if 'conn' in locals() and conn:
            conn.close()
            print("Połączenie z bazą danych zostało zamknięte.")


if __name__ == "__main__":
    import_excel_to_sql()