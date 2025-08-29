import pandas as pd
import pyodbc

from config import EXCEL_FILE_PATH, EXCEL_SHEET_NAME



DB_SERVER = "UJAMSSQL02.group.mubea.net"
DB_DATABASE = "it"
DB_TABLE = "dbo.WnioskiIdeaManagement"
# Klucz biznesowy do łączenia danych z Excela i SQL
KEY_COLUMN_EXCEL = 'Wniosek Nr'
KEY_COLUMN_SQL = 'WniosekNumer'

DB_CONNECTION_STRING = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={DB_SERVER};"
    f"DATABASE={DB_DATABASE};"
    f"Trusted_Connection=yes;"
)

# --- POPRAWIONE MAPOWANIE KOLUMN ---
# Pamiętaj, aby nazwy w EXCEL_COLUMNS odpowiadały tym w pliku Excel

#Kolumny w excelu nie odpowiadaja kolumnom w SQL 
#Zmiana nazwy kolumny w sharepoincie nie zmienia jej nazwy pierwotnej. 
#W momencie gdy nazwa pierwotna miała nazwę ImieNazwisko i w przyszłości zostala zmieniona na dział jej nameID pozostało takie same 

EXCEL_COLUMNS = [
    'Tytuł', 'Data zgłoszenia', 'Dział', 'Opis usprawnienia', 'Priorytet', 'Wniosek Nr', 'Data wdrożenia', 'Poprawa wpłynie na', 'Status wniosku', 'ProgressBar'
]

SQL_COLUMNS = [
    'Tytul', 'DataZgloszenia', 'ImieNazwisko', 'Dzial', 'Priorytet', 'WniosekNumer', 'DataWdrozenia', 'PoprawaWplynieNa', 'StatusWniosku', 'ProgressBar'
]

# --- Funkcje pomocnicze (bez zmian) ---
def convert_polish_float(value):
    if value is None or pd.isna(value) or str(value).strip() == '': return None
    try:
        return float(str(value).replace(',', '.'))
    except (ValueError, TypeError):
        return None

# --- GŁÓWNA LOGIKA SKRYPTU ---
def sync_excel_to_sql():
    conn = None
    cursor = None
    try:
        print(f"Wczytywanie danych z pliku: {EXCEL_FILE_PATH}...")
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=EXCEL_SHEET_NAME)

        # --- Czyszczenie i transformacja danych (bez zmian) ---
        print("Czyszczenie i przygotowywanie danych...")
        if 'Tytuł' in df.columns: df['Tytuł'].fillna('Brak tytułu', inplace=True)
        if 'ProgressBar' in df.columns: df['ProgressBar'] = df['ProgressBar'].apply(convert_polish_float)
        
        df_clean = df.where(pd.notnull(df), None)
        print(f"Przygotowano {len(df_clean)} wierszy do importu.")
        
        # --- Połączenie z bazą ---
        print(f"Łączenie z bazą danych '{DB_DATABASE}'...")
        conn = pyodbc.connect(DB_CONNECTION_STRING)
        cursor = conn.cursor()
        print("✅ Połączenie z bazą danych udane.")

        # --- NOWA SEKCJA: Czyszczenie tabeli przed nowym importem ---
        print(f"Czyszczenie tabeli docelowej: {DB_TABLE}...")
        cursor.execute(f"TRUNCATE TABLE {DB_TABLE}")
        print("Tabela została wyczyszczona.")
        # -----------------------------------------------------------

        # --- UPROSZCZONA LOGIKA IMPORTU (tylko INSERT) ---
        print("Rozpoczynanie importu danych...")
        
        # Upewnienie się, że bierzemy pod uwagę tylko zmapowane kolumny
        df_to_insert = df_clean[EXCEL_COLUMNS]

        insert_cols_str = ", ".join([f"[{col}]" for col in SQL_COLUMNS])
        insert_placeholders = ", ".join(["?" for _ in SQL_COLUMNS])
        insert_query = f"INSERT INTO {DB_TABLE} ({insert_cols_str}) VALUES ({insert_placeholders})"
        
        data_tuples = [tuple(x) for x in df_to_insert.to_numpy()]
        
        cursor.fast_executemany = True
        cursor.executemany(insert_query, data_tuples)
        
        conn.commit()
        print(f"✅ Pomyślnie zaimportowano {len(df_to_insert)} wierszy.")

    except Exception as e:
        print(f"!!! Wystąpił krytyczny błąd: {e}")
        if conn: conn.rollback()
    finally:
        if cursor: cursor.close()
        if conn:
            conn.close()
            print("Połączenie z bazą danych zostało zamknięte.")

if __name__ == "__main__":
    sync_excel_to_sql()