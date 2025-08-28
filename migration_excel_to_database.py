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

# --- Funkcje pomocnicze do czyszczenia danych ---
def convert_polish_float(value):
    if value is None or pd.isna(value) or str(value).strip() == '': return None
    try:
        return float(str(value).replace(',', '.'))
    except (ValueError, TypeError):
        return None

# --- GŁÓWNA LOGIKA SKRYPTU ---
def upsert_excel_to_sql():
    """Wczytuje dane z Excela i wykonuje operację UPSERT do bazy SQL Server."""
    conn = None
    cursor = None
    try:
        print(f"Wczytywanie danych z pliku: {EXCEL_FILE_PATH}...")
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=EXCEL_SHEET_NAME)

        # --- Czyszczenie i transformacja danych ---
        print("Czyszczenie i przygotowywanie danych...")
        if 'Tytuł' in df.columns: df['Tytuł'].fillna('Brak tytułu', inplace=True)
        if 'ProgressBar' in df.columns: df['ProgressBar'] = df['ProgressBar'].apply(convert_polish_float)
        
        # Globalna zamiana pustych wartości na None, które SQL rozumie jako NULL
        df_clean = df.where(pd.notnull(df), None)

        print(f"Przygotowano {len(df_clean)} wierszy do synchronizacji.")
        
        # --- Połączenie z bazą ---
        print(f"Łączenie z bazą danych '{DB_DATABASE}'...")
        conn = pyodbc.connect(DB_CONNECTION_STRING)
        cursor = conn.cursor()
        print("✅ Połączenie z bazą danych udane.")

        # --- Budowanie zapytań ---
        # Zapytanie INSERT
        insert_cols_str = ", ".join([f"[{col}]" for col in SQL_COLUMNS])
        insert_placeholders = ", ".join(["?" for _ in SQL_COLUMNS])
        insert_query = f"INSERT INTO {DB_TABLE} ({insert_cols_str}) VALUES ({insert_placeholders})"

        # Zapytanie UPDATE (bez klucza w części SET)
        update_cols_list = [f"[{col}] = ?" for col in SQL_COLUMNS if col != KEY_COLUMN_SQL]
        update_set_str = ", ".join(update_cols_list)
        update_query = f"UPDATE {DB_TABLE} SET {update_set_str} WHERE [{KEY_COLUMN_SQL}] = ?"

        # --- Pętla synchronizująca ---
        print("Rozpoczynanie synchronizacji (UPSERT)...")
        inserted_count = 0
        updated_count = 0

        for index, row in df_clean.iterrows():
            key_value = row[KEY_COLUMN_EXCEL]
            
            # Jeśli klucz jest pusty, pomiń wiersz, aby uniknąć błędów
            if key_value is None:
                continue

            # 1. Sprawdź, czy rekord o danym numerze wniosku już istnieje
            cursor.execute(f"SELECT 1 FROM {DB_TABLE} WHERE [{KEY_COLUMN_SQL}] = ?", key_value)
            exists = cursor.fetchone()

            if exists:
                # 2a. Jeśli tak -> zaktualizuj (UPDATE)
                update_values = [row.get(ex_col) for ex_col, sql_col in zip(EXCEL_COLUMNS, SQL_COLUMNS) if sql_col != KEY_COLUMN_SQL]
                update_values.append(key_value) # Dodaj klucz na końcu dla klauzuli WHERE
                cursor.execute(update_query, tuple(update_values))
                updated_count += 1
            else:
                # 2b. Jeśli nie -> wstaw nowy (INSERT)
                insert_values = [row.get(col) for col in EXCEL_COLUMNS]
                cursor.execute(insert_query, tuple(insert_values))
                inserted_count += 1
        
        conn.commit()
        print("✅ Synchronizacja zakończona pomyślnie.")
        print(f"Dodano nowych wierszy: {inserted_count}")
        print(f"Zaktualizowano istniejących wierszy: {updated_count}")

    except Exception as e:
        print(f"!!! Wystąpił krytyczny błąd: {e}")
        if conn: conn.rollback()
    finally:
        if cursor: cursor.close()
        if conn:
            conn.close()
            print("Połączenie z bazą danych zostało zamknięte.")

if __name__ == "__main__":
    upsert_excel_to_sql()