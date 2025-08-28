from flask import Flask, render_template, jsonify, request
import pyodbc
from datetime import date
import os
from dotenv import load_dotenv

# Wczytaj zmienne środowiskowe z pliku .env
load_dotenv()

# Inicjalizacja aplikacji Flask
app = Flask(__name__)

DB_CONFIG = {
    'server': os.getenv('DB_SERVER'),
    'database': os.getenv('DB_DATABASE'),
    'username': os.getenv('DB_USERNAME'),
    'password': os.getenv('DB_PASSWORD'),
    'driver': os.getenv('DB_DRIVER', '{ODBC Driver 17 for SQL Server}'),
}

def get_db_connection():
    """Funkcja tworząca połączenie z bazą danych."""
    conn_str = (
        f"DRIVER={DB_CONFIG['driver']};"
        f"SERVER={DB_CONFIG['server']};"
        f"DATABASE={DB_CONFIG['database']};"
        "Trusted_Connection=yes;"
    )
    return pyodbc.connect(conn_str)

@app.route('/')
def index():
    """Główna strona, która wyświetla szablon HTML."""
    return render_template('index.html')


@app.route('/api/raport', methods=['POST'])
def raport_liczba_wnioskow():
    """API, które zwraca liczbę wniosków złożonych w danym roku oraz liczbę wniosków dla działu w wybranym okresie."""
    try:
        # Pobierz parametry z żądania
        data = request.get_json()
        rok = int(data.get('rok', 2024))
        miesiac_start = int(data.get('miesiac_start', 1))
        miesiac_koniec = int(data.get('miesiac_koniec', 12))
        dzial = data.get('dzial', 'CP')

        # Przygotuj daty do zapytań
        data_poczatkowa = date(rok, 1, 1)
        data_koncowa = date(rok, miesiac_koniec, 1).replace(day=28)  # tymczasowo 28, zaraz poprawka
        # Ustal ostatni dzień miesiąca
        import calendar
        last_day = calendar.monthrange(rok, miesiac_koniec)[1]
        data_koncowa = data_koncowa.replace(day=last_day)

        # Stwórz połączenie z bazą danych
        conn = get_db_connection()
        cursor = conn.cursor()

        # Zgłoszone wnioski (wszystkie wnioski w wybranym okresie i dziale)
        sql_zgloszone = """
            SELECT COUNT(*) 
            FROM dbo.WnioskiIdeaManagement
            WHERE YEAR(DataZgloszenia) = ?
              AND MONTH(DataZgloszenia) >= ?
              AND MONTH(DataZgloszenia) <= ?
              AND Dzial = ?
        """
        cursor.execute(sql_zgloszone, rok, miesiac_start, miesiac_koniec, dzial)
        zgloszone = cursor.fetchone()[0]

        # Zrealizowane wnioski (StatusWniosku = 'Zrealizowany')
        sql_zrealizowane = """
            SELECT COUNT(*) 
            FROM dbo.WnioskiIdeaManagement
            WHERE YEAR(DataZgloszenia) = ?
              AND MONTH(DataZgloszenia) >= ?
              AND MONTH(DataZgloszenia) <= ?
              AND Dzial = ?
              AND StatusWniosku = 'Zrealizowany'
        """
        cursor.execute(sql_zrealizowane, rok, miesiac_start, miesiac_koniec, dzial)
        zrealizowane = cursor.fetchone()[0]

        # Odrzucone wnioski
        sql_odrzucone = """
            SELECT COUNT(*) 
            FROM dbo.WnioskiIdeaManagement
            WHERE StatusWniosku = 'Odrzucony'
              AND Dzial = ?
              AND DataZgloszenia BETWEEN ? AND ?
        """
        cursor.execute(sql_odrzucone, dzial, data_poczatkowa, data_koncowa)
        odrzucone = cursor.fetchone()[0]

        # Zaakceptowane zrealizowane (StatusWniosku = 'Zatwierdzony' i ProgressBar = 1)
        sql_zaakceptowane_zrealizowane = """
            SELECT COUNT(*) 
            FROM dbo.WnioskiIdeaManagement
            WHERE DataZgloszenia BETWEEN ? AND ?
              AND Dzial = ?
              AND StatusWniosku = 'Zatwierdzony'
              AND (ProgressBar = 1)
        """
        cursor.execute(sql_zaakceptowane_zrealizowane, data_poczatkowa, data_koncowa, dzial)
        zaakceptowane_zrealizowane = cursor.fetchone()[0]

        cursor.close()
        conn.close()

        # Zwróć wyniki jako odpowiedź JSON
        return jsonify({
            'opZgloszone': zgloszone,
            'opZrealizowane': zrealizowane,
            'opOdrzucone': odrzucone,
            'opZaakceptowaneZrealizowane': zaakceptowane_zrealizowane,
            'kpiWnioskiNaPracownika': 0,
            'kpiCzasRealizacji': 0,
            'kpiBenefitNaWniosek': 0
        })

    except Exception as e:
        # Obsługa błędów
        return jsonify({'error': str(e)}), 500

# Uruchomienie serwera
if __name__ == '__main__':
    app.run(debug=True)