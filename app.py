# TODO:
# 1. Przeanalizować plik Excel i zidentyfikować, w jaki sposób są obliczane wskaźniki/kpi (np. kpiWnioskiNaPracownika, kpiCzasRealizacji, kpiBenefitNaWniosek).
# 2. Na podstawie tej analizy przygotować odpowiednie zapytania SQL, które będą liczyć te wartości bezpośrednio w bazie danych. 
# 3. Dodać te zapytania do backendu (w tej funkcji) i zwracać ich wyniki przez API.
# 4. Na froncie (templates) podmieniać tylko textContent odpowiednich elementów na stronie na wartości zwrócone przez API (bez dodatkowych obliczeń po stronie JS).


from flask import Flask, render_template, jsonify, request
import pyodbc
from datetime import date
import os
from dotenv import load_dotenv
from queries import sql_queries 


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


        cursor.execute(sql_queries['get_zgloszone_wnioski'], rok,miesiac_koniec, dzial)
        zgloszone = cursor.fetchone()[0]

        cursor.execute(sql_queries['get_wnioski_otwarte'], rok, miesiac_koniec, dzial)
        otwarte = cursor.fetchone()[0]

        cursor.execute(sql_queries['get_odrzucone'], dzial, rok, miesiac_koniec)
        odrzucone = cursor.fetchone()[0]

        cursor.execute(sql_queries['get_wnioski_zaakceptowane_zrealizowane'], dzial, rok, miesiac_koniec)
        zaakceptowane_zrealizowane = cursor.fetchone()[0]


        cursor.execute(sql_queries['get_wnioski_otwarte_90dni'], rok, miesiac_koniec, dzial)
        otwarte_90dni = cursor.fetchone()[0]

        cursor.execute(sql_queries['get_sredni_czas_realizacji'], dzial, rok, miesiac_koniec)
        sredni_czas_realizacji = cursor.fetchone()[0]

        cursor.execute(sql_queries['get_zysk_wirtualny'], rok, miesiac_koniec, dzial)
        zysk_wirtualny = cursor.fetchone()[0]

        
        cursor.close()
        conn.close()

        # Zwróć wyniki jako odpowiedź JSON
        return jsonify({
            #KPI
            'kpiWnioskiNaPracownika': zgloszone ,
            'kpiCzasRealizacji': sredni_czas_realizacji,
            'kpiBenefitNaWniosek': 0, 

            #Raport Operacyjny 
            'opZgloszone': zgloszone,                                  #1. ZGŁOSZONE
            'opZgloszoneOtwarte': otwarte,                             #2. ZGŁOSZONE_OTWARTE
            'opOtwarte90dni': otwarte_90dni,                           #3. ZGŁOSZONE_OTWARTE > 90 dni
            'opZaakceptowaneZrealizowane': zaakceptowane_zrealizowane, #4. ZAAKCEPTOWANE_ZREALIZOWANE
            'opOdrzucone': odrzucone,                                  #5. ODRZUCONE
            'opSredniCzasRealizacji': sredni_czas_realizacji,          #6. Ø CZAS REALIZACJI                           
            'opZyskWirtualny': zysk_wirtualny,                         #7. ZYSK WIRTUALNY
            'opZyskMierzalny': 0                                       #8. ZYSK MIERZALNY
        })

    except Exception as e:
        # Obsługa błędów
        return jsonify({'error': str(e)}), 500

# Uruchomienie serwera
if __name__ == '__main__':
    app.run(debug=True)