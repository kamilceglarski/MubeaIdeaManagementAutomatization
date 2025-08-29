Mubea Idea Management - Automatyzacja
Projekt ten automatyzuje proces synchronizacji danych z pliku Excel (połączonego z listą SharePoint) do bazy danych SQL Server oraz udostępnia prosty interfejs webowy do generowania raportów.

Automatyzacja opiera się na dwóch skryptach uruchamianych cyklicznie przez Harmonogram Zadań Windows:

Skrypt PowerShell: Odświeża dane w pliku Excel (Refresh All).

Skrypt Python: Importuje zaktualizowane dane z pliku Excel do docelowej bazy danych SQL, realizując logikę UPSERT lub "Wyczyść i Wczytaj".

Dodatkowo, Aplikacja Flask serwuje interfejs webowy do dynamicznego raportowania na podstawie danych z bazy.

⚙️ Wymagania Wstępne
Przed wdrożeniem upewnij się, że na serwerze docelowym zainstalowane jest następujące oprogramowanie:

Python (zalecana wersja 3.8+).

Microsoft Excel (pełna wersja desktopowa, wymagana do odświeżania danych przez skrypt PowerShell).

Microsoft ODBC Driver for SQL Server.

Git do sklonowania repozytorium.

🚀 Instalacja i Konfiguracja
1. Klonowanie Repozytorium
Sklonuj repozytorium do wybranej lokalizacji na serwerze (np. C:\MubeaAutomation):

Bash

git clone https://github.com/kamilceglarski/MubeaIdeaManagementAutomatization.git
cd MubeaIdeaManagementAutomatization
2. Konfiguracja Środowiska Python
Zalecane jest użycie wirtualnego środowiska, aby odizolować zależności projektu od innych aplikacji.

Bash

# Utwórz wirtualne środowisko
python -m venv venv

# Aktywuj środowisko
.\venv\Scripts\activate

# Zainstaluj wszystkie wymagane biblioteki
pip install -r requirements.txt
3. Konfiguracja Zmiennych Środowiskowych
Projekt korzysta z pliku .env do bezpiecznego przechowywania wrażliwych danych, takich jak dane logowania do bazy czy ścieżki plików.

Stwórz kopię pliku .env.example i zmień jej nazwę na .env.

Otwórz plik .env w edytorze tekstu i uzupełnij go poprawnymi wartościami dla Twojego środowiska.

⏰ Wdrożenie - Synchronizacja Automatyczna
Otwórz Harmonogram Zadań (Task Scheduler) na serwerze i skonfiguruj dwa oddzielne zadania, które będą uruchamiać skrypty w odpowiedniej kolejności.

✅ Zadanie 1: Odświeżanie Pliku Excel (PowerShell)
To zadanie uruchomi skrypt odswiez_raport.ps1, który wykona operację "Refresh All" w pliku Excel.

Zakładka "Ogólne" (General):

Nazwa: 1 - Mubea IM - Odswiezanie Raportu Excel

Zaznacz: "Uruchom niezależnie od tego, czy użytkownik jest zalogowany".

Zaznacz: "Uruchom z najwyższymi uprawnieniami".

Zakładka "Wyzwalacze" (Triggers):

Nowy...: Ustaw harmonogram, np. Codziennie o 07:00:00.

Zakładka "Akcje" (Actions):

Nowy...:

Akcja: Uruchom program

Program/skrypt: powershell.exe

Dodaj argumenty: -ExecutionPolicy Bypass -File "C:\MubeaAutomation\odswiez_raport.ps1" (dostosuj ścieżkę do lokalizacji projektu).

✅ Zadanie 2: Import Danych do Bazy (Python)
To zadanie uruchomi skrypt migration_excel_to_database.py, który przeniesie odświeżone dane do bazy SQL. Musi być uruchamiane po zakończeniu zadania nr 1.

Zakładka "Ogólne" (General):

Nazwa: 2 - Mubea IM - Import Danych do SQL

Ustaw te same opcje co w Zadaniu 1.

Zakładka "Wyzwalacze" (Triggers):

Nowy...: Ustaw harmonogram przesunięty o kilka minut względem pierwszego zadania, np. Codziennie o 07:05:00.

Zakładka "Akcje" (Actions):

Nowy...:

Akcja: Uruchom program

Program/skrypt: Podaj pełną ścieżkę do python.exe z Twojego wirtualnego środowiska, np. C:\MubeaAutomation\venv\Scripts\python.exe.

Dodaj argumenty: Podaj pełną ścieżkę do skryptu importującego, np. "C:\MubeaAutomation\migration_excel_to_database.py".

Rozpocznij w (opcjonalnie): Wpisz ścieżkę do głównego folderu projektu, np. C:\MubeaAutomation\.

📊 Uruchomienie Aplikacji Webowej
Aby uruchomić interfejs do generowania raportów, wykonaj poniższe kroki.

Otwórz terminal w głównym folderze projektu.

Aktywuj wirtualne środowisko:

Bash

.\venv\Scripts\activate
Uruchom serwer deweloperski Flask:

Bash

flask run
Aplikacja będzie domyślnie dostępna w przeglądarce pod adresem http://127.0.0.1:5000.
