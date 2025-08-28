# MubeaIdeaManagementAutomatization


## Krok 1: Przygotowanie serwera
Zanim skonfigurujesz harmonogram, upewnij się, że na serwerze:

Znajdują się oba skrypty: odswiez_raport.ps1 (PowerShell) oraz import_excel.py (Python) w znanej lokalizacji (np. C:\Automatyzacja).

Znajduje się plik Excel: Serwer musi mieć dostęp do pliku Book1.xlsx.

Jest zainstalowany Microsoft Excel: Wymagany do uruchomienia skryptu PowerShell.

Jest zainstalowany Python oraz wymagane biblioteki: Upewnij się, że na serwerze jest Python i zainstalowane są pakiety pandas, openpyxl i pyodbc.

Są zainstalowane sterowniki ODBC: Serwer musi mieć zainstalowany sterownik "ODBC Driver for SQL Server".

## Krok 2: Konfiguracja zadań w Harmonogramie Zadań
Otwórz Harmonogram Zadań (Task Scheduler) na serwerze i postępuj zgodnie z poniższymi instrukcjami.

Zadanie 1: Automatyczne odświeżanie pliku Excel (PowerShell)
To zadanie uruchomi skrypt, który wykona "Refresh All" w Twoim pliku Excel.

Utwórz zadanie: W panelu Akcje kliknij Utwórz zadanie... (Create Task...).

Zakładka "Ogólne" (General):

Nazwa: Wpisz 1 - Odswiezanie Raportu Excel.

Zaznacz opcję "Uruchom niezależnie od tego, czy użytkownik jest zalogowany" (Run whether user is logged on or not).

Zaznacz opcję "Uruchom z najwyższymi uprawnieniami" (Run with highest privileges).

Zakładka "Wyzwalacze" (Triggers):

Kliknij Nowy... (New...).

Ustaw harmonogram, np. Codziennie (Daily) o godzinie 07:00:00.

Zakładka "Akcje" (Actions):

Kliknij Nowy... (New...).

Akcja: Uruchom program (Start a program).

Program/skrypt: powershell.exe

Dodaj argumenty: -ExecutionPolicy Bypass -File "C:\Automatyzacja\odswiez_raport.ps1" (pamiętaj, aby podać poprawną ścieżkę do pliku).

Zadanie 2: Import danych z Excela do bazy (Python)
To zadanie uruchomi skrypt Pythona, który przeniesie odświeżone dane do bazy SQL. Ustawimy je tak, aby uruchamiało się kilka minut po pierwszym zadaniu.

Utwórz zadanie: Ponownie kliknij Utwórz zadanie....

Zakładka "Ogólne" (General):

Nazwa: Wpisz 2 - Import Danych Excel do SQL.

Zaznacz opcje "Uruchom niezależnie od tego, czy użytkownik jest zalogowany" oraz "Uruchom z najwyższymi uprawnieniami".

Zakładka "Wyzwalacze" (Triggers):

Kliknij Nowy....

Ustaw ten sam harmonogram, co dla pierwszego zadania, ale przesunięty o 5 minut, np. Codziennie (Daily) o godzinie 07:05:00. To da pewność, że plik Excel zdążył się odświeżyć i zapisać.

Zakładka "Akcje" (Actions):

Kliknij Nowy....

Akcja: Uruchom program.

Program/skrypt: Podaj pełną ścieżkę do pliku wykonywalnego python.exe w Twojej instalacji Pythona na serwerze, np. C:\Python39\python.exe.

Dodaj argumenty: Podaj pełną ścieżkę do swojego skryptu importującego, np. "C:\Automatyzacja\import_excel.py".

Rozpocznij w (opcjonalnie): Wpisz ścieżkę do folderu, w którym znajduje się skrypt, np. C:\Automatyzacja\.