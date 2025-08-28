Mubea Idea Management - Automatyzacja
Poniższa instrukcja opisuje proces konfiguracji serwera w celu automatycznego odświeżania raportu w pliku Excel i importowania danych do bazy SQL Server. Proces składa się z dwóch zaplanowanych zadań.

Krok 1: Przygotowanie serwera ⚙️
Zanim skonfigurujesz harmonogram, upewnij się, że na serwerze spełnione są poniższe wymagania:

Lokalizacja skryptów: Oba skrypty, odswiez_raport.ps1 (PowerShell) oraz import_excel.py (Python), znajdują się w znanej lokalizacji (np. C:\Automatyzacja).

Dostęp do pliku Excel: Serwer musi mieć dostęp do pliku źródłowego Book1.xlsx.

Zainstalowany Microsoft Excel: Wymagany do uruchomienia skryptu PowerShell, który odświeża dane.

Zainstalowany Python i biblioteki: Na serwerze musi być zainstalowany Python oraz pakiety: pandas, openpyxl i pyodbc.

Zainstalowane sterowniki ODBC: Serwer musi mieć zainstalowany sterownik "ODBC Driver for SQL Server" do komunikacji z bazą danych.

Krok 2: Konfiguracja zadań w Harmonogramie Zadań ⏰
Otwórz Harmonogram Zadań (Task Scheduler) na serwerze i postępuj zgodnie z poniższymi instrukcjami, tworząc dwa oddzielne zadania.

✅ Zadanie 1: Automatyczne odświeżanie pliku Excel (PowerShell)
To zadanie uruchomi skrypt, który wykona "Refresh All" w Twoim pliku Excel.

Utwórz zadanie: W panelu Akcje kliknij Utwórz zadanie... (Create Task...).

Zakładka "Ogólne" (General):

Nazwa: 1 - Odswiezanie Raportu Excel

Zaznacz opcję "Uruchom niezależnie od tego, czy użytkownik jest zalogowany".

Zaznacz opcję "Uruchom z najwyższymi uprawnieniami".

Zakładka "Wyzwalacze" (Triggers):

Kliknij Nowy....

Ustaw harmonogram, np. Codziennie o godzinie 07:00:00.

Zakładka "Akcje" (Actions):

Kliknij Nowy....

Akcja: Uruchom program.

Program/skrypt: powershell.exe

Dodaj argumenty:

-ExecutionPolicy Bypass -File "C:\Automatyzacja\odswiez_raport.ps1"
✅ Zadanie 2: Import danych z Excela do bazy (Python)
To zadanie uruchomi skrypt Pythona, który przeniesie odświeżone dane do bazy SQL.

Utwórz zadanie: Ponownie kliknij Utwórz zadanie....

Zakładka "Ogólne" (General):

Nazwa: 2 - Import Danych Excel do SQL

Zaznacz opcje "Uruchom niezależnie od tego, czy użytkownik jest zalogowany" oraz "Uruchom z najwyższymi uprawnieniami".

Zakładka "Wyzwalacze" (Triggers):

Kliknij Nowy....

Ustaw harmonogram przesunięty o 5 minut w stosunku do pierwszego zadania, np. Codziennie o godzinie 07:05:00, aby dać czas na odświeżenie pliku.

Zakładka "Akcje" (Actions):

Kliknij Nowy....

Akcja: Uruchom program.

Program/skrypt: Podaj pełną ścieżkę do pliku python.exe, np.:

C:\Python39\python.exe
Dodaj argumenty: Podaj pełną ścieżkę do skryptu importującego w cudzysłowie, np.:

"C:\Automatyzacja\import_excel.py"
Rozpocznij w (opcjonalnie): Wpisz ścieżkę do folderu, w którym znajduje się skrypt:

C:\Automatyzacja\