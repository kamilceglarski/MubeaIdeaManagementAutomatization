# --- KONFIGURACJA ---
# Wpisz pelna sciezke do Twojego pliku Excel. Uzyj cudzyslowow.
$filePath = "C:\Users\ceglarskik\OneDrive - Mubea\Pulpit\Book1.xlsx"

# --- GLOWNA LOGIKA SKRYPTU ---
Write-Host "Rozpoczynanie procesu odswiezania pliku Excel..."

# Sprawdzenie, czy plik istnieje
if (-not (Test-Path $filePath)) {
    Write-Error "Plik nie zostal znaleziony pod sciezka: $filePath"
    exit
}

# Utworzenie niewidocznej instancji aplikacji Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false # Wylacza komunikaty typu "Czy na pewno zapisac?"

try {
    # Otwarcie skoroszytu
    Write-Host "Otwieranie pliku: $filePath"
    $workbook = $excel.Workbooks.Open($filePath)

    # Odswiezenie wszystkich polaczen danych w pliku
    Write-Host "Odswiezanie wszystkich polaczen danych..."
    $workbook.RefreshAll()

    # Zapisanie zmian
    Write-Host "Zapisywanie pliku..."
    $workbook.Save()

    Write-Host "Plik zostal pomyslnie odswiezony i zapisany."
}
catch {
    Write-Error "Wystapil blad: $($_.Exception.Message)"
}
finally {
    # Bardzo wazne: zamkniecie Excela i zwolnienie zasobow
    if ($workbook) { $workbook.Close() }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel, workbook
}