# Pomoc: `czytaj_modbus.py`

Skrypt cyklicznie odczytuje rejestry Modbus RTU przez bibliotekę `pymodbus` i
wyświetla je razem z opisami zmiennych z pliku Excel. Domyślnie odczyt jest
pokazywany w interaktywnej tabeli terminalowej.

## Wymagania

- Python 3.9 lub nowszy
- urządzenie Modbus RTU, port szeregowy i prawidłowy adres slave'a

## Przygotowanie środowiska

W PowerShellu, z katalogu projektu, utwórz i aktywuj wirtualne środowisko,
a następnie zainstaluj zależności:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r .\requirements.txt
```

W kolejnych sesjach PowerShell aktywuj środowisko poleceniem:

```powershell
.\.venv\Scripts\Activate.ps1
```

Skrypt używa komunikacji RTU z parzystością `even`, 8 bitami danych i 1 bitem
stopu (8E1). Wymaganie `pymodbus` jest przypięte do wersji 3.8.6, która
obsługuje Python 3.9 i nowszy.

## Uruchomienie

W PowerShellu, z katalogu zawierającego skrypt:

```powershell
.\.venv\Scripts\python.exe .\czytaj_modbus.py
```

Przykład z ustawieniami dla portu COM3:

```powershell
.\.venv\Scripts\python.exe .\czytaj_modbus.py --port COM3 --slave 1 --baud 38400
```

Pełną listę opcji można wyświetlić poleceniem:

```powershell
.\.venv\Scripts\python.exe .\czytaj_modbus.py --help
```

## Opcje

| Opcja | Domyślnie | Opis |
|---|---:|---|
| `--port` | `/dev/ttyUSB0` | Port szeregowy, np. `COM3` w Windows lub `/dev/ttyUSB0` w Linuxie. |
| `--slave` | `1` | Adres urządzenia Modbus. |
| `--baud` | `38400` | Prędkość transmisji. |
| `--start` | `8960` | Pierwszy adres rejestru zakresu. |
| `--count` | `103` | Liczba adresów w zakresie, począwszy od `--start`. |
| `--interval` | `1.0` | Odstęp między kolejnymi odczytami, w sekundach. |
| `--timeout` | `3.0` | Limit czasu odpowiedzi Modbus, w sekundach. |
| `--metadata` | `registers/status_var.xlsx` obok skryptu | Ścieżka do pliku XLSX z arkuszem `Status_var`. |
| `--no-clear` | wyłączone | Zamiast interaktywnej tabeli wypisuje kolejne raporty w terminalu; przydatne do przekierowania wyjścia do pliku. |
| `--no-color` | wyłączone | Wyłącza kolorowanie wartości. |

Przykład zmiany zakresu i częstotliwości odczytu:

```powershell
.\.venv\Scripts\python.exe .\czytaj_modbus.py --port COM3 --start 8960 --count 103 --interval 2
```

Przykład wskazania własnego pliku z opisami:

```powershell
.\.venv\Scripts\python.exe .\czytaj_modbus.py --port COM3 --metadata ".\registers\status_var.xlsx"
```

## Plik z opisami zmiennych

Domyślny plik to `registers/status_var.xlsx`, liczony względem katalogu, w którym
znajduje się skrypt. Musi zawierać arkusz `Status_var` i kolumny:

- `Address` — adres rejestru; musi dać się odczytać jako liczba całkowita,
- `Name` — nazwa zmiennej,
- `IEC type` — typ IEC zmiennej.

Dodatkowe kolumny `Description`, `Min`, `Max`, `Scale`, `Offset` i `Unit` są
wyświetlane, jeśli występują w arkuszu. Brakujące wartości są prezentowane jako
`-`. Skrypt pokazuje wartość odczytaną z rejestru; nie przelicza jej według
`Scale` ani `Offset`. Dla typu `BOOL` wartości `0` i `1` są dodatkowo
prezentowane jako `FALSE` i `TRUE`.

Skrypt odczytuje tylko adresy obecne w arkuszu i należące do wybranego zakresu.
Jeśli zmienna ma typ `UDINT`, odczytuje również następny rejestr, o ile mieści
się on w wybranym zakresie.

## Obsługa i wyniki

- W tabeli interaktywnej naciśnij `q`, aby zakończyć.
- `Home` i `End` przewijają tabelę na początek i koniec.
- W trybie `--no-clear` można przerwać działanie przez `Ctrl+C`.
- `BLAD:` oznacza błąd odczytu, w tym przekroczenie limitu czasu.
- Jeśli urządzenie zgłosi nieprawidłowy adres dla większego bloku, skrypt
  dzieli blok na mniejsze odczyty, aby wskazać rejestry powodujące problem.
- Jeżeli zakres nie zawiera żadnego adresu z arkusza, skrypt zakończy działanie
  z komunikatem o błędzie.

Przykładowe dopisywanie raportów do pliku:

```powershell
.\.venv\Scripts\python.exe .\czytaj_modbus.py --port COM3 --no-clear *> .\odczyt_modbus.log
```

## Typowe problemy

- **Błąd połączenia lub timeout** — sprawdź port, adres slave'a, prędkość,
  parametry transmisji urządzenia oraz połączenia RS-485.
- **Nie znaleziono pliku z mapą zmiennych** — sprawdź ścieżkę podaną przez
  `--metadata`; domyślnie plik powinien znajdować się pod
  `registers/status_var.xlsx`.
- **Brak arkusza lub wymaganych kolumn** — sprawdź nazwę arkusza
  (`Status_var`) i nagłówki `Address`, `Name`, `IEC type`.
- **Brak biblioteki Python** — zainstaluj zależności poleceniem z sekcji
  [Wymagania](#wymagania).
