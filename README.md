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

Domyślne parametry programu, w tym komunikacji RTU (8E1), znajdują się w pliku
`config.json` obok skryptu. Wymaganie `pymodbus` jest przypięte do wersji 3.8.6,
która obsługuje Python 3.9 i nowszy.

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
| `--config` | `config.json` obok skryptu | Plik JSON z domyślnymi ustawieniami. |
| `--port` | `/dev/ttyUSB0` | Port szeregowy, np. `COM3` w Windows lub `/dev/ttyUSB0` w Linuxie. |
| `--slave` | `1` | Adres urządzenia Modbus. |
| `--baud` | `38400` | Prędkość transmisji. |
| `--parity` | `E` | Parzystość: `N`, `E` lub `O`. |
| `--bytesize` | `8` | Liczba bitów danych: `5`, `6`, `7` lub `8`. |
| `--stopbits` | `1` | Liczba bitów stopu: `1`, `1.5` lub `2`. |
| `--start` | `8960` | Pierwszy adres rejestru zakresu. |
| `--address-base` | `1` | Numeracja adresów z arkusza i zakresu: `1` (jak `modpoll`) lub `0` (adres PDU). |
| `--count` | `103` | Liczba adresów w zakresie, począwszy od `--start`. |
| `--interval` | `1.0` | Odstęp między kolejnymi odczytami, w sekundach. |
| `--timeout` | `3.0` | Limit czasu odpowiedzi Modbus, w sekundach. |
| `--metadata` | `registers/status_var.xlsx` obok skryptu | Ścieżka do pliku XLSX z arkuszem `Status_var`. |
| `--no-clear` | wyłączone | Zamiast interaktywnej tabeli wypisuje kolejne raporty w terminalu; przydatne do przekierowania wyjścia do pliku. |
| `--clear` | — | Wymusza interaktywną tabelę, także gdy `no_clear` jest włączone w konfiguracji. |
| `--no-color` | wyłączone | Wyłącza kolorowanie wartości. |
| `--color` | — | Włącza kolorowanie, także gdy `no_color` jest włączone w konfiguracji. |

Plik `config.json` obsługuje komentarze rozpoczynające się od `//`, dzięki
czemu przy każdym ustawieniu znajduje się jego opis. Wartości w pliku są
domyślne; podanie opcji w wierszu poleceń
zastępuje odpowiadającą jej wartość z konfiguracji. Można wskazać inny plik:

```powershell
.\.venv\Scripts\python.exe .\czytaj_modbus.py --config ".\moja-konfiguracja.json"
```

Edytuj wartości w `config.json`, aby zmienić ustawienia domyślne. Ścieżka
względna w polu `metadata` jest liczona względem katalogu pliku konfiguracji.
Plik musi zawierać wszystkie klucze widoczne w domyślnym `config.json`.

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
znajduje się skrypt. Preferowany arkusz to `Status_var`. Jeśli plik ma tylko
jeden arkusz, skrypt użyje go również pod inną nazwą. Arkusz musi zawierać
kolumny:

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
`address_base` określa numerację adresów w arkuszu i opcjach `--start`/`--count`.
Przy wartości `1` (domyślnej, zgodnej z `modpoll`) skrypt odejmuje jeden przed
wywołaniem `pymodbus`, który przyjmuje adres PDU od zera. Ustaw `0`, jeśli mapa
zawiera adresy PDU i mają być przekazywane bez zmian.

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
- **Brak arkusza lub wymaganych kolumn** — sprawdź, czy plik zawiera arkusz
  `Status_var` (lub tylko jeden arkusz) oraz nagłówki `Address`, `Name`,
  `IEC type`.
- **Brak biblioteki Python** — zainstaluj zależności poleceniem z sekcji
  [Wymagania](#wymagania).
