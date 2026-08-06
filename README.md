# office-mcp

Serwer MCP, który pozwala Claude'owi sterować **otwartymi** aplikacjami Microsoft Office 2019
(PowerPoint, Excel, Word) na Windows przez automatyzację COM. Model tworzy i edytuje dokumenty
promptem, a zmiany widać na żywo w oknie aplikacji — bez pośredniego generowania plików i
otwierania ich ręcznie.

- 72 narzędzia MCP: `ppt_*` (24), `xl_*` (25), `doc_*` (23)
- pełny cykl: tworzenie, odczyt istniejących dokumentów, edycja, formatowanie, wykresy, obrazy, tabele
- podłącza się do już uruchomionej instancji Office zamiast otwierać drugą
- 216 testów jednostkowych i integracyjnych działających bez zainstalowanego Office

## Architektura

```
Claude Code / Claude Desktop
        │  (stdio, protokół MCP)
        ▼
┌─────────────────────┐
│   MCP Server        │  Python, oficjalny SDK `mcp`
│   (server.py)       │  Rejestruje narzędzia ppt_*, xl_*, doc_*
└──────────┬──────────┘
           │  TCP (localhost, JSON-line)
           ▼
┌─────────────────────┐
│   Office Bridge     │  Python, pywin32 (win32com.client)
│   (bridge/*.py)     │  Trzyma żywe połączenia COM per aplikacja
└──────────┬──────────┘
           │  COM
           ▼
  PowerPoint.Application / Excel.Application / Word.Application
```

**Dlaczego dwie warstwy:**

- klient MCP może restartować `server.py` do woli — Bridge żyje dalej i nie gubi połączeń COM
  ani otwartych dokumentów użytkownika,
- Bridge jest zwykłym serwerem TCP, więc można się do niego podpiąć czymkolwiek innym
  (skrypt debugowy, GUI) bez ruszania warstwy MCP,
- rozdzielenie ułatwia testy: kontrolery testuje się z zamockowanym COM, transport osobno.

**Izolacja aplikacji.** Każda aplikacja Office dostaje własny wątek COM (apartament STA),
własny obiekt połączenia i własny stan. Zawieszony Word nie blokuje Excela, a każde wywołanie
COM ma limit czasu (domyślnie 15 s) — po jego przekroczeniu wątek jest porzucany, połączenie
oznaczane jako martwe i odtwarzane przy następnym wywołaniu.

**Leniwe łączenie.** Aplikacja startuje dopiero przy pierwszym narzędziu, które jej dotyczy.
Bridge najpierw próbuje `GetActiveObject` (podłączenie do okna otwartego przez użytkownika),
a dopiero potem `Dispatch` (uruchomienie nowej instancji).

## Wymagania

| Element | Wersja |
|---|---|
| System | Windows 10/11 (**tylko Windows** — COM Office nie istnieje gdzie indziej) |
| Office | Microsoft Office 2019 (działa też z 2016/365 desktop) |
| Python | 3.11+ |
| Biblioteki | `mcp`, `pywin32`, `pytest` |

## Instalacja

```powershell
git clone https://github.com/JulianPoleszczuk/office-mcp.git
cd office-mcp

python -m venv .venv
.\.venv\Scripts\Activate.ps1

pip install -r requirements.txt
python .venv\Scripts\pywin32_postinstall.py -install
```

`pywin32_postinstall.py -install` rejestruje biblioteki DLL potrzebne do COM. Bez tego kroku
`win32com.client` potrafi rzucać `ImportError: DLL load failed` przy pierwszym `Dispatch`.

Szybki test, że wszystko widzi Office:

```powershell
python -m bridge.main --log-level DEBUG
```

## Konfiguracja Claude Desktop

Plik `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "office": {
      "command": "C:\\sciezka\\do\\office-mcp\\.venv\\Scripts\\python.exe",
      "args": ["C:\\sciezka\\do\\office-mcp\\server.py"],
      "env": {
        "OFFICE_BRIDGE_PORT": "8765"
      }
    }
  }
}
```

Po zapisaniu pliku zrestartuj Claude Desktop. Narzędzia `ppt_*`, `xl_*`, `doc_*` pojawią się
na liście dostępnych narzędzi.

W Claude Code wystarczy:

```powershell
claude mcp add office -- C:\sciezka\do\office-mcp\.venv\Scripts\python.exe C:\sciezka\do\office-mcp\server.py
```

Bridge startuje automatycznie przy pierwszym użyciu narzędzia — nie trzeba go uruchamiać ręcznie.

### Zmienne środowiskowe

| Zmienna | Domyślnie | Znaczenie |
|---|---|---|
| `OFFICE_BRIDGE_HOST` | `127.0.0.1` | Adres nasłuchu Bridge |
| `OFFICE_BRIDGE_PORT` | `8765` | Port Bridge |
| `OFFICE_BRIDGE_TIMEOUT` | `15` | Limit pojedynczego wywołania COM (s) |
| `OFFICE_BRIDGE_AUTOSTART` | `1` | `0` wyłącza automatyczny start Bridge przez serwer MCP |
| `OFFICE_MCP_CALL_TIMEOUT` | `60` | Limit oczekiwania serwera MCP na odpowiedź Bridge (s) |
| `OFFICE_BRIDGE_LOG_LEVEL` | `INFO` | Poziom logów Bridge |

## Uruchamianie Bridge osobno (debugowanie)

```powershell
python -m bridge.main --port 8765 --timeout 15 --log-level DEBUG
```

Bridge loguje każde żądanie i odpowiedź. Można go odpytać zwykłym socketem — jedna linia JSON
to jedno żądanie:

```powershell
python -c "import socket, json; s=socket.create_connection(('127.0.0.1',8765)); s.sendall(json.dumps({'id':'1','app':'excel','action':'get_workbook_info','params':{}}).encode()+b'\n'); print(s.recv(65536).decode())"
```

Format protokołu:

```json
{"id": "uuid", "app": "powerpoint", "action": "add_slide", "params": {"layout": "title_content"}}
{"id": "uuid", "ok": true, "result": {"slide_index": 2}}
{"id": "uuid", "ok": false, "error": {"type": "ComConnectionError", "message": "PowerPoint nie odpowiada"}}
```

## Narzędzia

Wszystkie narzędzia zwracają JSON w jednym formacie:
`{"ok": true, "result": {...}}` albo `{"ok": false, "error": {"type": ..., "message": ...}}`.

### PowerPoint (`ppt_*`)

| Narzędzie | Opis |
|---|---|
| `ppt_create_presentation(path, template=None)` | Nowa prezentacja, opcjonalnie z szablonu `.potx` |
| `ppt_open_presentation(path)` | Otwiera plik lub aktywuje już otwarty |
| `ppt_save(path=None)` | `Save` albo `SaveAs` |
| `ppt_close(save=True)` | Zamyka prezentację |
| `ppt_get_presentation_info()` | Liczba slajdów, rozmiar slajdu, motyw, ścieżka |
| `ppt_get_slide_content(slide_index)` | Kształty, pozycje, teksty, notatki |
| `ppt_list_slides()` | Tytuły i układy wszystkich slajdów |
| `ppt_add_slide(layout, index=None, title=None)` | Dodaje slajd o wybranym układzie |
| `ppt_delete_slide(slide_index)` | Usuwa slajd |
| `ppt_duplicate_slide(slide_index)` | Duplikuje slajd |
| `ppt_reorder_slide(from_index, to_index)` | Przenosi slajd |
| `ppt_set_title(slide_index, text)` | Ustawia tytuł slajdu |
| `ppt_add_textbox(slide_index, text, left, top, width, height, ...)` | Pole tekstowe |
| `ppt_add_bullet_list(slide_index, items, placeholder="content")` | Lista punktowana z poziomami |
| `ppt_find_replace_text(old_text, new_text, slide_index=None, match_case=False)` | Podmiana tekstu (też w tabelach i grupach) |
| `ppt_set_speaker_notes(slide_index, text)` | Notatki prelegenta |
| `ppt_set_text_style(slide_index, shape_id, ...)` | Czcionka, rozmiar, kolor, pogrubienie |
| `ppt_apply_theme(theme_name_or_path)` | Motyw z pliku `.thmx`/`.potx` lub galerii Office |
| `ppt_set_background(slide_index, color=None, image_path=None)` | Tło slajdu |
| `ppt_set_slide_layout(slide_index, layout_name)` | Zmiana układu |
| `ppt_add_image(slide_index, image_path, left, top, width=None, height=None)` | Obraz |
| `ppt_add_chart(slide_index, chart_type, categories, series_data, ...)` | Wykres z danymi |
| `ppt_add_table(slide_index, rows, cols, data, left, top, width, height)` | Tabela |
| `ppt_add_shape(slide_index, shape_type, left, top, width, height, ...)` | Kształt |

Układy: `title`, `title_content`, `two_content`, `title_only`, `blank`, `section_header`,
`comparison`, `picture_with_caption`, `content_with_caption`, `chart`, `table`, `four_objects`.
Wykresy: `bar`, `column`, `line`, `pie`, `area`, `scatter`, `doughnut`, `radar`, `bubble`.
Kształty: `rectangle`, `rounded_rectangle`, `oval`, `triangle`, `diamond`, `star`,
`arrow_right`, `callout`, `cloud`, `hexagon`, `chevron`.

Współrzędne podaje się w punktach: slajd 16:9 ma 960 × 540 pt, 1 cm = 28,35 pt.

### Excel (`xl_*`)

| Narzędzie | Opis |
|---|---|
| `xl_create_workbook(path)` | Nowy skoroszyt |
| `xl_open_workbook(path)` | Otwiera plik lub aktywuje już otwarty |
| `xl_save(path=None)` | `Save` albo `SaveAs` |
| `xl_close(save=True)` | Zamyka skoroszyt |
| `xl_add_sheet(name, index=None)` | Nowy arkusz |
| `xl_delete_sheet(name)` | Usuwa arkusz |
| `xl_rename_sheet(old_name, new_name)` | Zmienia nazwę arkusza |
| `xl_get_workbook_info()` | Arkusze, ich zakresy danych, aktywny arkusz, ścieżka |
| `xl_get_range_values(sheet, range_ref)` | Odczyt zakresu jako tablica 2D |
| `xl_get_used_range(sheet)` | Faktycznie wypełniony obszar wraz z danymi |
| `xl_set_cell(sheet, cell_ref, value)` | Wartość pojedynczej komórki |
| `xl_set_range(sheet, start_cell, values_2d)` | Wklejenie całej macierzy naraz |
| `xl_set_formula(sheet, cell_ref, formula)` | Formuła + wyliczony wynik |
| `xl_clear_range(sheet, range_ref, contents_only=True)` | Czyszczenie zakresu |
| `xl_insert_rows(sheet, start_row, count=1)` | Wstawianie wierszy |
| `xl_delete_rows(sheet, start_row, count=1)` | Usuwanie wierszy |
| `xl_insert_columns(sheet, start_col, count=1)` | Wstawianie kolumn (litera albo numer) |
| `xl_set_cell_format(sheet, range_ref, ...)` | Czcionka, kolory, format liczb, wyrównanie, zawijanie |
| `xl_set_column_width(sheet, column, width)` | Szerokość kolumny, `width="auto"` = autodopasowanie |
| `xl_merge_cells(sheet, range_ref, center=True)` | Scalanie komórek |
| `xl_apply_conditional_formatting(sheet, range_ref, rule_type, params)` | Formatowanie warunkowe |
| `xl_freeze_panes(sheet, cell_ref)` | Blokowanie okienek |
| `xl_add_chart(sheet, chart_type, data_range, left, top, width, height, title=None)` | Wykres |
| `xl_create_table(sheet, range_ref, table_name, has_headers=True, style=...)` | Natywna tabela Excela |
| `xl_add_pivot_table(sheet, source_range, dest_cell, rows, columns, values, ...)` | Tabela przestawna |

Reguły formatowania warunkowego: `cell_value` (operatory `greater`, `less`, `equal`,
`not_equal`, `greater_equal`, `less_equal`, `between`, `not_between`), `expression`,
`text_contains`, `color_scale`, `data_bar`.
Funkcje agregujące tabeli przestawnej: `sum`, `count`, `average`, `max`, `min`, `product`,
`count_numbers`, `std_dev`.

### Word (`doc_*`)

| Narzędzie | Opis |
|---|---|
| `doc_create_document(path, template=None)` | Nowy dokument, opcjonalnie z szablonu `.dotx` |
| `doc_open_document(path)` | Otwiera plik lub aktywuje już otwarty |
| `doc_save(path=None)` | `Save` albo `SaveAs` |
| `doc_close(save=True)` | Zamyka dokument |
| `doc_get_document_info()` | Liczba stron, słów, znaków, szablon, ścieżka |
| `doc_get_full_text()` | Cały tekst dokumentu |
| `doc_get_outline()` | Drzewo nagłówków z indeksami akapitów |
| `doc_add_paragraph(text, style=None)` | Akapit na końcu dokumentu |
| `doc_add_heading(text, level=1)` | Nagłówek poziomu 1–9 |
| `doc_insert_page_break()` | Twardy podział strony |
| `doc_find_replace(old_text, new_text, match_case=False)` | Podmiana tekstu |
| `doc_add_bullet_list(items)` | Lista punktowana z poziomami |
| `doc_add_numbered_list(items)` | Lista numerowana |
| `doc_set_text_style(paragraph_index, ...)` | Czcionka, rozmiar, kolor, pogrubienie, kursywa |
| `doc_set_paragraph_alignment(paragraph_index, alignment)` | `left`/`center`/`right`/`justify` |
| `doc_apply_style(paragraph_index, style_name)` | Styl akapitu |
| `doc_set_page_margins(top, bottom, left, right, unit="cm")` | Marginesy (`cm`, `mm`, `in`, `pt`) |
| `doc_insert_image(image_path, width=None, height=None, position="inline")` | Obraz w tekście lub pływający |
| `doc_insert_table(rows, cols, data=None, position=None)` | Tabela z obramowaniem |
| `doc_insert_header(text, section=1)` | Nagłówek strony |
| `doc_insert_footer(text, section=1)` | Stopka |
| `doc_add_page_numbers(alignment="center", first_page=True)` | Numery stron |
| `doc_insert_table_of_contents(levels=3, position="start")` | Spis treści ze stylów nagłówków |

Nazwy stylów można podawać po angielsku (`Heading 1`, `Normal`, `Quote`, `Caption`) także w
polskiej wersji Worda — kontroler mapuje je na wbudowane stałe `wdStyle`.

### Diagnostyka

| Narzędzie | Opis |
|---|---|
| `office_status()` | Stan Bridge i połączeń COM wszystkich trzech aplikacji |

## Obsługa błędów

Kontrolery nigdy nie przepuszczają surowego `pywintypes.com_error` — każdy błąd COM jest
mapowany na typ z `bridge/utils/errors.py`:

| Typ | Kiedy |
|---|---|
| `ComConnectionError` | Aplikacja zamknięta, nie odpowiada albo odrzuciła wywołanie (otwarty dialog) |
| `ComTimeoutError` | Wywołanie COM przekroczyło limit czasu |
| `DocumentNotFoundError` | Plik nie istnieje, katalog docelowy nie istnieje, brak otwartego dokumentu |
| `InvalidReferenceError` | Zły `slide_index`, nieistniejący arkusz, zły zakres, nieznana nazwa układu/wykresu |
| `UnsupportedOperationError` | Operacja niedostępna w zainstalowanej wersji Office |
| `ProtocolError` | Zła wiadomość protokołu, nieznana akcja, brakujące parametry |
| `BridgeUnavailable` | Serwer MCP nie dogadał się z procesem Bridge |

Przykładowa odpowiedź błędu narzędzia MCP:

```json
{
  "ok": false,
  "error": {
    "type": "InvalidReferenceError",
    "message": "slide_index = 7 poza zakresem 1..3"
  }
}
```

## Testy

```powershell
python -m pytest -q
```

216 testów, wszystkie bez zainstalowanego Office:

- `tests/test_bridge_protocol.py` — kodowanie/dekodowanie protokołu oraz test integracyjny
  serwera TCP (prawdziwy socket, atrapa kontrolera),
- `tests/test_powerpoint_controller.py`, `test_excel_controller.py`, `test_word_controller.py` —
  kontrolery z zamockowanym COM (`unittest.mock`),
- `tests/test_server_tools.py` — warstwa MCP; sprawdza między innymi, że każda akcja użyta
  w `server.py` istnieje w odpowiednim kontrolerze.

Scenariusze do ręcznego testu na żywym Office: `examples/example_prompts.md`.

## Znane ograniczenia

- **Tylko Windows.** Automatyzacja opiera się na COM; `server.py` i `bridge/main.py` kończą
  pracę czytelnym komunikatem na innych platformach.
- **Wymaga zainstalowanego Office** (desktop, nie wersji web). Bridge podłącza się do otwartej
  instancji albo uruchamia nową.
- **COM bywa kruche między wersjami.** Ta sama metoda potrafi być raz metodą, raz właściwością —
  stąd np. helper `com_address`. Wersje starsze niż 2016 mogą nie mieć `AddChart2`
  (jest fallback na `AddChart`) ani niektórych typów formatowania warunkowego.
- **Otwarte okno dialogowe blokuje aplikację.** Jeśli użytkownik ma otwarty np. dialog
  zapisu, Office odrzuca wywołania COM — narzędzie zwróci `ComConnectionError` z podpowiedzią.
- **Motywy PowerPointa** przyjmuje się jako ścieżkę do `.thmx`/`.potx` albo nazwę z galerii
  Office; COM nie udostępnia motywów po samej nazwie.
- **Zamiana tekstu w Wordzie** dopasowuje wielkość liter wstawianego tekstu do znalezionego,
  gdy `match_case=False` — tak samo jak okno „Znajdź i zamień”.
- **Zapis** wymaga ścieżki przy pierwszym zapisaniu nowego dokumentu (`ppt_save(path=...)`).
- Bridge nasłuchuje wyłącznie na localhost i nie ma uwierzytelniania — nie należy wystawiać
  jego portu na zewnątrz.

## Struktura projektu

```
office-mcp/
├── server.py                    # serwer MCP (stdio), rejestracja narzędzi, klient Bridge
├── bridge/
│   ├── main.py                  # serwer TCP, protokół JSON-line, routing do kontrolerów
│   ├── protocol.py              # Request/Response + kodowanie linii
│   ├── connection_manager.py    # leniwe łączenie COM, wątki STA, timeouty
│   ├── controllers/
│   │   ├── base.py              # routing akcji, mapowanie com_error, wspólne helpery
│   │   ├── powerpoint.py        # PowerPointController
│   │   ├── excel.py             # ExcelController
│   │   └── word.py              # WordController
│   └── utils/
│       ├── com_helpers.py       # kolory, jednostki, stałe Office, konwersje wartości
│       └── errors.py            # hierarchia wyjątków
├── tests/                       # testy protokołu, kontrolerów (mock COM) i warstwy MCP
└── examples/example_prompts.md  # scenariusze end-to-end na żywym Office
```

## Licencja

MIT
