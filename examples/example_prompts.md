# Przykładowe polecenia (test end-to-end na żywym Office 2019)

Scenariusze do ręcznego sprawdzenia w Claude Desktop / Claude Code z podłączonym `office-mcp`.
Każdy opisuje, co wpisać, jakich narzędzi model powinien użyć i co należy zobaczyć w oknie Office.

Przed startem: `office_status` powinien zwrócić `"ok": true` dla Bridge. Ścieżki w przykładach
warto podmienić na własne.

---

## 1. Prezentacja od zera z wykresem

> Stwórz prezentację o robotyce sumo w `C:\Users\ja\Documents\sumo.pptx`, 5 slajdów:
> tytułowy, zasady, budowa robota, wyniki z wykresem słupkowym (Robot A – 5 wygranych,
> Robot B – 3, Robot C – 8) i podsumowanie. Do slajdu z zasadami dodaj listę punktowaną.

Narzędzia: `ppt_create_presentation`, `ppt_add_slide`, `ppt_set_title`, `ppt_add_bullet_list`,
`ppt_add_chart`, `ppt_save`.
Sprawdź: 5 slajdów, wykres z podpisanymi seriami, punkty na slajdzie 2.

---

## 2. Edycja istniejącej prezentacji

> Otwórz `C:\prezentacje\raport_q3.pptx`, wypisz tytuły slajdów, a potem na slajdzie 3
> zmień tytuł na „Wyniki Q3 2026” i dodaj notatkę prelegenta z trzema zdaniami komentarza.

Narzędzia: `ppt_open_presentation`, `ppt_list_slides`, `ppt_set_title`, `ppt_set_speaker_notes`.
Sprawdź: model najpierw odczytuje strukturę, dopiero potem edytuje; okno przewija się na slajd 3.

---

## 3. Podmiana tekstu w całej prezentacji

> W otwartej prezentacji zamień wszystkie wystąpienia „2025” na „2026”, także w tabelach,
> i powiedz ile zmian zrobiłeś.

Narzędzia: `ppt_find_replace_text`.
Sprawdź: `replacements` zgadza się z liczbą wystąpień, tabele też zostały zaktualizowane.

---

## 4. Slajd składany ręcznie z obiektów

> Na nowym pustym slajdzie ułóż: pole tekstowe „Harmonogram” u góry (32 pt, pogrubione),
> tabelę 3×2 z etapami projektu po lewej, a po prawej zaokrąglony prostokąt w kolorze
> `#0070C0` z napisem „Start: marzec”.

Narzędzia: `ppt_add_slide(layout="blank")`, `ppt_add_textbox`, `ppt_add_table`, `ppt_add_shape`.
Sprawdź: elementy nie nachodzą na siebie (slajd 16:9 to 960 × 540 pt).

---

## 5. Motyw, tło i formatowanie

> Ustaw tło slajdu 1 na `#F2F2F2`, tytuł zmień na 40 pt w kolorze granatowym,
> a slajdowi 2 nadaj układ „Porównanie”.

Narzędzia: `ppt_get_slide_content` (po `shape_id`), `ppt_set_background`, `ppt_set_text_style`,
`ppt_set_slide_layout`.
Sprawdź: model najpierw pobiera `shape_id` z zawartości slajdu, zamiast go zgadywać.

---

## 6. Arkusz budżetu z formatowaniem warunkowym

> W arkuszu `budzet.xlsx` dodaj kolumnę „Suma” z formułą sumującą wiersz,
> pogrub nagłówki, ustaw format walutowy na kwotach i pokoloruj na czerwono
> wszystkie komórki większe niż 1000.

Narzędzia: `xl_open_workbook`, `xl_get_used_range`, `xl_set_formula`, `xl_set_cell_format`,
`xl_apply_conditional_formatting`.
Sprawdź: reguła `cell_value` z operatorem `greater` i progiem 1000; format liczb w kolumnie kwot.

---

## 7. Wklejenie całej tabeli danych naraz

> Stwórz skoroszyt `C:\dane\wyniki.xlsx` z arkuszem „Zawody” i wstaw dane:
> nagłówki Robot / Kategoria / Punkty oraz pięć wierszy wyników. Zablokuj pierwszy wiersz
> i dopasuj szerokość kolumn.

Narzędzia: `xl_create_workbook`, `xl_rename_sheet`, `xl_set_range`, `xl_freeze_panes`,
`xl_set_column_width`.
Sprawdź: dane wchodzą jednym wywołaniem `xl_set_range`, a nie komórka po komórce.

---

## 8. Wykres i tabela Excela

> Z danych w `A1:C6` arkusza „Zawody” zrób wykres kolumnowy zatytułowany „Punkty robotów”
> obok danych, a sam zakres zamień w tabelę o nazwie „Wyniki”.

Narzędzia: `xl_add_chart`, `xl_create_table`.
Sprawdź: wykres nie zasłania danych, tabela ma filtry w nagłówkach.

---

## 9. Tabela przestawna

> W `budzet.xlsx` zbuduj na nowym arkuszu tabelę przestawną z zakresu `A1:D200`:
> w wierszach kategorie, w kolumnach miesiące, wartości to suma kwot.

Narzędzia: `xl_add_sheet`, `xl_add_pivot_table`.
Sprawdź: nazwy pól muszą pokrywać się z nagłówkami zakresu źródłowego — inaczej wraca
`InvalidReferenceError` z czytelnym komunikatem.

---

## 10. Odczyt i analiza danych

> Odczytaj zakres `A1:D50` z arkusza „Sprzedaz”, powiedz który miesiąc miał najwyższą sprzedaż
> i wpisz ten wynik do komórki F1 razem z opisem.

Narzędzia: `xl_get_range_values`, `xl_set_cell`.
Sprawdź: model analizuje zwrócone dane po swojej stronie, a wynik ląduje w arkuszu.

---

## 11. Raport w Wordzie od zera

> Stwórz dokument `C:\raporty\zawody.docx`: nagłówek 1 „Raport z zawodów”, akapit wstępu,
> nagłówek 2 „Wyniki” z listą numerowaną trzech pierwszych miejsc, tabelę 3×2 z punktacją,
> stopkę z numerami stron i spis treści na początku.

Narzędzia: `doc_create_document`, `doc_add_heading`, `doc_add_paragraph`, `doc_add_numbered_list`,
`doc_insert_table`, `doc_add_page_numbers`, `doc_insert_table_of_contents`.
Sprawdź: spis treści zawiera oba nagłówki, numeracja stron widoczna w stopce.

---

## 12. Porządkowanie istniejącego dokumentu

> Otwórz `raport.docx`, znajdź wszystkie wystąpienia „TODO” i zamień je na „Zrobione”,
> a potem pokaż strukturę nagłówków dokumentu.

Narzędzia: `doc_open_document`, `doc_find_replace`, `doc_get_outline`.
Sprawdź: liczba zamian w odpowiedzi; przy `match_case=False` Word dopasowuje wielkość liter
wstawianego tekstu do znalezionego.

---

## 13. Formatowanie akapitów i marginesów

> W otwartym dokumencie wyjustuj akapit 3, ustaw w nim czcionkę Calibri 12 pt,
> a marginesy strony zmień na 2 cm z każdej strony.

Narzędzia: `doc_get_outline` (po indeksy), `doc_set_paragraph_alignment`, `doc_set_text_style`,
`doc_set_page_margins`.
Sprawdź: indeksy akapitów model bierze z dokumentu, a nie z powietrza.

---

## 14. Dokument z obrazem i nagłówkiem firmowym

> Wstaw do dokumentu logo z `C:\grafika\logo.png` o szerokości 120 pt, dodaj nagłówek strony
> „Firma sp. z o.o.” i podziel dokument stroną przed sekcją „Załączniki”.

Narzędzia: `doc_insert_image`, `doc_insert_header`, `doc_insert_page_break`.
Sprawdź: brak pliku → `DocumentNotFoundError` zamiast wyjątku Pythona.

---

## 15. Przypadki błędne (celowo)

> Dodaj punktory do slajdu 99.
> Zapisz nowy skoroszyt bez podawania ścieżki.
> Wstaw wykres typu „spirala”.

Oczekiwane odpowiedzi:

- `InvalidReferenceError` — `slide_index = 99 poza zakresem 1..N`,
- `InvalidReferenceError` — „Skoroszyt nie ma jeszcze pliku - podaj parametr path”,
- `InvalidReferenceError` — nieznany typ wykresu z listą dostępnych.

Sprawdź: żadne narzędzie nie zwraca stack trace'a, a Office pozostaje sprawny — kolejne
polecenia działają dalej.

---

## 16. Test odporności połączenia

> Zamknij ręcznie okno PowerPointa, a potem poproś o `ppt_list_slides`.
> Następnie poproś o `xl_get_workbook_info`.

Oczekiwane: PowerPoint zwraca `ComConnectionError` albo `DocumentNotFoundError` i podnosi się
przy kolejnym poleceniu, a Excel działa niezależnie — awaria jednej aplikacji nie dotyka reszty.
