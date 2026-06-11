# IDoc → Excel Documentation Generator

Skrypt Python konwertujący pliki IDoc SAP (format flat file `.txt`) do czytelnej dokumentacji Excel.

## Obsługiwane typy IDoc

| Typ IDoc | Opis |
|----------|------|
| ORDERS05 | Zamówienie zakupu |
| DELFOR02 | Harmonogram dostaw (scheduling agreement release) |

## Wymagania

- Python 3.8+
- openpyxl

```bash
pip install openpyxl
```

## Użycie

```bash
# Podstawowe – plik wyjściowy generowany automatycznie
python idoc_parser.py ORDERS_344242684.txt

# Z podaną nazwą pliku wyjściowego
python idoc_parser.py DELFOR_344224115.txt moja_dokumentacja.xlsx
```

## Struktura wyjściowego pliku Excel

### ORDERS05

| Arkusz | Zawartość |
|--------|-----------|
| `IDoc - Pola i wartości` | Pełny rozpad każdego segmentu na pola: nazwa, długość, opis, wartość, znaczenie kodu |
| `Podsumowanie` | Lista wszystkich segmentów z liczbą wystąpień i opisem |
| `Pozycje zamówienia` | Zestawienie pozycji: materiał, ilość, cena, data dostawy, zakład |

### DELFOR02

| Arkusz | Zawartość |
|--------|-----------|
| `IDoc - Pola i wartości` | Pełny rozpad segmentów |
| `Podsumowanie` | Lista segmentów |
| `Harmonogram dostaw` | Linie E2EDP16 z datami i ilościami |
| `Historia dostaw` | Wpisy E2EDP36 – potwierdzenia GR/dostaw |

## Funkcje

- Automatyczne wykrycie typu IDoc z nagłówka `EDI_DC40`
- Kolumna **Znaczenie** dla pól kodowanych (`PARVW`, `QUALF`) – zielone tło
- Puste pola wyróżnione żółtym tłem
- Kodowanie latin-1 (obsługa znaków SAP)

## Słowniki kodów

Skrypt zawiera wbudowane słowniki znaczeń dla:

- `PARVW` – role partnerów (AG, LF, WE, EK, AP i inne)
- `QUALF` w `E2EDK14` – jednostki organizacyjne (zakład, org. zakupów itd.)
- `QUALF` w `E2EDK03` – kwalifikatory dat
- `QUALF` w `E2EDK02` – typy referencji dokumentów
- `QUALF` w `E2EDK17` / `E2EDP17` – warunki dostawy / Incoterms
- `QUALF` w `E2EDK18` – warunki płatności
- `QUALF` w `E2EDP19` – identyfikacja materiału
- `QUALF` w `ZE1EDKEMAIL` – typ adresu e-mail

## Dodawanie nowych segmentów

W słowniku `SEGMENT_FIELDS` dodaj nowy wpis:

```python
'E2EDKNEW': [
    ('POLE1', 3,  'Opis pola 1'),
    ('POLE2', 35, 'Opis pola 2'),
    # suma długości musi odpowiadać obszarowi danych segmentu
],
```

## Dodawanie nowych słowników kodów

```python
QUALF_MÓJSEG = {
    '001': 'Opis wartości 001',
    '002': 'Opis wartości 002',
}

# Dodaj do FIELD_LOOKUPS:
('E2EDKMOJ', 'QUALF'): QUALF_MÓJSEG,
```

## Format pliku IDoc (flat file)

```
<nazwa_segmentu:30><pola_kontrolne:33><dane_segmentu>
```

- Bajty 0–29: nazwa segmentu (np. `E2EDK01006`)
- Bajty 30–62: pola kontrolne (DOCNUM, SEGNUM itd.)
- Bajty 63+: dane segmentu – pola ułożone kolejno według zdefiniowanych długości

## Struktura projektu

```
idoc-to-excel/
├── idoc_parser.py    ← główny skrypt
├── README.md
├── requirements.txt
├── .gitignore
└── examples/         ← przykładowe pliki IDoc (opcjonalnie)
```
