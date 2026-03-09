# Validator kapacit

Desktopová aplikace pro automatickou kontrolu Excelových souborů s kapacitními daty.

Vytvořeno pro **[PARAMA Software](mailto:info@parama.cz)**.

---

## Stažení

**→ [Stáhnout ValidatorKapacit.exe](https://github.com/radek-livecka/capacity-excel-validator/releases/latest)**

Soubor `.exe` nevyžaduje instalaci Pythonu ani žádných dalších nástrojů — stačí stáhnout a spustit.

---

## Co aplikace dělá

Prochází zadaný Excel soubor a hledá neúplné nebo chybové záznamy v kapacitních datech:

- Kontrolují se listy, jejichž název **nezačíná** podtržítkem `_`
- Oblast kontroly: sloupce **A–H** od řádku **14**
- Pravidlo: pokud sloupec **B** obsahuje hodnotu (kód projektu), pak sloupce A a C–H **nesmí být prázdné** ani obsahovat Excel chybu (`#N/A`, `#REF!` apod.)
- Výsledky lze uložit jako `.txt` report

---

## Spuštění ze zdrojového kódu

Vyžaduje Python 3.9+.

```bash
pip install -r requirements.txt
python validator.py
```

## Build do .exe

```bash
build.bat
```

Skript nainstaluje závislosti, vygeneruje grafické assety a sestaví `dist/ValidatorKapacit.exe`.

---

## Závislosti

| Knihovna | Účel |
|---|---|
| `openpyxl` | Čtení Excel souborů |
| `Pillow` | Zobrazení loga v GUI |

---

## Struktura projektu

```
validator.py              # Hlavní aplikace
requirements.txt          # Závislosti
setup.bat                 # Instalace závislostí (pro spuštění ze zdroje)
build.bat                 # Sestavení .exe
graphics/
  01-parama-primary-dark.svg   # Logo PARAMA Software (zdroj)
  05-parama-symbol.svg         # Symbol PARAMA Software (zdroj)
  create_assets.py             # Generátor PNG assetů pro GUI
  parama-symbol.png            # Vygenerovaný asset (header)
  parama-icon.png              # Vygenerovaný asset (ikona okna)
```

---

*PARAMA Software — info@parama.cz*
