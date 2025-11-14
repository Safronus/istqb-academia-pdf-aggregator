# ISTQB Academia PDF Aggregator

**Aktuální verze:** 0.11t  
**Datum vydání:** 2025-11-14  
**Platforma:** macOS (PySide6, dark‑theme friendly)

Aplikace slouží k vyhledávání, parsování, třídění a exportu údajů z
*ISTQB Academia Recognition Program – Application Form* (PDF). Vše je navrženo
jako **minimal‑change** nad stávajícím projektem: zachováváme původní chování,
UI, názvy widgetů a datové struktury; přidáváme pouze to, co bylo výslovně
požadováno.

---

## Obsah
- [Požadavky](#požadavky)
- [Instalace & spuštění (macOS)](#instalace--spuštění-macos)
- [Struktura aplikace](#struktura-aplikace)
- [Nová pole z PDF (Section 6 a 7)](#nová-pole-z-pdf-section-6-a-7)
- [Záložka Overview](#záložka-overview)
- [Záložka PDF Browser](#záložka-pdf-browser)
- [Záložka Sorted PDFs](#záložka-sorted-pdfs)
- [Exporty (XLSX/CSV/TXT)](#exporty-xlsxcsvtxt)
- [Databáze `sorted_db.json`](#databáze-sorted_dbjson)
- [Verzování a git příkazy](#verzování-a-git-příkazy)
- [Changelog od 0.11](#changelog-od-011)
- [Smoke test (rychlé ověření)](#smoke-test-rychlé-ověření)

---

## Požadavky
- **Python 3.10+** (doporučeno 3.11)
- **PySide6**
- Ostatní závislosti dle původního projektu (nepřidáváme nové knihovny bez důvodu).

> Aplikace je vyvíjena s ohledem na macOS (HiDPI/Retina, dark theme).

---

## Instalace & spuštění (macOS)
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt  # je-li k dispozici; jinak: pip install PySide6
python main.py --root "/cesta/k/kořeni/PDF"
```
- `--root` míří na **kořen adresáře s PDF** (podadresáře = boards jako `CFTL`, `LTSTQB`, …).  
- Při prvním spuštění se vytvoří (pokud chybí) adresář **Sorted PDFs** pro lokální DB a exporty.

---

## Struktura aplikace
- **Overview** – tabulka všech nalezených PDF (skupiny sloupců, možnost filtrování a exportu).
- **PDF Browser** – strom souborů (nahoře) a panel detailů (dole).
- **Sorted PDFs** – vlastní databáze vybraných PDF (pravý panel je editovatelný; hodnoty se ukládají do `sorted_db.json`).

Design, názvy a pořadí prvků zachováváme dle původního projektu. Novinky jsou
vložené tak, aby **nerušily** stávající používání.

---

## Nová pole z PDF (Section 6 a 7)
Z aplikace 0.11 umíme číst a zobrazovat **další pole z PDF** (mohou být prázdná; nic si nedomýšlíme):

**Section 6 – Declaration and Consent**
- `Printed Name, Title`

**Section 7 – For ISTQB Academia Use Only**
- `Receiving Member Board`
- `Date Received`
- `Validity Start Date`
- `Validity End Date` *(libovolný text, ne jen datum)*

Získávání probíhá **pouze z AcroForm polí**; když pole v PDF chybí, hodnota zůstane prázdná.

---

## Záložka Overview
- `Printed Name, Title` je **před** `Signature Date`.
- Sekce 7 (`Receiving Member Board`, `Date Received`, `Validity Start Date`, `Validity End Date`)
  je sdružena **za** `Signature Date` pod nadpisem *For ISTQB Academia Purpose Only*.
- Nové sloupce jsou **barevně** odlišené: blok Consent a blok ISTQB‑internal mají jednotnou barvu.
- Sloupec **Sorted** zůstal **poslední** a plní se jako dříve (bez změn logiky).

---

## Záložka PDF Browser
- Levý strom zobrazuje **podsložky kořene** a soubory (výchozí chování `QFileSystemModel`).  
- **Abecedné řazení** je zapnuté; po načtení se jemně přizpůsobí šířky sloupců.  
- Pravý detail je **sdružen do sekcí 1–7** dle formuláře; hlavičky sekcí se vkládají po sestavení formuláře.

---

## Záložka Sorted PDFs
- Levý strom: skupiny podle **Board**, položky = soubory.  
- Pravý formulář: **všechna** dosavadní pole + nová (Consent + ISTQB‑internal).  
- `File name` je **read‑only**. `Board` je také read‑only.  
- Tlačítko **Edit** přepne editovatelnost polí; **Save to DB** uloží změny do `sorted_db.json`.  
- Levý strom je **abecedně seřazen** – **boards** i **PDF soubory**.

---

## Exporty (XLSX/CSV/TXT)
- V **Overview** i **Sorted PDFs** lze exportovat vybrané sloupce do **.xlsx**, **.csv** i **.txt**.  
- TXT export má tematické bloky včetně **Consent** a **For ISTQB Academia Purpose Only**.  
- Pořadí sloupců kopíruje Overview (včetně nových).

---

## Databáze `sorted_db.json`
- Schéma beze změn (verze `1.0`). Nová pole se ukládají do bloku `data`.  
- `Rescan Sorted` strom staví ze záznamů DB; interně používáme **DB‑klíč** relativní k adresáři *Sorted PDFs*,
  aby nedocházelo k chybám při výběru PDF mimo tento adresář.
- Při otevírání detailu se podle potřeby **doplní chybějící klíče** z čerstvého parsování PDF (jen pro UI).

---

## Verzování a git příkazy
Projekt používá schéma **major.minor + patch‑písmeno**. Hlavní verzi neměníme bez výslovné žádosti.

```bash
# ulož změny
git add app/*.py README.md

# sémantický commit
git commit -m "fix(sorted): add missing 'File name' field and status label; keep alphabetical sorting (v0.11s)"

# tag patch verzi
git tag -a v0.11s -m "v0.11s"

# push do repozitáře
git push && git push --tags

# zobraz hash posledního commitu (po push)
git rev-parse HEAD
```

---

## Changelog od 0.11
### 0.11t — 2025-11-14
- fix(parser): textové fallbacky pro `University website links`, `Any additional relevant information or documents` a `Printed Name, Title`. Formulářová pole mají stále prioritu.
- fix(scanner): při prázdných hodnotách z AcroForm doplní `uni_links` a `additional_information` z parseru.

### 0.11s — 2025-11-14
- **Sorted PDFs:** doplněno chybějící pole **`File name`** (read‑only) a **status label** `lbl_sorted_status` vpravo dole.  
  Odstraňuje chyby `AttributeError: ed_filename` a `AttributeError: lbl_sorted_status`.  
  Abecední řazení stromu zůstává zapnuté.

### 0.11r
- **Sorted PDFs:** strom abecedně řadí boards i PDF soubory (zapnuto `setSortingEnabled(True)` + `sortItems(...)`).

### 0.11q
- **PDF Browser:** po sestavení pravého formuláře se vkládají **sekční hlavičky** (1–7).

### 0.11p
- **PDF Browser:** návrat k výchozímu chování `QFileSystemModel` (bez filtrů), přidáno pouze řazení.

*(starší položky viz předchozí verze)*

---

## Smoke test (rychlé ověření)
1. Spusť aplikaci s `--root` ke kořeni PDF (adresář s podsložkami boardů).  
2. V **Sorted PDFs** klikni **Rescan Sorted** → top‑level boards i jejich PDF položky jsou abecedně.  
3. Vyber libovolné PDF v **Sorted PDFs** → pravý formulář ukazuje i **File name** a stav dole.  
4. V **PDF Browseru** vlevo vidíš strom; klik na PDF → vpravo náhled se sekcemi 1–7.  
5. Ověř Overview/Sorted exporty (XLSX/CSV/TXT).  
6. Prázdná PDF pole zůstávají prázdná; `Validity End Date` je libovolný text.
