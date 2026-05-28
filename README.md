# ISTQB Academia PDF Aggregator

**Aktuální verze:** 0.14b  
**Datum vydání:** 2026-05-28  
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

## Instalace & spuštění (macOS, iCloud-friendly)

> ⚠️ Repo leží v iCloud-synchronizované složce (Desktop). **Venv proto patří mimo
> iCloud** – jinak iCloud vytvoří duplicitní `… 2.*` soubory a rozbije instalaci
> PySide6 (typicky chyba `Could not find the Qt platform plugin "cocoa"`).
> Venv postavíme v `~/.venvs/` a v projektu na něj uděláme **symlink `.venv`**,
> takže běžné `source .venv/bin/activate` funguje dál.

```bash
# 1) Mimo iCloud postav venv
mkdir -p ~/.venvs
deactivate 2>/dev/null
rm -rf .venv
/opt/anaconda3/bin/python3 -m venv ~/.venvs/istqb-academia   # nebo jiný python3.10+

# 2) V projektu vytvoř symlink na ten venv
ln -s ~/.venvs/istqb-academia .venv

# 3) Aktivuj a nainstaluj závislosti
source .venv/bin/activate
pip install PySide6 pypdf PyPDF2 pdfminer.six openpyxl

# 4) Spusť
python main.py --pdf-root "/cesta/k/kořeni/PDF"
```

> Pozn.: pokud máš v aktuálním terminálu starou aktivaci (`(.venv)` v promptu),
> nejdřív `deactivate` a pak znovu `source .venv/bin/activate`.
- `--pdf-root` míří na **kořen adresáře s PDF** (podadresáře = boards jako `CFTL`, `LTSTQB`, …). Je **nepovinný** — bez něj se použije naposledy uložená složka, jinak výchozí `./PDF`.  
- Složku s PDF i složku **Sorted PDFs** lze kdykoli změnit v menu **File** (volba se uloží).  
- Při prvním spuštění se vytvoří (pokud chybí) adresář **Sorted PDFs** pro lokální DB a exporty.

### Nastavení (settings.json)
Aplikace si pamatuje nastavení mezi spuštěními v souboru `settings.json` v konfiguračním adresáři OS
(`QStandardPaths.AppConfigLocation`, na macOS typicky `~/Library/Application Support/istqb-academia-aggregator/`).
Ukládá se: poslední PDF složka, složka Sorted PDFs, geometrie okna, poslední záložka a Overview filtry.
Soubor je čistě lokální (není verzovaný).

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

Získávání probíhá primárně z **AcroForm polí**. Pokud PDF žádná pole nemá
(flattened/„zploštělý" formulář), použije se od 0.13 **textový fallback** nad
textovou vrstvou. U naskenovaných PDF bez textové vrstvy zůstanou hodnoty
prázdné a PDF se v UI **označí** k ručnímu vyplnění (viz Changelog 0.13).

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
### 0.14b — 2026-05-28
- **feat(overview):** **zpětná vazba editovaných hodnot do tabulky Overview.** Hodnoty doplněné/změněné ručně v záložce Sorted PDFs se nyní promítnou přímo do buněk Overview (zeleně zvýrazněný text + tooltip „Edited in Sorted PDFs"), včetně ikonek Yes/No u *Academia/Certified Recognition*. Aktualizuje se hned po **Save to DB** i při rescanu.

### 0.14a — 2026-05-28
- **fix(export):** **Board** se při exportu z Overview určoval substringově (`"board"`), takže ho chybně „přebil" sloupec *Receiving Member Board* → PDF se ukládalo pod **Unsorted**. Nově se bere přesně sloupec *Board* (poslední řádek popisku). Opraveno mj. pro nové boardy (např. `FISTB`).
- **fix(export):** export nově zapisuje záznam jako **parsed (edited=False)** přes `upsert_parsed` (dřív `mark_edited`). Stav **„Edited"** v Overview se tak objeví **až po skutečné ruční editaci** v Sorted PDFs; opakovaný export navíc **nepřepíše** již ručně doplněné hodnoty.

### 0.14 — 2026-05-28
- **fix(sorted):** oprava kritické chyby ve `rescan_sorted` – položky stromu měly **chybnou absolutní cestu** (chyběl segment `Sorted PDFs/`), takže `sorted_db.get()` házel `ValueError`. To rozbíjelo výběr ve stromu, **Edit → Save to DB** i zobrazení nově **exportovaných** PDF z Overview. Nyní se klíč DB správně převádí na absolutní cestu pod `sorted_root`.
- **feat(overview):** sloupec **Sorted** nově rozlišuje tři stavy: prázdné (není v Sorted), **„Yes"** (zkopírováno do Sorted) a **„Edited"** (ručně doplněno v záložce Sorted PDFs). Editované záznamy mají **zelené pozadí** a **tooltip se seznamem polí**, která byla v Sorted doplněna/změněna oproti automaticky vytěženým hodnotám.
- **i18n:** uživatelské texty kolem exportu a varování přepsány do **angličtiny** (kontextová akce „Export to 'Sorted PDFs' folder", stavové hlášky, varování u skenů).

### 0.13a — 2026-05-28
- **docs/build:** popsán **iCloud-friendly venv** postup (venv v `~/.venvs/`, v projektu symlink `.venv`). Řeší opakovanou chybu `Could not find the Qt platform plugin "cocoa"` způsobenou tím, že iCloud duplikuje soubory v in-folder `.venv`.
- **chore(gitignore):** ignoruj i symlink `.venv` (nejen adresář `.venv/`).

### 0.13 — 2026-05-28
- **feat(parser):** nový textový fallback pro **flattened PDF bez AcroForm polí**. Když formulář nemá interaktivní pole, vytěží se kontaktní blok (instituce, kandidát, jméno, e‑mail, telefon, adresa), datum podpisu a *Printed Name, Title* přímo z textové vrstvy. Podporuje dvě rozložení: hodnoty v blocích za labely i hodnoty zřetězené na jednom řádku (s očištěním e‑mailu od slepeného jména). Spouští se **jen** při prázdném AcroFormu, takže nemění chování běžných PDF.
- **feat(scanner):** kontaktní pole se nově doplňují z tohoto fallbacku (dřív se braly jen z AcroFormu).
- **feat(ui):** PDF bez AcroFormu a bez vytěžitelných hodnot (naskenovaný obraz / prázdný nebo poškozený formulář) se nově **označí**: v **PDF Browseru** se zobrazí varovný banner, v **Sorted PDFs** stavový text „⚠ Sken/prázdný formulář – vyplňte ručně". Označení zmizí, jakmile jsou hodnoty doplněny.
- **fix(parser):** šablonový boilerplate sekce 5 se už nevyhodnocuje jako *Additional information* (odstraněn falešný pozitiv u prázdných formulářů).

### 0.12a — 2026-05-28
- **feat(ui):** okno se při startu **defaultně maximalizuje** (není-li uložená geometrie okna). S uloženou geometrií se chování nemění — respektuje se poslední velikost/pozice.
- **poznámka(prostředí):** chyba startu `Could not find the Qt platform plugin "cocoa"` byla způsobena poškozenou instalací PySide6 ve `.venv`. Řešení: `pip install --force-reinstall --no-cache-dir PySide6==6.11.1`.

### 0.12 — 2026-05-28
- **feat(settings):** PDF kořen i složka **Sorted PDFs** jsou nově **vybíratelné za běhu** přes menu **File → Open PDF folder… / Open Sorted PDFs folder…** (`QFileDialog`).
- **feat(settings):** aplikace si **pamatuje nastavení** mezi spuštěními v JSON souboru (`settings.json` v config adresáři OS, `QStandardPaths.AppConfigLocation`): poslední PDF složka, složka Sorted PDFs, velikost/pozice okna, poslední záložka a Overview filtry (hledání, board, *Hide Sorted*).
- **fix:** `sorted_root` už není natvrdo vázán na aktuální pracovní adresář — bere se z uloženého nastavení (fallback: `./Sorted PDFs`).
- Priorita PDF kořene: argument `--pdf-root` > uložené nastavení > výchozí `./PDF`.

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
