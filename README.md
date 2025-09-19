# Zapoctovy-program---tvroba-rozvrhu
Stručný popis:
Program načte požadavky na školní rozvrh z excelové tabulky, najde nejlépe vyhovující řešení (pokud nějaké existuje) a to vrátí jednak do konzole a jednak do excelového souboru.
# Dokumentace
# Instalace & spuštění:
Ke spuštění programu je nejprve nutno nainstalovat python 3 a do něj potřebné knihovny. Tyto instalujeme v příkazovém řádku tímto příkazem:
python -m pip install pulp highspy pandas openpyxl.
Poté se v Powershellu otevře složka, v které jsou uloženy program i excelová tabulka s daty k načtení a nakonec se spustí samotný program na vybrané vstupní data. Toto provedeme dvěma příkazy:      
(1) cd "C:\cesta\ke\složce"
(2) & "C:\cesta\k\python.exe" solve_rozvrh.py rozvrh_skola.xlsx

# Použití:
Program slouží jako nástroj pro sestavení školního rozvrhu na základě požadavků a omezení specifikovaných v excelovém souboru.
Stažený soubor rozvrh_skola.xlsx může člověk přepsat podle vlastních požadavků, nicméně kdyby si chtěl vytvořit svůj od základu, tak soubor je tvaru:
1. list: 'Ucitele' sloupce: Ucitel, Uvazek  
                   obsah: jména učitelů a počet hodin týdně podle jejich úvazku..
2. list: 'Kompetence' sloupce: Ucitel, Predmet
                        obsah: ke každému jménu učitele píšeme do vedlejší buňky právě jeden předmět
                               pokud jich je víc, musíme učitelovo jméno opsat víckrát pro každý jeho předmět odděleně.
3. list: 'Kurikulum' sloupce: Trida, Predmet, Hodiny
                       obsah: požadavky pro jednotlivé třídy,
                              kolik hodin kterého předmětu týdně má být vyučováno.
4. list: 'Ucebny' sloupce: Predmet, Ucebna
                    obsah: přiřazení předmětu do konkrétní učebny 
                          (např. fyzika-> laboratoř, matematika-> běžná učebna).
5. list: 'Sloty' sloupce: Slot, Priorita
                   obsah: časové sloty (např. "Po 1", "Út 2"…) 
                          a jejich preference (1 = preferovaný, 2 = méně vhodný, 3 = nedostupný).
6. list: 'Dostupnost' sloupce: Ucitel, Slot, Priorita
                        obsah: informace, zda je učitel v daný čas dostupný a s jakou prioritou.
Program vrátí jednak excelovou tabulku s rozvrhy pro jednotlivé třídy na oddělených listech, čili např. list pojemenovaný 2.A obsahuje rozvrh pro třídu 2.A. 
a jednak řešení do konzole, to vypadá např. takto:
<img width="1711" height="310" alt="image" src="https://github.com/user-attachments/assets/90d32e12-61ff-4247-b6c4-061d11ef2853" />

Uživatel může v programu v sekci model, proměnné, omezení a cíl ve funkci build_model požadavky na minimální počet dní s výukou a omezení na
maximální počet hodin za den (min_dni_s_vyukou: int = 5 & max_hodin_den: int = 6)
Při každém novém spuštění se soubor rozvrh_vysledek přepíše na aktualní výsledek, a objeví se ve stejné složce jako jsou solve_rozvrh a rozvrh_skola. Samotný soubor rozvrh_skola se ale nemění.
  
# Technický popis
Program je napsán v Pythonu a je strukturován do několika částí:

Načítání dat (funkce load_data)
Zajišťuje zpracování excelového souboru rozvrh_skola.xlsx. Z jednotlivých listů vytvoří odpovídající Python datové struktury (slovníky, seznamy, datové rámce). Díky tomu se vstupní data převedou do podoby vhodné pro optimalizační model.

Model (funkce build_model)

Vytváří matematický model rozvrhu jako úlohu lineárního programování, kde čveřici učitel–předmět–třída–slot–učebna přiřadí 1, nebo 0 v závislosti na tom, jestli daný učitel učí daný předmět v daný čas v dané učebně, nebo nikoliv. Při hledání řešení máme některá tvrdá omezení, jako např.

učitel nemůže učit dvě hodiny najednou,

učitel může učit jen předměty ze svých kompetencí,

dodržuje se úvazek učitele,

jednotlivé předměty musí být odučeny podle kurikula,

každá třída má v daném čase právě nejvýše hodinu,

respektuje se dostupnost učeben,

a také měkká omezení, která určují prioritizovaný výsledek:

Respektují se časové dostupnosti učitelů (sloty + priority).


Řešení (funkce solve_model)
Používá knihovny pulp a highspy pro nalezení optimálního rozvrhu. Minimalizuje penalizace za nevyhovující sloty a snaží se přiblížit preferovaným časovým rozvrhům. Pokud řešení neexistuje, program tuto skutečnost oznámí.
Tvorba rozvrhu samotného (funkce extract_schedule)
vybere příslušné kombinace učitel–předmět–třída–slot–učebna, tedy ty označené 1 a vytvoří čitelný rozvrh.
Výstup (hlavní spouštěč): 
Výsledný rozvrh se uloží do konzole, aby uživatel viděl přehled při běhu programu a do excelového souboru rozvrh_vysledek.xlsx, kde má každá třída svůj list se sestaveným rozvrhem.
