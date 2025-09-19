# Zapoctovy-program---tvroba-rozvrhu
Stručný popis:
Program načte požadavky na školní rozvrh z excelové tabulky, najde nejlépe vyhovující řešení (pokud nějaké existuje) a to vrátí jednak do konzole a jednak do excelového souboru.
# Dokumentace
# Instalace & spuštění:
Ke spuštění programu je nejprve nutno nainstalovat python 3 a do něj potřebné knihovny. Ty instalujeme v příkazovém řádku tímto příkazem: python -m pip install pulp highspy pandas openpyxl. Poté se v Powershellu otevře složka, v které jsou uloženy program i excelová tabulka s daty k načtení a nakonec se spustí samotný program na vybrané vstupní data. Toto provedeme dvěma příkazy:      
(1) cd "C:\cesta\k\rozvrh_skola.xlsx"
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

Rozvrh pro 2.A:
Hodina                        1                     2                        3                         4                        5                      6
Den
Po      Tělesná výchova (Hrubá)    Matematika (Novák)       Matematika (Novák)       Český jazyk (Novák)                        -     Angličtina (Beneš)
Út           Matematika (Novák)   Vlastivěda (Dvořák)                        -        Matematika (Novák)      Český jazyk (Kovář)        Prvouka (Novák)
St                            -   Vlastivěda (Dvořák)     Český jazyk (Dvořák)  Tělesná výchova (Veselý)      Český jazyk (Kovář)    Prvouka (Svobodová)
Čt          Matematika (Dvořák)  Český jazyk (Dvořák)       Vlastivěda (Kovář)                         -      Angličtina (Veselý)  Přírodověda (Novotná)
Pá                            -  Angličtina (Novotná)  Český jazyk (Svobodová)     Přírodověda (Novotná)  Český jazyk (Svobodová)    Angličtina (Veselý)

Uživatel může v programu v sekci model, proměnné, omezení a cíl hned na začátku měnit požadavky na minimální počet dní s výukou a omezení na
maximální počet hodin za den.
def build_model(data: dict,
                pokuta_nedoplnene: int = 5000,     
                pokuta_priorita2: int = 1000,       
                pokuta_nerovnomernost: int = 1000,  
                min_dni_s_vyukou: int = 5,         
                rozdil_hodin_den: int = 2,         
                max_hodin_den: int = 6)
Při každém novém spuštění se soubor rozvrh_vysledek přepíše na aktualní výsledek, a objeví se ve stejné složce jako jsou solve_rozvrh a rozvrh_skola.
  



