# Zapoctovy-program---tvroba-rozvrhu
Stručný popis:
Program načte požadavky na školní rozvrh z excelové tabulky, najde nejlépe vyhovující řešení (pokud nějaké existuje) a to vrátí jednak do konzole a jednak do excelového souboru.
# Dokumentace
Spuštění:
Ke spuštění programu je nejprve nutno nainstalovat python 3 a do něj potřebné knihovny. Ty instalujeme v příkazovém řádku tímto příkazem: python -m pip install pulp highspy pandas openpyxl. Poté se v Powershellu otevře složka, v které je program i excelová tabulka s daty k načtení a nakonec se spustí samotný program na vybrané vstupní data. Toto provedeme dvěma příkazy:      
(1) cd "C:\cesta\k\rozvrh_skola.xlsx"
(2) & "C:\cesta\k\python.exe" solve_rozvrh.py rozvrh_skola.xlsx
