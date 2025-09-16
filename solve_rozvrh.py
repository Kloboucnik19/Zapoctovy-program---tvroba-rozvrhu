# Načítá se excel soubor s listy:
# - Ucitele(Ucitel, Uvazek)
# - Kompetence(Ucitel, Predmet)
# - Kurikulum(Trida, Predmet, Hodiny)
# - Ucebny(Predmet, Ucebna)
# - Sloty(Slot, Priorita)- volitelná *globální* (školní) priorita slotu
# - Dostupnost(Ucitel, Slot, Priorita)-  *individuální* priorita učitele v daném slotu
#
# Priority: 1 = OK/preference, 2 = lze s penalizací, 3 = zakázáno.
# Pokud je slot definován i globálně i u učitele, používám max(ucitelova_priorita, global_priorita), tj. když škola slot zakáže (3), je to 3 pro všechny.
#
# Spuštění:
# (pro mě osobně): otevřít PowerShell a spustit:
#     cd "C:\Users\Jan Pavlů\OneDrive - Univerzita Karlova\Desktop\MATFYZ\Programování\Zápočtový program"
#     & "C:\Users\Jan Pavlů\AppData\Local\Programs\Python\Python313\python.exe" solve_rozvrh.py rozvrh_skola.xlsx
# obecně:
#     cd "C:\cesta\k\rozvrh_skola_even.xlsx"
#     & "C:\cesta\k\python.exe" solve_rozvrh.py rozvrh_skola_even.xlsx
#
# Potřeba nainstalovat:
#     python -m pip install pulp highspy pandas openpyxl
#
# Výstup:
# - stav řešení do konzole,
# - plus Excel `rozvrh_vysledek.xlsx` s tabulkou rozvrhu.

from __future__ import annotations
import sys
from typing import Dict, List, Set, Tuple
import pandas as pd
import pulp as pl


# Načtení dat z excelu
def load_data(xlsx_path: str):
  '''
  Čtení listů z excelu, na vstupu je cesta k souboru, na výstupu slovník s daty tohoto tvaru:
  - ucitele: List[str]
  - uvazky: Dict[str, int]
  - dostupnost: Dict[(teacher, slot), int]
  - kompetence: Dict[teacher, Set[subject]]
  - tridy: List[str]
  - predmety: List[str]
  - pozadavky: Dict[(class, subject), int]
  - predmet_ucebna: Dict[subject, room]
  - ucebny: List[str]
  - casove_sloty: List[str]
  - dny: List[str]   (např. ["Po","Út",...])
  - hodiny: List[str] (hodiny jako "1","2",...)
  '''
  df_u = pd.read_excel(xlsx_path, sheet_name="Ucitele")
  df_k = pd.read_excel(xlsx_path, sheet_name="Kompetence")
  df_c = pd.read_excel(xlsx_path, sheet_name="Kurikulum")
  df_r = pd.read_excel(xlsx_path, sheet_name="Ucebny")
  df_s = pd.read_excel(xlsx_path, sheet_name="Sloty")
  df_d = pd.read_excel(xlsx_path, sheet_name="Dostupnost")

  # všechny časové sloty/hodiny(tím je myšleno po 1, po 2, ut 1 atd.) - uloženo jako list stringů
  casove_sloty: List[str] = df_s["Slot"].astype(str).tolist()
  # dále obdobně načtu i zbývající listy z excelu do příslušných struktur
  # Ucitelé + úvazky
  ucitele: List[str] = df_u["Ucitel"].astype(str).tolist()
  uvazky: Dict[str,int] = {r["Ucitel"]: int(r["Uvazek"]) for _, r in df_u.iterrows()}

  #globální priorita slotu (pokud nevyplněno, tak default = 1)
  globalni_priorita: Dict[str, int] = {r["Slot"]: int(r["Priorita"]) for _, r in df_s.iterrows()}

  # dostupnost učitelů (pokud nevyplněno, tak default = 1)
  priorita_ucitele: Dict[Tuple[str,str], int] = {}
  for _, r in df_d.iterrows():
    priorita_ucitele[(str(r["Ucitel"]), str(r["Slot"]))] = int(r["Priorita"])

  # Kombinovaná dostupnost: max(teacher, global)
  dostupnost: Dict[Tuple[str,str], int] = {}
  for t in ucitele:
    for l in casove_sloty:
      tp = priorita_ucitele.get((t,l), 1)
      gp = globalni_priorita.get(l, 1)
      dostupnost[(t,l)] = max(tp, gp) # bereme přísnější prioritu (3 = zakázáno)

  # Kompetence učitelů
  kompetence: Dict[str, Set[str]] = {t:set() for t in ucitele}
  for _, r in df_k.iterrows():
    kompetence.setdefault(str(r["Ucitel"]), set()).add(str(r["Predmet"])) #set abychom zamezili potenciálním duplicitám, samozřejmě jeden učitel může učit více předmětů

  # Kurikulum: kolik hodin týdně třída potřebuje z předmětu
  tridy: List[str] = sorted(df_c["Trida"].astype(str).unique())
  predmety: List[str] = sorted(df_c["Predmet"].astype(str).unique())
  pozadavky: Dict[Tuple[str,str], int] = {}
  for _, r in df_c.iterrows():
    pozadavky[(str(r["Trida"]), str(r["Predmet"]))] = int(r["Hodiny"])

  # Předmětu je přiřazena učebna
  predmet_ucebna: Dict[str,str] = {str(r["Predmet"]): str(r["Ucebna"]) for _, r in df_r.iterrows()}
  ucebny: List[str] = sorted(set(predmet_ucebna.values()))

  # Rozdělení slotů na dny a hodiny
  dny: List[str] = sorted(list(set([l.split()[0] for l in casove_sloty])))
  hodiny: List[str] = sorted(list(set([l.split()[1] for l in casove_sloty])), key=int) # rozdělení na dny a hodiny se hodí později pro formátování

  # sanity-checks
  _missing_slots = [l for l in casove_sloty if l not in df_d["Slot"].unique()]
  if _missing_slots:
    print(f"[Info] Sloty bez explicitní dostupnosti učitelů: {', '.join(map(str,_missing_slots))} → default 1")

  return {
    "ucitele": ucitele,
    "uvazky": uvazky,
    "dostupnost": dostupnost,
    "kompetence": kompetence,
    "tridy": tridy,
    "predmety": predmety,
    "pozadavky": pozadavky,
    "predmet_ucebna": predmet_ucebna,
    "ucebny": ucebny,
    "casove_sloty": casove_sloty,
    "dny": dny,
    "hodiny": hodiny
  }

# ------------------------------
# Model: proměnné, omezení, cíl
# ------------------------------
def build_model(data: dict,
                pokuta_nedoplnene: int = 5000,     # int udává "závažnost přestupku", doplnění všech hodin je hlavní priorita
                pokuta_priorita2: int = 1000,       
                pokuta_nerovnomernost: int = 1000,  
                min_dni_s_vyukou: int = 5,         
                rozdil_hodin_den: int = 2,         
                max_hodin_den: int = 6):            
  """
  MILP model pro rozvrh:
    - Rozhodovací proměnné: binárky vyuka[ucitel,trida,predmet,slot], které říkají, jestli  danýý učitel učí danou třídu
      daný předmět ve daném slotu, pokud ano 1, pokud ne 0.
    - nedoplnene_hodiny[trida,predmet] = kolik hodin daného předmětu dané třídy se nepodařilo zahrnout do vystupu
    - denni_hodiny[trida][den] = kolik hodin má třída v den.
    - vyuka_den[trida][den] = 1, pokud má třída v den aspoň jednu hodinu (jinak 0).
    - min_dni_s_vyukou = minimální počet dnů, kdy má třída výuku.
    - rozdil_hodin_den = povolený rozdíl mezi "nejplnějším" a "nejprázdnějším" dnem (zavádím pro rovnoměrnost rozvrhu).
  """

  # Rozbalení dat z připraveného slovíku(který jsme extrahovali z excelu)
  # ucitele, tridy, predmety, casove_sloty, dny jsou mnoziny. 
  # ucebny je seznam, dostupnost, kompetence, pozadavky, predmet_ucebna, uvazky jsou slovníky
  ucitele = data["ucitele"]      
  tridy = data["tridy"]           
  predmety = data["predmety"]     
  casove_sloty = data["casove_sloty"]  # množina časových slotů (např. "Po 1", "Po 2", ...)
  dny = data["dny"]               
  ucebny = data["ucebny"]         
  dostupnost = data["dostupnost"] # dostupnost učitele (1=ok, 2=spíš ne, 3=zakázáno)
  kompetence = data["kompetence"] #key je ucitel, value jsou předměty v jeho kompetenci
  pozadavky = data["pozadavky"]   # kolik hodin každá třída potřebuje pro daný předmět
  predmet_ucebna = data["predmet_ucebna"]  # ke každému předmětu je určená učebna
  uvazky = data["uvazky"]         

  # Definice našeho problému, snažím se minimalizovat penalizace
  model = pl.LpProblem("RozvrhSimp", pl.LpMinimize)

###################################
#Proměnné (definice a vysvětlení): 
###################################
  # vyuka[ucitel,trida,predmet,slot] = 1, pokud učitel učí třídu předmět ve slotu
  vyuka = {}
  for ucitel in ucitele:
    povolene_predmety = kompetence.get(ucitel, set())  # předměty, které učitel smí učit
    if not povolene_predmety:
      continue
    for trida in tridy:
      for predmet in povolene_predmety:
        for slot in casove_sloty:
          # pokud dostupnost[(ucitel,slot)] == 3, zakázano, tedy vynechám 
          if dostupnost.get((ucitel, slot), 1) != 3:
            vyuka[(ucitel, trida, predmet, slot)] = pl.LpVariable(f"vyuka_{ucitel}_{trida}_{predmet}_{slot}", cat=pl.LpBinary)

  # nedoplnene_hodiny[trida,predmet] = celé číslo >=0, kolik hodin předmětu ve třídě se nepodařilo splnit
  nedoplnene_hodiny = {
    (trida, predmet): pl.LpVariable(f"nedoplnene_{trida}_{predmet}", lowBound=0, cat=pl.LpInteger)
    for (trida, predmet) in pozadavky.keys()
  }

  # denni_hodiny[trida][den] = kolik hodin má třída v den, abych mohl ověřit rovnoměrnost
  denni_hodiny = pl.LpVariable.dicts("denni_hodiny", (tridy, dny), lowBound=0, cat=pl.LpInteger)

  # vyuka_den[trida][den] = 1, pokud třída má aspoň jednu hodinu v den, na základce asi nechceme žádný volný den, už jen kvůli rodičům hehe
  vyuka_den = pl.LpVariable.dicts("vyuka_den", (tridy, dny), cat=pl.LpBinary)

  # pomocné proměnné pro rovnoměrnost: maximum a minimum počtu hodin v jednotlivých dnech
  denni_max = pl.LpVariable.dicts("denni_max", tridy, lowBound=0, cat=pl.LpInteger)
  denni_min = pl.LpVariable.dicts("denni_min", tridy, lowBound=0, cat=pl.LpInteger)

  ##################
  # Omezení(tvrdá)
  ##################   

  # 1 # Naplnění kurikula: pro každý (trida,predmet) musí součet naplánovaných hodin + nedoplnene_hodiny == požadavek
  for trida, predmet in pozadavky.keys():
    suma_hodin = 0
    for ucitel in ucitele:
        for slot in casove_sloty:
            suma_hodin += vyuka.get((ucitel, trida, predmet, slot), 0)
    model += suma_hodin + nedoplnene_hodiny[(trida, predmet)] == pozadavky[(trida, predmet)], f"kurikulum_{trida}_{predmet}"
    #bacha += značí přidej do modelu, ne přičti doslova jako by byl model float
  # 2 # Každá třída má v daném slotu max 1 předmět
  for trida in tridy:  
    for slot in casove_sloty:
      suma = 0
      for ucitel in ucitele:
        for predmet in predmety:
          suma += vyuka.get((ucitel, trida, predmet, slot), 0)
      model += suma <= 1, f"trida_slot_{trida}_{slot}" 

  # 3 # Učitel má v jednom slotu max 1 hodinu
  for ucitel in ucitele:
    for slot in casove_sloty:
      suma = 0
      for trida in tridy:
        for predmet in predmety:
          suma += vyuka.get((ucitel, trida, predmet, slot), 0)
      model += suma <= 1, f"ucitel_slot_{ucitel}_{slot}"

  # 4 # Každá učebna má v jednom slotu max 1 předmět (rozhoduje predmet_ucebna)
  for ucebna in ucebny:
    for slot in casove_sloty:
      suma = 0
      for ucitel in ucitele:
        for trida in tridy:
          for predmet in predmety:
            if predmet_ucebna.get(predmet) == ucebna:
              suma += vyuka.get((ucitel, trida, predmet, slot), 0)
      model += suma <= 1, f"ucebna_{ucebna}_{slot}"

  # 5 # Úvazek učitele: celkem hodin ≤ uvazek, aby se vešel do uvazku
  for ucitel in ucitele:
    suma = 0
    for trida in tridy:
      for predmet in predmety:
        for slot in casove_sloty:
          suma += vyuka.get((ucitel, trida, predmet, slot), 0)
    model += suma <= uvazky[ucitel], f"uvazek_{ucitel}"

  # 6 # Denní hodiny a propojení s vyuka_den
  BIGM = sum(pozadavky.values()) + 5  # dost velké číslo pro "implication trick"
  for trida in tridy:
    for den in dny:
      # denní hodiny = součet vyuka, kde slot patří do dne
      suma = 0
      for ucitel in ucitele:
        for predmet in predmety:
          for slot in casove_sloty:
            if slot.split()[0] == den:
              suma += vyuka.get((ucitel, trida, predmet, slot), 0)
      model += denni_hodiny[trida][den] == suma, f"denni_{trida}_{den}"

      # pokud denni_hodiny > 0 -> vyuka_den=1 (a naopak)
      model += denni_hodiny[trida][den] <= BIGM * vyuka_den[trida][den], f"link_vyuka_den_up_{trida}_{den}"
      model += denni_hodiny[trida][den] >= vyuka_den[trida][den], f"link_vyuka_den_down_{trida}_{den}"

  # 7 # Každá třída musí mít výuku aspoň v min_dni_s_vyukou dnech, možno změnit vbuilt_model ale nechavam na 5 dnech
  for trida in tridy:
    suma = 0
    for den in dny:
      suma += vyuka_den[trida][den]
    model += suma >= min_dni_s_vyukou, f"min_dni_{trida}"
  
  # 8 # Omezení na mx počet hodin denně
  for trida in tridy:
    for den in dny:
        model += denni_hodiny[trida][den] <= max_hodin_den, f"max_hodin_den_{trida}_{den}"


  # 9 # Rovnoměrnost: rozdíl mezi nejplnějším a nejprázdnějším dnem ≤ rozdil_hodin_den
  for trida in tridy:
    for den in dny:
      model += denni_max[trida] >= denni_hodiny[trida][den], f"max_link_{trida}_{den}"
      model += denni_min[trida] <= denni_hodiny[trida][den], f"min_link_{trida}_{den}"
    model += denni_max[trida] - denni_min[trida] <= rozdil_hodin_den, f"rovnomernost_{trida}"

  ####################
  # Objektivka(měkká)
  #####################
  #přidělím penalizace a vyberu nejméně špatné řešení z těch co prošli tvrdá omezení

  # 1 # Penalizace za chybějící hodiny
  cast_nedoplnene = pokuta_nedoplnene * pl.lpSum(nedoplnene_hodiny.values())

  # 2 # Penalizace za výuku v horších časech (dostupnost == 2)
  suma_priorita2 = []
  for key in vyuka:
      ucitel, trida, predmet, slot = key
      if dostupnost.get((ucitel, slot), 1) == 2:
          suma_priorita2.append(vyuka[key])
  cast_priorita2 = pokuta_priorita2 * pl.lpSum(suma_priorita2)

  # 3 # Penalizace za nerovnoměrné rozložení hodin během týdne
  suma_nerovnomernost = []
  for trida in tridy:
      rozdil = denni_max[trida] - denni_min[trida]
      suma_nerovnomernost.append(rozdil)
  cast_nerovnomernost = pokuta_nerovnomernost * pl.lpSum(suma_nerovnomernost)

  # 4 # Složení všech částí do jedné cílové funkce
  cilova_funkce = (
      cast_nedoplnene
      + cast_priorita2
      + cast_nerovnomernost
  )

  # 5 # Vložení cílové funkce do modelu (solver ji bude minimalizovat)
  model += cilova_funkce

  # Vrací model a důležitý proměnný
  return model, vyuka, nedoplnene_hodiny



##########################
# Funkce pro řešení modelu
##########################
# postupně zkoušíme jednotlivé solvery, 
# od nejrychlejšího, ale někdy méně dostupného k těch spolehlivějším ale pomalejším...

def solve_model(model: pl.LpProblem, prefer_highs: bool = True):
  solver = None
  status = None
  if prefer_highs:
    try:
      # zeprvé zkusíme použít rychlý solver HiGHS (pokud je k dispozici)
      solver = pl.HiGHS_CMD(msg=False)
      status = model.solve(solver)
      return status
    except Exception:
      pass
  try:
    # pokud není možno použít highs, použijeme CBC solver (default v PuLP)
    solver = pl.PULP_CBC_CMD(msg=False)
    status = model.solve(solver)
    return status
  except Exception:
    # pokud nefungoval ani jeden ze solveru výše, tak dávám solve() bez explicitního solveru
    status = model.solve()
    return status
  
############################################
# Funkce pro extrakci rozvrhu z proměnných
############################################
#tady vytvořím čitelný rozvrh pomocí proměnných, které dostanou ze solveru hodnoty 0 nebo 1, vyberu samozřejmě ty s jedničkou, jeilkož ty se konají
def extract_schedule(vyuka: Dict[Tuple[str,str,str,str], pl.LpVariable], data: dict) -> pd.DataFrame:
  rows = []
  for (ucitel, trida, predmet, slot), var in vyuka.items():    
    if var.varValue and var.varValue > 0.5: #ta první podmínka je kvůli tomu kdyby to byl None, druhou kvuli floatům dávám radši >0.5 než ==1 
      # najdeme správnou učebnu pro předmět
      ucebna = data["predmet_ucebna"].get(predmet)
      try:
        den, hodina = slot.split() #poze zapis slotu rozdělim na den a hodinu
      except ValueError:
        # fallback: když by slot nebyl dobře zapsaný
        den, hodina = slot, "?"
      rows.append({
        "Den": den,
        "Hodina": hodina,
        "Trida": trida,
        "Predmet": predmet,
        "Ucitel": ucitel,
        "Ucebna": ucebna,
        "Slot": slot
      })
  # vytvoření DataFrame z vybraných řádků, rows je seznam slovníků popisující jednotlive hodiny
  df = pd.DataFrame(rows)

  if not df.empty:
    #seřazení podle dnů a hodin
    day_order = {den: i for i, den in enumerate(data["dny"])}
    df["_d"] = df["Den"].map(day_order).fillna(999).astype(int) #fillna hází neznáme dny na konec
    with pd.option_context('mode.chained_assignment', None):
      try:
        # hodiny převedeme na čísla kvůli správnému řazení
        df["_p"] = df["Hodina"].astype(int)
      except Exception: #pokud třeba hodina bude "?" nebo něco jiného
        df["_p"] = 0
    # seřazení rozvrhu podle dne, hodiny a třídy
    df = df.sort_values(["_d", "_p", "Trida"]).drop(columns=["_d", "_p"]) #pomocné sloupce smažu
  return df

############################
# Hlavní spouštěč programu
############################
def main(argv: List[str]):
  if len(argv) < 2:
    print("Použití: python solve_rozvrh.py <cesta_k_excelu>")
    return 2

  # cesta k excelovému souboru s daty, tu dostanu od uživatele jakožto argument při spouštění
  xlsx_path = argv[1]
  data = load_data(xlsx_path)

  model, vyuka, nedoplnene_hodiny = build_model(data) #stavba modelu

  #zavolám solver
  status = solve_model(model)
  print("Status:", pl.LpStatus[status])

  if pl.LpStatus[status] == "Infeasible":
    print("Model je nerealizovatelný.")
    print("Možné příčiny: moc velké požadavky v kurikulu, málo slotů, nebo učitelé nemají kompetence.")
    return 1

  # spočítám, kolik hodin se do rozvrhu nevešlo
  total_nedoplnene = int(sum((v.varValue or 0) for v in nedoplnene_hodiny.values()))
  print(f"Nedoplněné hodiny: {total_nedoplnene}")

  # ty nasledně vypišu
  if total_nedoplnene > 0:
    print("Seznam nedoplněných hodin:")
    for (trida, predmet), var in nedoplnene_hodiny.items():
      missing = int(var.varValue or 0)
      if missing > 0:
        print(f"  - {trida} {predmet}: {missing} hod.")

  # převedeme výsledek do DataFrame
  df = extract_schedule(vyuka, data)

  if df.empty:
    print("Nepodařilo se najít rozvrh. Zkuste snížit požadavky, navýšit úvazky, nebo uvolnit sloty.")
  else:    
    df["Obsah"] = df["Predmet"] + " (" + df["Ucitel"] + ")"#tvořím buňky do klasickeho rozvrhu
    df["Hodina"] = df["Hodina"].astype(int)
    order = ["Po", "Út", "St", "Čt", "Pá"]
    df["Den"] = pd.Categorical(df["Den"], categories=order, ordered=True)#všechno pěkně seřadit a pak tisknout

    print("\n=== Ukázka rozvrhů ===")
    with pd.ExcelWriter("rozvrh_vysledek.xlsx") as writer:
      for trida in df["Trida"].unique():
        print(f"\nRozvrh pro {trida}:")
        df_trida = df[df["Trida"] == trida].copy()     
        pivot = df_trida.pivot(index="Den", columns="Hodina", values="Obsah")  # pivot tabulka: dny = řádky, hodiny = sloupce     
        pivot = pivot.fillna("-")#prázdný místa = pomlčka
        print(pivot.to_string())
        
        # uložím na samostatný sheet v Excelu
        pivot.to_excel(writer, sheet_name=trida)

    print("\nVýsledek uložen do: rozvrh_vysledek.xlsx")

  return 0

if __name__ == "__main__":
  sys.exit(main(sys.argv))
