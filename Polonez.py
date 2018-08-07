import win32com.client
import os
import pandas as pd


VISUM_PATH = os.path.join(os.getcwd(),"visum.ver")
ATTS = ["LENGTH",	"IMPEDANCE",	"T0",	"TCUR",	"V0",	"VCUR"]
ATT = "Length"
ATT_PUT = "Length"

TSYS = "SO"
MODE = "PuT"
DEP_TIME = "08:00"
KRYTERIUM = 3
MATRIX_NO = 102

BUDGET = 120


POI_CAT = 1
"""
Attributes for SPS PrT
Member                  Value
criteria_AddVal1        4  
criteria_AddVal2        5  
criteria_AddVal3        6 
criteria_Distance       3 
criteria_Impedance      2  
criteria_t0             0 
criteria_tCur           1 
PrTSearchCriterionT_End 7   
"""


"""
Attributes for SPS PuT
Arr  
ArrDay  
ArrTime  
Dep  
DepDay  
DepTime  
OrigZoneNo  
DestZoneNo  
FromStopAreaNo  
ToStopAreaNo  
FromNodeNo  
ToNodeNo  
Length  
Time 
"""
def SPS_PrT(_from, _to):
    # Szuka sciezki w sieci drogowej dla zadanego kryterium pomiedzy zadana para punktow, zwraca zadany atrybut
    RouteSearch = Visum.Analysis.RouteSearchPrT
    Route = Visum.CreateNetElements()
    RouteSearch.Clear()
    Route.Add(_from)
    Route.Add(_to)
    RouteSearch.Execute(Route, TSYS, KRYTERIUM)
    return RouteSearch.AttValue(ATT)

def SPS_PuT(_from, _to):
    # Szuka sciezki w sieci KZ dla zadanego kryterium pomiedzy zadana para punktow, zwraca zadany atrybut
    RouteSearch = Visum.Analysis.RouteSearchPuT
    Route = Visum.CreateNetElements()
    RouteSearch.Clear()
    Route.Add(_from)
    Route.Add(_to)
    RouteSearch.Execute(Route,MODE, DEP_TIME)
    return RouteSearch.AttValue(ATT_PUT)

def POI2NearestNode(Visum):
    Visum.Graphic.StopDrawing = True
    mm = Visum.Net.CreateMapMatcher()
    Iterator = Visum.Net.POICategories.ItemByKey(POI_CAT).POIs.Iterator
    while Iterator.Valid:
        POI = Iterator.Item
        nearest_node = mm.GetNearestNode(POI.AttValue("XCoord"), POI.AttValue("YCoord"), 500, False)
        if nearest_node.Success:
            POI.SetAttValue("Node", nearest_node.Node.AttValue("No"))
            POI.SetAttValue("Dist_PrT", nearest_node.Distance)
        Iterator.Next()
    Visum.Graphic.StopDrawing = False


def POI2NearestSPoint(Visum):
    Visum.Graphic.StopDrawing = True
    mm = Visum.Net.CreateMapMatcher()
    Iterator = Visum.Net.POICategories.ItemByKey(POI_CAT).POIs.Iterator
    while Iterator.Valid:
        POI = Iterator.Item
        nearest_node = mm.GetNearestNode(POI.AttValue("XCoord"), POI.AttValue("YCoord"), 1000, True)
        if nearest_node.Success:
            POI.SetAttValue("SPoint", nearest_node.Node.AttValue("CONCATENATE:STOPAREAS\NO"))
            POI.SetAttValue("Dist_PuT", nearest_node.Distance)
        Iterator.Next()
    Visum.Graphic.StopDrawing = False

def MainLoopStages(Visum):
    Visum.Graphic.StopDrawing = True  # nie rysuj (przyspieszenie)
    # inicjalizacja bazy dancyh (do csv)
    df_s1 = pd.DataFrame(columns=["Z_Rejon", "POI", "Czas_PrT", "Czas_PuT"])
    df_s2 = pd.DataFrame(columns=["POI", "Do_Rejon", "Czas_PrT", "Czas_PuT"])

    Zones = Visum.Net.Zones.GetMultiAttValues("No")  # rejony do iteracji
    # dane o POI
    POIs = Visum.Net.POICategories.ItemByKey(1).POIs.GetMultipleAttributes(["No", "Node", "SPoint", "Dist_PrT", "Dist_PuT"])

    # glowna petla
    for OZone in Zones:
        Z = Visum.Net.Zones.ItemByKey(OZone[1]) # para rejonow Z
        for POI in POIs:
            #S1
            Przez_PrT = Visum.Net.Nodes.ItemByKey(POI[1]) #Punkt w sieci dla POI
            CzPrT = SPS_PrT(Z, Przez_PrT) + POI[3] # Oblicz czas PrT (2x dojscie do POI)

            if POI[2] is not None:
                Przez_PuT = Visum.Net.StopAreas.ItemByKey(POI[2]) # Przystanek dla POI (jesli jest)
                CzPuT = SPS_PuT(Z, Przez_PuT) + POI[4]   # Oblicz czas PuT (2x dojscie do POI)
            else:
                CzPuT = 999999
            print("From {} to {} in {} PrT, {} PuT".format(OZone[1], int(POI[0]), CzPrT, CzPuT))
            df_s1.loc[df_s1.shape[0] + 1] = [OZone[1], int(POI[0]), CzPrT, CzPuT] # zapisz rekord w bazie danych

            #s2
            CzPrT = SPS_PrT(Przez_PrT, Z) + POI[3]  # Oblicz czas PrT (2x dojscie do POI)

            if POI[2] is not None:
                Przez_PuT = Visum.Net.StopAreas.ItemByKey(POI[2])  # Przystanek dla POI (jesli jest)
                CzPuT = SPS_PuT(Przez_PuT,Z) + POI[4]  # Oblicz czas PuT (2x dojscie do POI)
            else:
                CzPuT = 999999
            print("From {} to {} in {} PrT, {} PuT".format(OZone[1], int(POI[0]), CzPrT, CzPuT))
            df_s2.loc[df_s1.shape[0] + 1] = [int(POI[0]), OZone[1], CzPrT, CzPuT]  # zapisz rekord w bazie danych

    df_s1.to_csv("POI_1Stage.csv")  # zapisz baze do pliku
    df_s2.to_csv("POI_2Stage.csv")  # zapisz baze do pliku
    Visum.Graphic.StopDrawing = False


def MainLoop(Visum):
    Visum.Graphic.StopDrawing = True  # nie rysuj (przyspieszenie)
    # inicjalizacja bazy dancyh (do csv)
    df = pd.DataFrame(columns=["Z_Rejon", "Do_Rejon", "L_Podrozy", "Przez_POI", "Czas_PrT", "Czas_PuT"])

    Zones = Visum.Net.Zones.GetMultiAttValues("No")  # rejony do iteracji
    # dane o POI
    POIs = Visum.Net.POICategories.ItemByKey(1).POIs.GetMultipleAttributes(["No", "Node", "SPoint", "Dist_PrT", "Dist_PuT"])

    # glowna petla
    for OZone in Zones:
        for DZone in Zones:
            if OZone != DZone:
                nTrips = Visum.Net.Matrices.ItemByKey(MATRIX_NO).GetValue(OZone[1],DZone[1]) # liczba podrozy (z macierzy)
                Z = Visum.Net.Zones.ItemByKey(OZone[1]) # para rejonow Z
                Do = Visum.Net.Zones.ItemByKey(DZone[1]) # i do
                for POI in POIs:
                    Przez_PrT = Visum.Net.Nodes.ItemByKey(POI[1]) #Punkt w sieci dla POI
                    CzPrT = SPS_PrT(Z, Przez_PrT) + SPS_PrT(Przez_PrT, Do) + 2 * POI[3] # Oblicz czas PrT (2x dojscie do POI)

                    if POI[2] is not None:
                        Przez_PuT = Visum.Net.StopAreas.ItemByKey(POI[2]) # Przystanek dla POI (jesli jest)
                        CzPuT = SPS_PuT(Z, Przez_PuT) + SPS_PuT(Przez_PuT, Do)+ 2*POI[4]   # Oblicz czas PuT (2x dojscie do POI)
                    else:
                        CzPuT = 999999
                    print("From {} to {} Via {} in {} PrT, {} PuT".format(OZone[1], DZone[1], int(POI[0]), CzPrT, CzPuT))
                    df.loc[df.shape[0] + 1] = [OZone[1], DZone[1], nTrips, int(POI[0]), CzPrT, CzPuT] # zapisz rekord w bazie danych
    df.to_csv("POIs.csv")  # zapisz baze do pliku
    Visum.Graphic.StopDrawing = False

def read_csvs():
    df1 = pd.read_csv('POI_1Stage.csv')
    df2 = pd.read_csv('POI_2Stage.csv')
    result = pd.merge(df1, df2, how='outer', on=['POI'])
    result = result[['Z_Rejon',"POI", "Do_Rejon", u'Czas_PrT_x', u'Czas_PuT_x', u'Czas_PrT_y', u'Czas_PuT_y']]
    result.to_csv("Joined.csv")  # zapisz baze do pliku
    # Zones = df1.Z_Rejon.unique()
    # POIs = df1.Do_POI.unique()
    # df = pd.DataFrame(columns=["Z_Rejon", "Do_Rejon", "L_Podrozy", "Przez_POI", "Czas_PrT", "Czas_PuT"])
    # for OZone in Zones:
    #     for DZone in Zones:
    #         for POI in POIs:
    #             S1 = df1[(df1.Z_Rejon==OZone) & (df1.Do_POI==POI)]
    #             S2 = df2[(df2.Z_POI == POI) & (df2.Do_Rejon == DZone)]
    #             df.loc[df.shape[0] + 1] = [OZone, DZone, 0, POI, S1.Czas_PrT+S2.Czas_PrT,
    #                                        S1.Czas_PuT + S2.Czas_PuT]  # zapisz rekord w bazie danych
    #             print("From {} to {} Via {} in {} PrT, {} PuT".format(OZone, DZone, int(POI), 0,0))

    # df.to_csv("Joined.csv")  # zapisz baze do pliku

def make_matrix():
    import matplotlib.pyplot as plt
    df = pd.read_csv('Joined.csv')
    (df.Czas_PrT_x+df.Czas_PrT_y).hist()

    plt.show()
    print(df.shape)
    print(df.head())

    matrix = df[(df.Czas_PrT_x + df.Czas_PrT_y) < 100000].groupby(by=['Z_Rejon', 'Do_Rejon'])  # group with two indexes
    OD = matrix.size()  # make trip matrix
    pd.options.display.float_format = '{:,.0f}'.format
    OD = OD.unstack().fillna(0)  # fill na with zeros, unstack the column matrix to classic view
    print(OD.head())

if __name__ == "__main__":

    make_matrix()
    quit()

    Visum = win32com.client.Dispatch("Visum.Visum")  # uruchom Visum
    Visum.LoadVersion(VISUM_PATH)  # zaladuj plik

    POI2NearestNode(Visum) # przypisz wezly sieci do POI
    POI2NearestSPoint(Visum) # przypisz przystanki do POI

    MainLoopStages(Visum) # glowny algorytm
    read_csvs()

    #Visum.SaveVersion(VISUM_PATH)








