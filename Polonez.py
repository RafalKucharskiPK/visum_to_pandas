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


if __name__ == "__main__":

    Visum = win32com.client.Dispatch("Visum.Visum")  # uruchom Visum
    Visum.LoadVersion(VISUM_PATH)  # zaladuj plik

    POI2NearestNode(Visum) # przypisz wezly sieci do POI
    POI2NearestSPoint(Visum) # przypisz przystanki do POI

    MainLoop(Visum) # glowny algorytm
    #Visum.SaveVersion(VISUM_PATH)








