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
    RouteSearch = Visum.Analysis.RouteSearchPrT
    Route = Visum.CreateNetElements()
    RouteSearch.Clear()
    Route.Add(_from)
    Route.Add(_to)
    RouteSearch.Execute(Route, TSYS, KRYTERIUM)
    return RouteSearch.AttValue(ATT)

def SPS_PuT(_from, _to):
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
    df = pd.DataFrame(columns=["Z_Rejon", "Do_Rejon", "L_Podrozy", "Przez_POI", "Czas_PrT", "Czas_PuT"])

    Zones = Visum.Net.Zones.GetMultiAttValues("No")
    POIs = Visum.Net.POICategories.ItemByKey(1).POIs.GetMultipleAttributes(["No", "Node", "SPoint", "Dist_PrT", "Dist_PuT"])

    for OZone in Zones:
        for DZone in Zones:
            if OZone != DZone:
                nTrips = Visum.Net.Matrices.ItemByKey(MATRIX_NO).GetValue(OZone[1],DZone[1])
                for POI in POIs:
                    Z = Visum.Net.Zones.ItemByKey(OZone[1])
                    Przez_PrT = Visum.Net.Nodes.ItemByKey(POI[1])
                    Do = Visum.Net.Zones.ItemByKey(DZone[1])
                    CzPrT = SPS_PrT(Z, Przez_PrT) + SPS_PrT(Przez_PrT, Do) + 2 * POI[3]

                    if POI[2] is not None:
                        Przez_PuT = Visum.Net.StopAreas.ItemByKey(POI[2])
                        CzPuT = SPS_PuT(Z, Przez_PuT) + SPS_PuT(Przez_PuT, Do)+ 2*POI[4]
                    else:
                        CzPuT = 999999
                    print("From {} to {} Via {} in {} PrT, {} PuT".format(OZone[1], DZone[1], int(POI[0]), CzPrT, CzPuT))
                    df.loc[df.shape[0] + 1] = [OZone[1], DZone[1], nTrips, int(POI[0]), CzPrT, CzPuT]
    df.to_csv("POIs.csv")


if __name__ == "__main__":

    Visum = win32com.client.Dispatch("Visum.Visum")
    Visum.LoadVersion(VISUM_PATH)

    POI2NearestNode(Visum)
    POI2NearestSPoint(Visum)

    MainLoop(Visum)
    #Visum.SaveVersion(VISUM_PATH)








