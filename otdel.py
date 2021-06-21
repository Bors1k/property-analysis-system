from shipment import Shipment
# класс отдела


class Otdel:

    def __init__(self, name):
        self.name = name
        self.shipments = []

    def addNewShipment(self,shipments, shipment: Shipment,count,srok):
        flag = True
        if(len(shipments)!=0):
            for TempShipment in shipments:
                if(shipment.getName() == TempShipment.getName()):
                    TempShipment.increaseCount(count,srok)
                    flag = False
            if(flag):
                shipment.increaseCount(count,srok)
                shipments.append(shipment)
        else:
            shipment.increaseCount(count,srok)
            shipments.append(shipment)

    def getName(self):
        return self.name

    def getShipments(self):
        return self.shipments
