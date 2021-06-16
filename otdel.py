from shipment import Shipment
# класс отдела


class Otdel:

    def __init__(self, name):
        self.name = name
        self.shipments = []

    def addNewShipment(self, shipment: Shipment,count):
        flag = True
        if(len(self.shipments)!=0):
            for TempShipment in self.shipments:
                if(shipment.getName() == TempShipment.getName()):
                    TempShipment.increaseCount(count)
                    flag = False
            if(flag):
                self.shipments.append(shipment)
        else:
            self.shipments.append(shipment)

    def getName(self):
        return self.name

    def getShipments(self):
        return self.shipments
