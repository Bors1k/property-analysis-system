from shipment import Shipment
# класс отдела


class Otdel:

    def __init__(self, name):
        self.name = name
        self.shipments = [Shipment]

    def addNewGood(self, shipment: Shipment):
        flag = True
        if(len(self.shipments)!=0):
            for TempShipment in self.shipments:
                if(shipment.getName() == TempShipment.getName()):
                    TempShipment.increaseCount()
                    flag = False
            if(flag):
                self.shipments.append(shipment)
        else:
            self.shipments.append(shipment)
