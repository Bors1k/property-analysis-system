from shipment import Shipment
# класс отдела


class Otdel:

    def __init__(self, name):
        self.name = name
        self.shipments = [Shipment]

    def addNewGood(self, shipment: Shipment):
        if(self.checkForNewOne(shipment.getName())):
            self.shipments.append()

    def checkForNewOne(self, shipmentName):
        for shipment in self.shipments:
            if(shipmentName == shipment.getName()):
                return False
