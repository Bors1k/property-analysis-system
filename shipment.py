#класс имущества
class Shipment:

    def __init__(self,name):
        self.name = name
        self.shipCount = 0
        self.expiredShipCount = 0

    def getName(self):
        return self.name
    
    def increaseCount(self,count):
        self.shipCount += count