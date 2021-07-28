#класс имущества
class Shipment:

    def __init__(self,name):
        self.name = name
        self.shipCount: int = 0
        self.expiredShipCount: int = 0


    def getName(self):
        return self.name
    
    def increaseCount(self,count,srok):
        if('в пределах' in srok):
            self.shipCount += count
        else:
            self.expiredShipCount += count
            self.shipCount += count