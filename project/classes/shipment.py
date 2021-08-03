#класс имущества
class Shipment:

    def __init__(self,name):
        self.name = name
        self.shipCount: int = 0
        self.expiredShipCount: int = 0
        self.expiredInNextYearCount: int = 0


    def getName(self):
        return self.name
    
    def increaseCount(self,count,srok,expiredNextYearFlag: bool):
        if('в пределах' in srok):
            if(expiredNextYearFlag == True):
                self.expiredInNextYearCount+=count
            self.shipCount += count
        else:
            if(expiredNextYearFlag == True):
                self.expiredInNextYearCount+=count
            self.expiredShipCount += count
            self.shipCount += count