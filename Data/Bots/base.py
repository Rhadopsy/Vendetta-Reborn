from Vendetta import *

class Character(Character_Base):
	def Harvest(self, identifier, amount):
		if (amount > 0):
			Ressource = self.GetRessource(identifier)
			if (ressource):
				self.GoObject(Ressource)
				self.Work(Ressource, amount)
	def Build(self, identifier):
		Building = self.SeatBuilding(identifier)
		if (Building):
			TypeBuilding = Building.GetType()
			
			materialCount = TypeBuilding.GetRequisiteMaterialCount()
			for i in range(materialCount):
				material = TypeBuilding.GetRequisiteMaterialId(i)
				amount = TypeBuilding.GetRequisiteMaterialAmount(i)
				self.Harvest(material, amount)
			
			self.GoObject(Building)
			self.Work(Building, TypeBuilding.GetMaxHP())
			self.Enter(Building)
	
	def onWhenever(self):
		pass
	def onSpawn(self):
		pass
	def onAttacked(self, Object):
		pass
	def onTargetSelected(self, Object):
		pass
	def onGoCoords(self, x, y):
		pass
	def onGoObject(self, Object):
		pass
