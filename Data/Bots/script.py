from Vendetta import *
import random

class Character(Character):
	target = 0
	def onWhenever(self):
		if (self.target and not self.isAttacking(self.target)):
			if (Range(self, self.target) < 100):
				self.ResetActions();
				self.Attack(self.target)
			elif (not self.isGoing(self.target)):
				self.ResetActions();
				self.GoObject(self.target)
	def onSpawn(self):
		self.Build(random.randint(1,1))
	def onAttacked(self, Character):
		if (not self.isActive()):
			self.target = Character
	def onTargetSelected(self, Object):
		self.ResetActions();
		if (Object.GetClass() == Types.Character):
			self.target = Object;
		else:
			self.GoObject(Object)
			if (Object.GetClass() == Types.Building):
				self.Enter(Object)
			self.Work(Object)
	def onGoCoords(self, x, y):
		print "Je vais en " + str(int(x)) + ":" + str(int(y))
		self.SetSay("Je vais en " + str(int(x)) + ":" + str(int(y)))
		self.Say("J'arrive en " + str(int(x)) + ":" + str(int(y)))
		self.target = 0
	def onGoObject(self, Object):
		pass
