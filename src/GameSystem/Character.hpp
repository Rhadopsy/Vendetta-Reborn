//*	Game (RPG, management and strategy)
//*	Copyright (C) 2011  Santos Quentin
//*	
//*	This program is free software: you can redistribute it and/or modify
//*	it under the terms of the GNU General Public License as published by
//*	the Free Software Foundation, either version 3 of the License, or
//*	(at your option) any later version.
//*	
//*	This program is distributed in the hope that it will be useful,
//*	but WITHOUT ANY WARRANTY; without even the implied warranty of
//*	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//*	GNU General Public License for more details.
//*	
//*	You should have received a copy of the GNU General Public License
//*	along with this program.  If not, see <http:#www.gnu.org/licenses/>.
//*	
//*	
#ifndef CHARACTER_HPP
#define CHARACTER_HPP

#include "../myvector.hpp"

#include "Object.hpp"
#include "TypeRace.hpp"
#include "Ressource.hpp"
#include "Item.hpp"
#include "Building.hpp"

#include "../Interface/Text.hpp"

namespace GameSystem
{
	namespace TypesActions
	{
		enum Names
		{
			Wait,
			Say,
			GoCoords,
			GoObject,
			Enter,
			Work,
			Attack
		};
	}
	
	struct ActionElement
	{
		GameSystem::TypesActions::Names type;
		void** params;
		
		GameSystem::ActionElement* nextAction;
	};
	
	class Character : public GameSystem::Object
	{
		public:
			Character();
			virtual ~Character();
			
			void Load();
			TiXmlElement* Save() const;
			
			// Déroulement du jeu
			bool Catch(const double& x, const double& y) const;
			void ReceiveString(const sf::String& txt);
			void Process(const double& time);
			void draw(sf::RenderTarget* const render);
			
			GameSystem::Ressource* GetRessource(const unsigned int& id);
			
			double GetXSprite() const;
			double GetYSprite() const;
			unsigned int GetNation() const;
			
			// Actions
			void AddAction(GameSystem::ActionElement* action);
			void NextAction();
			void ResetActions();
			
			void SetSay(const char* txt);
			void Wait(double time);
			void Say(const char* txt);
			void GoCoords(const double& x, const double& y);
			void GoObject(GameSystem::Object* const target);
			void Enter(GameSystem::Building* const building);
			void Work(GameSystem::Object* const target, const double& amount = 0);
			void Attack(GameSystem::Object* const target);
			GameSystem::Building* SeatBuilding(const unsigned int& id);
			
			// Tests d'activité
			bool isActive() const;
			bool isGoing(GameSystem::Object* const target) const;
			bool isAttacking(GameSystem::Object* const target) const;
			GameSystem::Object* GetTarget() const;
			
			// Déclencheurs
			//virtual void onWhenever() {};
			//virtual void onSpawn() {};
			//virtual void onAttacked(GameSystem::Object*) {};
			virtual void onTargetSelected(GameSystem::Object*) {};
			virtual void onGoCoords(double, double) {};
			virtual void onGoObject(GameSystem::Object*) {};
			
			Types::Race* GetRace() const;
			std::string GetName() const;
			GameSystem::Building* GetCurrentBuilding() const;
			std::string GetTextTypeBuilding(const Types::Building* const typeBuilding) const;
			std::string GetTextObject(GameSystem::Object* const object) const;
			
			double GetMaterialStockAmount(const unsigned int& materialId) const;
			double AddMaterial(const unsigned int& materialId, const double& amount);
			double RemoveMaterial(const unsigned int& materialId, const double& amount);
			
			Vector <GameSystem::Item*>* GetInventory();
			void AddItemToStock(GameSystem::Item* const item);
			
			Vector <GameSystem::Item*>* GetItemsEquipes();
		
		private:
			bool mVisible;
			Types::Race* mRace;
			unsigned int mNation;
			unsigned int mSpriteWidth;
			unsigned int mSpriteHeight;
			double mSpeed;
			std::string mName;
			unsigned int mDirection;
			double mStep;
			
			GameSystem::ActionElement* mActionList;
			GameSystem::Building* mCurrentBuilding;
			
			InterfaceModule::Text* mText;
			
			Vector <double> mMaterialAmounts;
			
			Vector <double> mSkillLevels;
			Vector <double> mAlteredSkills;
			Vector <GameSystem::Item*> mInventory;
			Vector <GameSystem::Item*> mEquipment;
			
			legacy::Image* mImage;
			legacy::Sprite* mSprite;
	};
}

#endif
