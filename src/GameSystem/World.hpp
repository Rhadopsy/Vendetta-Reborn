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
#ifndef WORLD_HPP
#define WORLD_HPP

#include "../tinyxml/tinyxml.h"
#include "../myvector.hpp"

#include "Ressource.hpp"
#include "Character.hpp"
#include "Building.hpp"
#include "Item.hpp"

namespace GameSystem
{
	class World
	{
		public:
			World();
			virtual ~World();
			
			void Load(const TiXmlElement* const element = NULL);
			TiXmlElement* Save() const;
			
			void SetPlayer(GameSystem::Character** const personnage);
			
			void draw(sf::RenderTarget* const render);
			void Process(const double& time);
			
			double getWidth() const;
			double getHeight() const;
			
			GameSystem::Object* GetObjectPoint(const double& x, const double& y) const;
			int GetIdField(const double& x, const double& y) const;
			
			void AjouteObject(GameSystem::Object* const object, const TiXmlElement* const element = NULL);
			GameSystem::Object* GetObject(const unsigned int& objectId) const;
			void RemoveObject(const unsigned int& objectId);
			
			GameSystem::Ressource* FindRessource(const GameSystem::Object* const origin, Types::Ressource* const type) const;
			unsigned int AssignNation(GameSystem::Character* character);
			bool RessourceFits(GameSystem::Ressource* ressource, const double& x, const double& y) const;
			bool Buildable(const double& x, const double& y, Types::Building* const building) const;
			GameSystem::Building* SeatBuilding(const double& x, const double& y, Types::Building* const type);
			GameSystem::Item* CreateItem(Types::Item* const type = NULL);
		
		private:
			double mWidth;
			double mHeight;
			unsigned int mWidthEnFields;
			unsigned int mHeightEnFields;
			
			unsigned int* mGrilleFields;
			legacy::Image* mImage;
			legacy::Sprite* mSprite;
			
			Vector <GameSystem::Object*> mObjects;
			Vector <GameSystem::Ressource*> mRessources;
			Vector <GameSystem::Character*> mCharacters;
			GameSystem::Character** mCharacterJoueur;
			Vector <GameSystem::Building*> mBuildings;
			Vector <GameSystem::Item*> mItems;
	};
}

#endif
