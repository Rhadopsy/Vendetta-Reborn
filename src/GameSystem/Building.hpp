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
#ifndef BUILDING_HPP
#define BUILDING_HPP

#include "../xml.hpp"
#include "../myvector.hpp"

#include "Object.hpp"
#include "Item.hpp"
#include "TypeBuilding.hpp"
#include "TypeMaterial.hpp"

namespace GameSystem
{
	// WorkList
	struct WorkElement
	{
		bool material;
		unsigned int id;
		double amount;
		
		GameSystem::WorkElement* nextElement;
	};
	
	class Building : public GameSystem::Object
	{
		public:
			Building(Types::Building* const type = NULL);
			virtual ~Building();
			
			virtual void Load();
			virtual TiXmlElement* Save() const;
			
			std::string getText() const;
			void draw(sf::RenderTarget* const render);
			
			double Work(GameSystem::Object* const worker, const double& time);
			
			void AddMaterialToWorlList(const unsigned int& materialId, const unsigned int& amount);
			void AddItemToWorkList(const unsigned int& itemId);
			void RemoveFromWorkList(const unsigned int& elementId);
			
			bool AddItemToStock(GameSystem::Item* const objet);
			GameSystem::Item* RemoveFromStock(const unsigned int& itemId);
			
			double GetHP() const;
			Types::Building* GetType() const;
			bool IsBuilt() const;
			
			double GetRequisiteMaterialAmount(const unsigned int& materialId) const;
			unsigned int GetStockItemCount(const unsigned int& typeItemId) const;
			double GetItemManufacturingProgress(const unsigned int& typeItemId) const;
			GameSystem::WorkElement* GetWorkList() const;
			
			bool Buildable(const double& x, const double& y, const unsigned int& width, const unsigned int& height) const;
		
		private:
			Types::Building* mType;
			
			bool mBuilt;
			
			Vector <double> mRequisiteMaterial;
			Vector <double> mItemMaking;
			Vector <unsigned int> mItemStock;
			
			struct GameSystem::WorkElement* mWorkList;
	};
}

#endif
