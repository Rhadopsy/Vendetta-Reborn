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
#ifndef TYPE_BUILDING_HPP
#define TYPE_BUILDING_HPP

#include "../compat/SfmlLegacy.hpp"
#include "../tinyxml/tinyxml.h"
#include "../myvector.hpp"

#include "TypeItem.hpp"

namespace Types
{
	class Building
	{
		public:
			Building(TiXmlElement* const config = NULL);
			~Building();
			
			void DrawStructure(sf::RenderTarget* const render, const double& x, const double& y, const double& HP);
			void SetPositionButton(const double& x, const double& y);
			void DrawButton(sf::RenderTarget* const render);
			
			bool Catch(const int& x, const int& y) const;
			
			unsigned int GetId() const;
			std::string GetName() const;
			std::string GetDescription() const;
			double GetMaxHP() const;
			unsigned int GetStructureWidth() const;
			unsigned int GetStructureHeight() const;
			
			unsigned int GetRequisiteMaterialCount() const;
			unsigned int GetRequisiteMaterialId(const unsigned int& materialId) const;
			double GetRequisiteMaterialAmount(const unsigned int& materialId) const;
			
			unsigned int GetAvailableMaterialCount() const;
			unsigned int GetMaterialIdFromList(const unsigned int& materialId) const;
			
			unsigned int GetAvailableItemCount() const;
			Types::Item* GetItemFromList(const unsigned int& idItem);
			
		private:
			unsigned int mId;
			std::string mName;
			std::string mDescription;
			double mMaxHP;
			unsigned int mStepCount;
			
			Vector <unsigned int> mRequisiteMaterials;
			Vector <double> mRequisiteMaterialAmounts;
			Vector <unsigned int> mAvailableMaterials;
			Vector <unsigned int> mAvailableItems;
			
			unsigned int mSpriteWidth;
			unsigned int mSpriteHeight;
			unsigned int mStructureWidth;
			unsigned int mStructureHeight;
			legacy::Image* mStructureImage;
			legacy::Sprite* mStructureSprite;
			legacy::Sprite* mFoundationsSprite;
			
			double mButtonX;
			double mButtonY;
			legacy::Image* mButtonImage;
			legacy::Sprite* mButtonSprite;
	};
}

#endif
