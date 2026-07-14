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
#ifndef TYPE_ITEM_HPP
#define TYPE_ITEM_HPP

#include "../compat/SfmlLegacy.hpp"
#include "../tinyxml/tinyxml.h"
#include "../myvector.hpp"

#include "TypeItemCategory.hpp"
#include "TypeSkill.hpp"

namespace Types
{
	class Item
	{
		public:
			Item(const TiXmlElement* const config);
			~Item();
			
			void DrawIcon(sf::RenderTarget* const render, const double& x, const double& y);
			bool Catch(const int& x, const int& y, const double& posX, const double& posY);
			
			unsigned int GetId() const;
			std::string GetName() const;
			std::string GetDescription() const;
			double GetManufacturingDuration() const;
			Types::ItemCategory* GetItemCategory() const;
			double GetBonus(const unsigned int id) const;
			Types::Skill* GetSkill() const;
			
			unsigned int GetRequisiteMaterialCount() const;
			unsigned int GetRequisiteMaterialId(const unsigned int& materialId) const;
			double GetRequisiteMaterialAmount(const unsigned int& materialId) const;
		
		private:
			unsigned int mId;
			std::string mName;
			std::string mDescription;
			double mManufacturingTime;
			Types::ItemCategory* mItemCategory;
			Types::Skill* mSkill;
			
			Vector <unsigned int> mTypesMaterialsNecessaires;
			Vector <double> mMaterialAmountsNecessaires;
			
			Vector <double> mBonus;
			
			legacy::Image* mImageItem;
			legacy::Sprite* mSpriteItem;
	};
}

#endif
