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
#ifndef TYPE_MATERIAL_HPP
#define TYPE_MATERIAL_HPP

#include "../compat/SfmlLegacy.hpp"
#include "../tinyxml/tinyxml.h"
#include "../myvector.hpp"

namespace Types
{
	class Material
	{
		public:
			Material(TiXmlElement* const config);
			~Material();
			
			void DrawIcon(sf::RenderTarget* const render, const double& x, const double& y);
			bool Catch(const int& x, const int& y, const double& posX, const double& posY);
			
			unsigned int GetId() const;
			std::string GetName() const;
			double GetProductionRate() const;
			//unsigned int GetSkillId() const;
			
			unsigned int GetRequisiteMaterialCount() const;
			unsigned int GetRequisiteMaterialId(const unsigned int& id) const;
			double GetRequisiteMaterialAmount(const unsigned int& id) const;
		
		private:
			unsigned int mId;
			std::string mName;
			double mProductionRate;
			//unsigned int mSkillId;
			
			Vector <unsigned int> mRequisiteMaterials;
			Vector <double> mRequisiteMaterialAmounts;
			
			legacy::Image* mImage;
			legacy::Sprite* mSprite;
	};
}

#endif
