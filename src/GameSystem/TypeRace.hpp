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
#ifndef TYPE_RACE_HPP
#define TYPE_RACE_HPP

#include "../compat/SfmlLegacy.hpp"
#include <string>
#include "../tinyxml/tinyxml.h"
#include "../myvector.hpp"

namespace Types
{
	class Race
	{
		public:
			Race(TiXmlElement* const config);
			
			legacy::Image* getImage();
			unsigned int getWidth() const;
			unsigned int getHeight() const;
			unsigned int GetSpriteWidth() const;
			unsigned int GetSpriteHeight() const;
			double GetSpeed() const;
			
			unsigned int GetSpriteEquipementIllustration() const;
			Vector <unsigned int> GetEquipmentSpotCategories() const;
			Vector <unsigned int> GetEquipmentSpotAbscissas() const;
			Vector <unsigned int> GetEquipmentSpotOrdinates() const;
		
		private:
			unsigned int mImageId;
			unsigned int mWidth;
			unsigned int mHeight;
			unsigned int mSpriteWidth;
			unsigned int mSpriteHeight;
			double mSpeed;
			
			std::string mDescription;
			std::string mName_Male_Single;
			std::string mName_Male_Several;
			std::string mName_Female_Single;
			std::string mName_Female_Several;
			
			unsigned int mEquipementIllustration;
			Vector <unsigned int> mEquipmentSpotCategories;
			Vector <unsigned int> mEquipmentSpotAbscissas;
			Vector <unsigned int> mEquipmentSpotOrdinates;
	};
}

#endif
