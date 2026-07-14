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
#include "TypeRace.hpp"

#include "../xml.hpp"
#include "../Configuration.hpp"

namespace Types
{
	Race::Race(TiXmlElement* const config) :
		mImageId(0),
		mWidth(24),
		mHeight(32),
		mSpriteWidth(24),
		mSpriteHeight(32),
		mSpeed(7),
		mDescription(""),
		mName_Male_Single(""),
		mName_Male_Several(""),
		mName_Female_Single(""),
		mName_Female_Several(""),
		mEquipementIllustration(0)
	{
		LoadXmlValue(config, "Image", mImageId);
		
		LoadXmlValue(config, "SpriteWidth", mSpriteWidth);
		LoadXmlValue(config, "SpriteHeight", mSpriteHeight);
		LoadXmlValue(config, "CharacterWidth", mWidth);
		LoadXmlValue(config, "CharacterHeight", mHeight);
		LoadXmlValue(config, "Speed", mSpeed);
		
		LoadXmlValue(config, "Description", mDescription);
		TiXmlElement* selement = config->FirstChildElement("Names");
		if (selement)
		{
			LoadXmlValue(selement, "MaleSingle", mName_Male_Single);
			LoadXmlValue(selement, "MaleSeveral", mName_Male_Several);
			LoadXmlValue(selement, "FemaleSingle", mName_Female_Single);
			LoadXmlValue(selement, "FemaleSeveral", mName_Female_Several);
		}
		
		selement = config->FirstChildElement("Equipment");
		if (selement)
		{
			LoadXmlValue(selement, "Illustration", mEquipementIllustration);
			
			int categorieItem;
			int x;
			int y;
			
			selement = selement->FirstChildElement("Spot");
			while (selement)
			{
				LoadXmlValue(selement, "ItemCategory", categorieItem);
				LoadXmlValue(selement, "x", x);
				LoadXmlValue(selement, "y", y);
				
				if (categorieItem > 0 && x > 0 && y > 0)
				{
					mEquipmentSpotCategories.push_back(categorieItem);
					mEquipmentSpotAbscissas.push_back(x);
					mEquipmentSpotOrdinates.push_back(y);
				}
				
				selement = selement->NextSiblingElement("Spot");
			}
		}
	}
	
	legacy::Image* Race::getImage()
	{
		return Configuration->getImage(mImageId);
	}
	
	unsigned int Race::getWidth() const
	{
		return mWidth;
	}
	
	unsigned int Race::getHeight() const
	{
		return mHeight;
	}
	
	unsigned int Race::GetSpriteWidth() const
	{
		return mSpriteWidth;
	}
	
	unsigned int Race::GetSpriteHeight() const
	{
		return mSpriteHeight;
	}
	
	double Race::GetSpeed() const
	{
		return mSpeed;
	}
	
	unsigned int Race::GetSpriteEquipementIllustration() const
	{
		return mEquipementIllustration;
	}
	
	Vector <unsigned int> Race::GetEquipmentSpotCategories() const
	{
		return mEquipmentSpotCategories;
	}
	
	Vector <unsigned int> Race::GetEquipmentSpotAbscissas() const
	{
		return mEquipmentSpotAbscissas;
	}
	
	Vector <unsigned int> Race::GetEquipmentSpotOrdinates() const
	{
		return mEquipmentSpotOrdinates;
	}
	
}
