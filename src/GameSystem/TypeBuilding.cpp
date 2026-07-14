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
#include "TypeBuilding.hpp"

#include <iostream>

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "../data.hpp"
#include "../functions.hpp"

namespace Types
{
	Building::Building(TiXmlElement* const config) :
		mId(0),
		mName(""),
		mDescription(""),
		mMaxHP(0),
		mStepCount(0),
		mSpriteWidth(0),
		mSpriteHeight(0),
		mStructureWidth(0),
		mStructureHeight(0),
		mStructureImage(new legacy::Image()),
		mStructureSprite(new legacy::Sprite()),
		mFoundationsSprite(Configuration->GetSprite(Sprites::FoundationsSprite)),
		mButtonX(0),
		mButtonY(0),
		mButtonImage(new legacy::Image()),
		mButtonSprite(new legacy::Sprite())
	{
		LoadXmlValue(config, "id", mId);
		LoadXmlValue(config, "Name", mName);
		LoadXmlValue(config, "Description", mDescription);
		LoadXmlValue(config, "MaxHP", mMaxHP);
		LoadXmlValue(config, "SpriteWidth", mSpriteWidth);
		LoadXmlValue(config, "SpriteHeight", mSpriteHeight);
		LoadXmlValue(config, "StructureWidth", mStructureWidth);
		LoadXmlValue(config, "StructureHeight", mStructureHeight);
		LoadXmlValue(config, "StepCount", mStepCount);
		
		std::string adresse;
		LoadXmlValue(config, "StructureImage", adresse);
		if (mStructureImage->loadFromFile(R_IMAGES_BATIMENTS_STRUCTURES + adresse))
			mStructureSprite->setImage(*mStructureImage);
		else
		{
			mStructureSprite->setImage(legacy::Image(1, 1, sf::Color::Transparent));
			std::cerr << "Erreur lors du chargement de l'image d'un building\n";
		}
		
		if (mFoundationsSprite)
			mFoundationsSprite->resize(mStructureWidth, mStructureHeight);
		else
			std::cerr << "L'image de debut de construction n'a pas put object chargee\n";
		
		LoadXmlValue(config, "ButtonImage", adresse);
		if (mButtonImage->loadFromFile(R_IMAGES_BATIMENTS_BOUTONS + adresse))
			mButtonSprite->setImage(*mButtonImage);
		else
		{
			mStructureSprite->setImage(legacy::Image(1, 1, sf::Color::Transparent));
			std::cerr << "Erreur lors du chargement de l'icone d'un building\n";
		}
		
		TiXmlElement* selement;
		if ( (selement = config->FirstChildElement("Requisite")) )
		{
			int materialId;
			int amount;
			selement = selement->FirstChildElement("Material");
			while (selement)
			{
				materialId = 0;
				amount = 0;
				
				LoadXmlValue(selement, "Type", materialId);
				LoadXmlValue(selement, "Quantite", amount);
				
				if (materialId > 0 && amount > 0)
				{
					mRequisiteMaterials.push_back(materialId);
					mRequisiteMaterialAmounts.push_back(amount);
				}
				
				selement = selement->NextSiblingElement("Material");
			}
		}
		
		selement = config->FirstChildElement("Available");
		if (selement)
		{
			unsigned int id;
			TiXmlElement* sselement;
			
			sselement = selement->FirstChildElement("Material");
			while (sselement)
			{
				if (LoadXmlValue(sselement, "Type", id))
					mAvailableMaterials.push_back(id);
				
				sselement = sselement->NextSiblingElement("Material");
			}
			
			
			sselement = selement->FirstChildElement("Item");
			while (sselement)
			{
				if (LoadXmlValue(sselement, "Type", id))
					mAvailableItems.push_back(id);
				
				sselement = sselement->NextSiblingElement("Item");
			}
		}
	}
	
	Building::~Building()
	{
		delete mStructureImage;
		delete mStructureSprite;
		
		delete mButtonImage;
		delete mButtonSprite;
	}
	
	void Building::DrawStructure(sf::RenderTarget* const render, const double& x, const double& y, const double& HP)
	{
		unsigned int step = static_cast<unsigned int>( (mStepCount - 1) * HP / mMaxHP );
		
		if (step == 0)
		{
			if (mFoundationsSprite)
			{
				mFoundationsSprite->setPosition(x, y);
				render->draw(*mFoundationsSprite);
			}
		}
		else
		{
			step--;
			
			mStructureSprite->setSubRect( sf::IntRect( 0, step * mSpriteHeight, mSpriteWidth, (step + 1) * mSpriteHeight) );
			mStructureSprite->resize(mStructureWidth, mStructureHeight);
			mStructureSprite->setPosition(x, y);
			render->draw(*mStructureSprite);
		}
		
		DrawGauge(render, x, y + mStructureHeight, mStructureWidth, 5, HP / mMaxHP);
	}
	
	void Building::SetPositionButton(const double& x, const double& y)
	{
		mButtonX = x;
		mButtonY = y;
		mButtonSprite->setPosition(mButtonX, mButtonY);
	}
	
	void Building::DrawButton(sf::RenderTarget* const render)
	{
		mButtonSprite->setSubRect( sf::IntRect(0, Prgm::TypesBuildings::SpriteButtonSize, Prgm::TypesBuildings::SpriteButtonSize, 2 * Prgm::TypesBuildings::SpriteButtonSize) );
		mButtonSprite->resize(Prgm::TypesBuildings::ButtonSize, Prgm::TypesBuildings::ButtonSize);
		render->draw(*mButtonSprite);
	}
	
	bool Building::Catch(const int& x, const int& y) const
	{
		return (mButtonX <= x && x <= mButtonX + Prgm::TypesBuildings::ButtonSize && mButtonY <= y && y <= mButtonY + Prgm::TypesBuildings::ButtonSize);
	}
	
	unsigned int Building::GetId() const
	{
		return mId;
	}
	
	std::string Building::GetName() const
	{
		return mName;
	}
	
	std::string Building::GetDescription() const
	{
		return mDescription;
	}
	
	double Building::GetMaxHP() const
	{
		return mMaxHP;
	}
	
	unsigned int Building::GetStructureWidth() const
	{
		return mStructureWidth;
	}
	
	unsigned int Building::GetStructureHeight() const
	{
		return mStructureHeight;
	}
	
	unsigned int Building::GetRequisiteMaterialCount() const
	{
		return mRequisiteMaterials.size();
	}
	
	unsigned int Building::GetRequisiteMaterialId(const unsigned int& materialId) const
	{
		return materialId < mRequisiteMaterials.size() ? mRequisiteMaterials[materialId] : 0;
	}
	
	double Building::GetRequisiteMaterialAmount(const unsigned int& materialId) const
	{
		return materialId < mRequisiteMaterials.size() ? mRequisiteMaterialAmounts[materialId] : 0;
	}
	
	unsigned int Building::GetAvailableMaterialCount() const
	{
		return mAvailableMaterials.size();
	}
	
	unsigned int Building::GetMaterialIdFromList(const unsigned int& materialId) const
	{
		return materialId < mAvailableMaterials.size() ? mAvailableMaterials[materialId] : 0;
	}
	
	unsigned int Building::GetAvailableItemCount() const
	{
		return mAvailableItems.size();
	}
	
	Types::Item* Building::GetItemFromList(const unsigned int& idItem)
	{
		return idItem < mAvailableItems.size() ? Configuration->GetTypeItem(mAvailableItems[idItem]) : NULL;
	}
}
