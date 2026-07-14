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
#include "TypeItem.hpp"

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "../data.hpp"

namespace Types
{
	Item::Item(const TiXmlElement* const config) :
		mId(0),
		mName(""),
		mDescription(""),
		mManufacturingTime(0),
		mItemCategory(NULL),
		mSkill(NULL),
		mImageItem(new legacy::Image()),
		mSpriteItem(new legacy::Sprite())
	{
		LoadXmlValue(config, "id", mId);
		LoadXmlValue(config, "Name", mName);
		LoadXmlValue(config, "Description", mDescription);
		
		LoadXmlValue(config, "ManufacturingTime", mManufacturingTime);
		
		unsigned int idCategory = 0;
		LoadXmlValue(config, "Category", idCategory);
		mItemCategory = Configuration->GetItemCategory(idCategory);
		
		unsigned int skillId;
		LoadXmlValue(config, "Skill", skillId);
		mSkill = Configuration->GetSkill(skillId);
		
		std::string imageFile = "";
		LoadXmlValue(config, "ImageFile", imageFile);
		if (mImageItem->loadFromFile(R_IMAGES_OBJETS + imageFile))
		{
			mSpriteItem->setImage(*mImageItem);
			mSpriteItem->resize(Prgm::Items::IconSize, Prgm::Items::IconSize);
		}
		else
			mSpriteItem->setImage(legacy::Image(1, 1, sf::Color::Transparent));
		
		const TiXmlElement* selement;
		if ( (selement = config->FirstChildElement("Requisite")) )
		{
			selement = selement->FirstChildElement("Material");
			while (selement)
			{
				int materialId = 0;
				double amount = 0;
				
				LoadXmlValue(selement, "Type", materialId);
				LoadXmlValue(selement, "Amount", amount);
				
				if (materialId > 0 && amount > 0)
				{
					mTypesMaterialsNecessaires.push_back(materialId);
					mMaterialAmountsNecessaires.push_back(amount);
				}
				
				selement = selement->NextSiblingElement("RequisiteMaterial");
			}
		}
		/*
		if ( (selement = config->FirstChildElement("Bonus")) )
		{
			const TiXmlElement* sselement = selement->FirstChildElement("Caracteristique");
			while (sselement)
			{
				std::string nomSkill = "";
				unsigned int idSkill = 0;
				double valeurBonus = 0;
				if (LoadXmlValue(sselement, "Type", nomSkill) &&
					(idSkill = Configuration->GetSkillId(nomSkill)) &&
					LoadXmlValue(sselement, "Value", valeurBonus) )
				{
					mBonus[idSkill] = valeurBonus;
				}
				
				sselement = sselement->NextSiblingElement("Caracteristique");
			}
		}*/
	}
	
	Item::~Item()
	{
		delete mImageItem;
		delete mSpriteItem;
	}
	
	void Item::DrawIcon(sf::RenderTarget* const render, const double& x, const double& y)
	{
		if (mSpriteItem)
		{
			mSpriteItem->setPosition(x, y);
			render->draw(*mSpriteItem);
		}
	}
	
	bool Item::Catch(const int& x, const int& y, const double& posX, const double& posY)
	{
		return mSpriteItem
				&& posX <= x && x <= posX + mSpriteItem->getSize().x
				&& posY <= y && y <= posY + mSpriteItem->getSize().y;
	}
	
	unsigned int Item::GetId() const
	{
		return mId;
	}
	
	std::string Item::GetName() const
	{
		return mName;
	}
	
	std::string Item::GetDescription() const
	{
		return mDescription;
	}
	
	double Item::GetManufacturingDuration() const
	{
		return mManufacturingTime;
	}
	
	Types::ItemCategory* Item::GetItemCategory() const
	{
		return mItemCategory;
	}
	
	double Item::GetBonus(const unsigned int id) const
	{
		return mBonus[id];
	}
	
	Types::Skill*  Item::GetSkill() const
	{
		return mSkill;
	}
	
	unsigned int Item::GetRequisiteMaterialCount() const
	{
		return mTypesMaterialsNecessaires.size();
	}
	
	unsigned int Item::GetRequisiteMaterialId(const unsigned int& materialId) const
	{
		return materialId < mTypesMaterialsNecessaires.size() ? mTypesMaterialsNecessaires[materialId] : 0;
	}
	
	double Item::GetRequisiteMaterialAmount(const unsigned int& materialId) const
	{
		return materialId < mMaterialAmountsNecessaires.size() ? mMaterialAmountsNecessaires[materialId] : 0;
	}
}
