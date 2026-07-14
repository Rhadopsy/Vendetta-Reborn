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
#include "TypeMaterial.hpp"

#include "../xml.hpp"
#include "../data.hpp"

namespace Types
{
	Material::Material(TiXmlElement* const config) :
		mId(0),
		mName(""),
		mProductionRate(0),
		//mSkillId(0),
		mImage(new legacy::Image()),
		mSprite(new legacy::Sprite())
	{
		LoadXmlValue(config, "id", mId);
		LoadXmlValue(config, "Name", mName);
		
		std::string imageFile = "";
		LoadXmlValue(config, "ImageFile", imageFile);
		if (mImage->loadFromFile(R_IMAGES_RESSOURCES + imageFile))
		{
			mSprite->setImage(*mImage);
			mSprite->resize(Prgm::Materials::IconSize, Prgm::Materials::IconSize);
		}
		else
			mSprite->setImage(legacy::Image(1, 1, sf::Color::Transparent));
		
		LoadXmlValue(config, "ConversionSpeed", mProductionRate);
		//LoadXmlValue(config, "Skill", mSkillId);
		
		TiXmlElement* selement = config->FirstChildElement("Requisite");
		if (selement)
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
					mRequisiteMaterials.push_back(materialId);
					mRequisiteMaterialAmounts.push_back(amount);
				}
				
				selement = selement->NextSiblingElement("Material");
			}
		}
	}
	
	Material::~Material()
	{
		delete mImage;
		delete mSprite;
	}
	
	void Material::DrawIcon(sf::RenderTarget* const render, const double& x, const double& y)
	{
		if (mSprite)
		{
			mSprite->setPosition(x, y);
			render->draw(*mSprite);
		}
	}
	
	bool Material::Catch(const int& x, const int& y, const double& posX, const double& posY)
	{
		return mSprite
				&& posX <= x && x <= posX + mSprite->getSize().x
				&& posY <= y && y <= posY + mSprite->getSize().y;
	}
	
	unsigned int Material::GetId() const
	{
		return mId;
	}
	
	std::string Material::GetName() const
	{
		return mName;
	}
	
	double Material::GetProductionRate() const
	{
		return mProductionRate;
	}
	/*
	unsigned int Material::GetSkillId() const
	{
		return mSkillId;
	}
	*/
	unsigned int Material::GetRequisiteMaterialCount() const
	{
		return mRequisiteMaterials.size();
	}
	
	unsigned int Material::GetRequisiteMaterialId(const unsigned int& materialId) const
	{
		if (materialId < mRequisiteMaterials.size())
			return mRequisiteMaterials[materialId];
		else
			return 0;
	}
	
	double Material::GetRequisiteMaterialAmount(const unsigned int& materialId) const
	{
		if (materialId < mRequisiteMaterialAmounts.size())
		return mRequisiteMaterialAmounts[materialId];
		else
			return 0;
	}
}
