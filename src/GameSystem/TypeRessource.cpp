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
#include "TypeRessource.hpp"

#include "../xml.hpp"
#include "../Configuration.hpp"

namespace Types
{
	Ressource::Ressource(TiXmlElement* const config) :
		mId(0),
		mSprite(NULL),
		mWidth(0),
		mHeight(0),
		mName(""),
		mActionList(""),
		mMaterial(NULL),
		mHarvestRate(0)
	{
		LoadXmlValue(config, "id", mId);
		
		mSprite = Configuration->GetSprite(config, "Sprite");
		
		LoadXmlValue(config, "Width", mWidth);
		LoadXmlValue(config, "Height", mHeight);
		if (mSprite)
			mSprite->resize(mWidth, mHeight);
		
		LoadXmlValue(config, "Name", mName);
		LoadXmlValue(config, "Action", mActionList);
		
		unsigned int materialId = 0;
		LoadXmlValue(config, "Material", materialId);
		mMaterial = Configuration->GetTypeMaterial(materialId);
		
		LoadXmlValue(config, "HarvestSpeed", mHarvestRate);
	}
	
	Ressource::~Ressource()
	{
		delete mSprite;
	}
	
	void Ressource::draw(sf::RenderTarget* const render, const double& x, const double& y)
	{
		if (mSprite)
		{
			mSprite->setPosition(x, y);
			render->draw(*mSprite);
		}
	}
	
	unsigned int Ressource::GetId() const
	{
		return mId;
	}
	
	unsigned int Ressource::getWidth() const
	{
		return mWidth;
	}
	
	unsigned int Ressource::getHeight() const
	{
		return mHeight;
	}
	
	unsigned int Ressource::GetResId() const
	{
		return mMaterial ? mMaterial->GetId() : 0;
	}
	
	std::string Ressource::GetAction() const
	{
		return mActionList;
	}
	
	Types::Material* Ressource::GetTypeMaterial() const
	{
		return mMaterial;
	}
	
	double Ressource::GetHarvestRate() const
	{
		return mHarvestRate;
	}
}
