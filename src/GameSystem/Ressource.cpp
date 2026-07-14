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
#include "Ressource.hpp"

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "Character.hpp"

namespace GameSystem
{
	Ressource::Ressource() : GameSystem::Object(),
		mType(NULL)
	{
		mClass = GameSystem::Classes::Ressource;
	}
	
	void Ressource::Load()
	{
		if (mXml)
		{
			LoadXmlValue(mXml, "X", mX);
			LoadXmlValue(mXml, "Y", mY);
			
			unsigned int type;
			if (LoadXmlValue(mXml, "Type", type))
				this->SetType(Configuration->GetTypeRessource(type));
		}
	}
	
	TiXmlElement* Ressource::Save() const
	{
		TiXmlElement* saveNode = new TiXmlElement("Ressource");
		saveNode->SetAttribute("id", mId);
		saveNode->SetDoubleAttribute("X", mX);
		saveNode->SetDoubleAttribute("Y", mY);
		saveNode->SetAttribute("Type", mType->GetId());
		
		return saveNode;
	}
	
	double Ressource::Harvest(GameSystem::Object* const collecteur, const double& time)
	{
		if (collecteur->GetClass() == GameSystem::Classes::Character)
		{
			GameSystem::Character* character = static_cast<GameSystem::Character*>(collecteur);
			return character->AddMaterial(mType->GetResId(), mType->GetHarvestRate() * time);
		}
		else
			return 0;
	}
	
	std::string Ressource::getText() const
	{
		return mType->GetAction();
	}
	
	void Ressource::draw(sf::RenderTarget* const render)
	{
		if (mType)
			mType->draw(render, mX, mY);
	}
	
	void Ressource::SetType(Types::Ressource* const type)
	{
		mType = type;
		
		if (mType)
		{
			mWidth = mType->getWidth();
			mHeight = mType->getHeight();
		}
	}
	
	Types::Ressource* Ressource::GetTypeRessource() const
	{
		return mType;
	}
	
	bool Ressource::Buildable(const double& x, const double& y, const unsigned int& width, const unsigned int& height) const
	{
		return mX <= x + width && mY <= y + height && x <= mX + mWidth && y <= mY + mHeight;
	}
}
