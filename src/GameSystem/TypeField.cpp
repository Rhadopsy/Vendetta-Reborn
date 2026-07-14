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
#include "TypeField.hpp"

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "../data.hpp"

namespace Types
{
	Field::Field(TiXmlElement* const config) :
		mProbability(0),
		mSpeedRatio(0),
		mBuildable(config->FirstChildElement("Constructible") != NULL),
		mAcceptRessource(config->FirstChildElement("Ressources") != NULL),
		mSprite(NULL)
	{
		int idSprite;
		LoadXmlValue(config, "Sprite", idSprite);
		mSprite = Configuration->GetSprite(idSprite);
		if (mSprite)
			mSprite->resize(Prgm::Fields::Width, Prgm::Fields::Height);
		
		LoadXmlValue(config, "FacteurDeplacement", mSpeedRatio);
		LoadXmlValue(config, "Probabilite", mProbability);
	}
	
	legacy::Sprite* Field::GetSprite() const
	{
		return mSprite;
	}
	
	bool Field::AcceptRessource() const
	{
		return mAcceptRessource;
	}
	
	double Field::GetSpeedRatio() const
	{
		return mSpeedRatio;
	}
	
	bool Field::Buildable() const
	{
		return mBuildable;
	}
	
	double Field::GetProbability() const
	{
		return mProbability;
	}
}
