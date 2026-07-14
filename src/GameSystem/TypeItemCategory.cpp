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
#include "TypeItemCategory.hpp"

#include "../xml.hpp"

namespace Types
{
	ItemCategory::ItemCategory(const TiXmlElement* const config) :
		mId(0),
		mName("")
	{
		if (LoadXmlValue(config, "id", mId))
			mId--;
		
		LoadXmlValue(config, "Name", mName);
	}
	
	std::string ItemCategory::GetName() const
	{
		return mName;
	}
	
	unsigned int ItemCategory::GetId() const
	{
		return mId;
	}
}
