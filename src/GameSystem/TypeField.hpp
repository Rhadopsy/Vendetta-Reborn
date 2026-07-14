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
#ifndef TYPE_FIELD_HPP
#define TYPE_FIELD_HPP

#include "../compat/SfmlLegacy.hpp"
#include "../tinyxml/tinyxml.h"

namespace Types
{
	class Field
	{
		public:
			Field(TiXmlElement* const config);
			
			legacy::Sprite* GetSprite() const;
			bool AcceptRessource() const;
			double GetSpeedRatio() const;
			bool Buildable() const;
			double GetProbability() const;
		
		private:
			double mProbability;
			double mSpeedRatio;
			bool mBuildable;
			bool mAcceptRessource;
			legacy::Sprite* mSprite;
	};
}

#endif
