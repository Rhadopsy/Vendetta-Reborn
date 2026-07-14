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
#ifndef TYPE_SKILL_HPP
#define TYPE_SKILL_HPP

#include <string>
#include "../tinyxml/tinyxml.h"

namespace Types
{
	class Skill
	{
		public:
			Skill(const TiXmlElement* const config);
			
			unsigned int GetId() const;
			std::string GetName() const;
			double GetRatio(const double& level) const;
		
		private:
			unsigned int mId;
			std::string mName;
			double mLearningRatio;
			double mEfficiencyRatio;
	};
}

#endif
