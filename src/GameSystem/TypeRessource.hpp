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
#ifndef TYPE_RESSOURCE_HPP
#define TYPE_RESSOURCE_HPP

#include "../tinyxml/tinyxml.h"

#include "TypeMaterial.hpp"

namespace Types
{
	class Ressource
	{
		public:
			Ressource(TiXmlElement* const config);
			~Ressource();
			
			void draw(sf::RenderTarget* const render, const double& x, const double& y);
			
			unsigned int GetId() const;
			unsigned int getWidth() const;
			unsigned int getHeight() const;
			unsigned int GetResId() const;
			std::string GetAction() const;
			Types::Material* GetTypeMaterial() const;
			double GetHarvestRate() const;
		
		private:
			unsigned int mId;
			legacy::Sprite* mSprite;
			unsigned int mWidth;
			unsigned int mHeight;
			std::string mName;
			std::string mActionList;
			Types::Material* mMaterial;
			double mHarvestRate;
	};
}

#endif
