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
#ifndef TYPE_STATUS_HPP
#define TYPE_STATUS_HPP

#include "../compat/SfmlLegacy.hpp"
#include "../tinyxml/tinyxml.h"

namespace Types
{
	class Status
	{
		public:
			Status(TiXmlElement* const config);
			~Status();
			
			void draw(sf::RenderTarget* const render);
			
			void SetValue(const double& value);
			void setPosition(const double& x, const double& y);
			
			std::string GetName() const;
			double GetDefault() const;
		
		private:
			std::string mName;
			double mDefaultValue;
			double mValue;
			
			double mX;
			double mY;
			
			legacy::Image* mBorderImage;
			legacy::Sprite* mBorderSprite;
			
			legacy::Image* mBackgroundImage;
			legacy::Sprite* mBackgroundSprite;
			
			legacy::Image* mProgressImage;
			legacy::Sprite* mProgressSprite;
			
			legacy::Text* mText;
	};
}

#endif
