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
#ifndef TEXTE_HPP
#define TEXTE_HPP

#include "../compat/SfmlLegacy.hpp"
#include "../tinyxml/tinyxml.h"

namespace InterfaceModule
{
	class Text
	{
		public:
			Text();
			~Text();
			
			void Configure(const TiXmlElement* const element = NULL);
			
			void draw(sf::RenderTarget* const render);
			
			void setPosition(const double& x, const double& y);
			void setText(const std::string& text);
			void setText(const sf::String& text);
			
			sf::FloatRect getRect() const;
		
		private:
			legacy::Sprite* mBackground;
			legacy::Text* mText;
	};
}

#endif
