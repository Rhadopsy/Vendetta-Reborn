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
#ifndef CURSOR_HPP
#define CURSOR_HPP

#include "../compat/SfmlLegacy.hpp"
#include "../tinyxml/tinyxml.h"

namespace InterfaceModule
{
	struct CursorStruct
	{
		CursorStruct();
		void Configure(TiXmlElement* const element);
		void draw(sf::RenderTarget* const render, bool constructible = false, double width = 0, double height = 0);
		
		void SetShift(double x, double y);
		void setPosition(const int& x, const int& y, sf::RenderWindow* const render, sf::View View);
		
		int mRenderX;
		int mRenderY;
		double mWorldX;
		double mWorldY;
		unsigned int mSize;
		
		double mShiftX;
		double mShiftY;
		legacy::Sprite* mDefault;
		legacy::Sprite* mSuitableBuilding;
		legacy::Sprite* mUnsuitableBuilding;
	};
}

#endif
