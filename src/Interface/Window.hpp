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
#ifndef WINDOW_HPP
#define WINDOW_HPP

#include "../compat/SfmlLegacy.hpp"
#include "../tinyxml/tinyxml.h"

#include "Dragndrop.hpp"

namespace InterfaceModule
{
	class Window
	{
		public:
			Window();
			virtual ~Window();
			
			virtual void Configure(const TiXmlElement* const element = NULL);
			virtual void Load(const TiXmlElement* const element = NULL);
			virtual TiXmlElement* Save(TiXmlElement* const element = NULL) const;
			
			virtual void draw(sf::RenderTarget* const render);
			virtual void setPosition(const int& x, const int& y);
			virtual bool Catch(const int& x, const int& y, const sf::Event& Event, InterfaceModule::DragNDrop* const drag);
			
			virtual std::string getText(const int& x, const int& y) const;
			
			virtual void SetWidth(const unsigned int& width);
			virtual void SetHeight(const unsigned int& height);
		
		protected:
			bool mVisible;
			bool mDraging;
			unsigned int mDragX;
			unsigned int mDragY;
			int mX;
			int mY;
			int mWidth;
			int mHeight;
			
			double mMargin;
			double mTextSize;
			
			std::string mTitle;
			legacy::Text* mTitleSprite;
			double mTitleX;
			double mTitleY;
			
			unsigned int mPageIssue;
			unsigned int mPageAmount;
			legacy::Sprite* mLeftArrow;
			legacy::Sprite* mRightArrow;
			double mArrowSize;
			
			sf::Keyboard::Key mShowKey;
			
			legacy::Sprite* mSprite;
			int mBorderWidth;
			int mBorderHeight;
	};
}

#endif
