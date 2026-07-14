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
#include "Text.hpp"

#include "../Configuration.hpp"

namespace InterfaceModule
{
	Text::Text() :
		mBackground(NULL),
		mText(new legacy::Text(""))
	{
		mText->setColor(sf::Color(0, 0, 0, 255));
		mText->setSize(15);
	}
	
	Text::~Text()
	{
		delete mText;
	}
	
	void Text::Configure(const TiXmlElement* const element)
	{
		mBackground = Configuration->GetSprite(element, "Background");
	}
	
	void Text::draw(sf::RenderTarget* const render)
	{
		if (static_cast<std::string>(mText->getText()) != "")
		{
			if (mBackground)
				render->draw(*mBackground);
			
			render->draw(*mText);
		}
	}
	
	void Text::setPosition(const double& x, const double& y)
	{
		if (mBackground)
			mBackground->setPosition(x, y + mText->getSize() / 5);
		
		mText->setPosition(x, y);
	}
	
	void Text::setText(const std::string& text)
	{
		mText->setText(text);
		
		if (mBackground)
			mBackground->resize(mText->getRect().width, mText->getRect().height);
	}
	
	void Text::setText(const sf::String& text)
	{
		mText->setText(text);
		
		if (mBackground)
			mBackground->resize(mText->getRect().width, mText->getRect().height);
	}
	
	sf::FloatRect Text::getRect() const
	{
		return mText->getRect();
	}
}
