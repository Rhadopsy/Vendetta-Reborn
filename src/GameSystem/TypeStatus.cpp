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
#include "TypeStatus.hpp"

#include "../xml.hpp"
#include "../data.hpp"

namespace Types
{
	Status::Status(TiXmlElement* const config) :
			mName(""),
			mDefaultValue(0),
			mValue(0),
			mX(0),
			mY(0),
			mBorderImage(new legacy::Image(1, 1, sf::Color(127, 127, 0, 255))),
			mBorderSprite(new legacy::Sprite(*mBorderImage)),
			mBackgroundImage(new legacy::Image(1, 1, sf::Color(0, 0, 0, 255))),
			mBackgroundSprite(new legacy::Sprite(*mBackgroundImage)),
			mProgressImage(new legacy::Image(1, 1, sf::Color(255, 255, 255, 255))),
			mProgressSprite(new legacy::Sprite(*mProgressImage)),
			mText(new legacy::Text(mName))
	{	
		LoadXmlValue(config, "Name", mName);
		
		LoadXmlValue(config, "DefaultValue", mDefaultValue);
		if (mDefaultValue < 0)
			mDefaultValue = 0;
		
		mBorderSprite->resize(Prgm::Statuses::Width, Prgm::Statuses::Height);
		mBackgroundSprite->resize(Prgm::Statuses::Width - 2 * Prgm::Statuses::Border, Prgm::Statuses::Height - 2 * Prgm::Statuses::Border);
		
		mText->setColor(sf::Color(0, 0, 0, 255));
		mText->setSize(Prgm::Statuses::TextSize);
	}
	
	Status::~Status()
	{
		delete mBorderImage;
		delete mBorderSprite;
		
		delete mBackgroundImage;
		delete mBackgroundSprite;
		
		delete mProgressImage;
		delete mProgressSprite;
		
		delete mText;
	}
	
	void Status::draw(sf::RenderTarget* const render)
	{
		mBorderSprite->setPosition(mX, mY);
		render->draw(*mBorderSprite);
		
		mBackgroundSprite->setPosition(mX + Prgm::Statuses::Border, mY + Prgm::Statuses::Border);
		render->draw(*mBackgroundSprite);
		
		mProgressSprite->setPosition(mX + Prgm::Statuses::Border, mY + Prgm::Statuses::Border);
		render->draw(*mProgressSprite);
		
		mText->setPosition(mX + Prgm::Statuses::Border + Prgm::Statuses::Width / 2 - mText->getRect().width / 2, mY + Prgm::Statuses::Border);
		render->draw(*mText);
	}
	
	void Status::SetValue(const double& value)
	{
		mValue = value;
		mProgressSprite->resize(Prgm::Statuses::Width * mValue / 20, Prgm::Statuses::Height);
	}
	
	void Status::setPosition(const double& x, const double& y)
	{
		mX = x;
		mY = y;
	}
	
	std::string Status::GetName() const
	{
		return mName;
	}
	
	double Status::GetDefault() const
	{
		return mDefaultValue;
	}
}
