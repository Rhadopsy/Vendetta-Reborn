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
#include "Cursor.hpp"

#include "../Configuration.hpp"
#include "../xml.hpp"

namespace InterfaceModule
{
	CursorStruct::CursorStruct() :
		mRenderX(0),
		mRenderY(0),
		mWorldX(0),
		mWorldY(0),
		mSize(24),
		mShiftX(0),
		mShiftY(0),
		mDefault(NULL),
		mSuitableBuilding(NULL),
		mUnsuitableBuilding(NULL)
	{
	}
	
	void CursorStruct::Configure(TiXmlElement* const element)
	{
		if (element)
		{
			LoadXmlValue(element, "Size", mSize);
			if ( (mDefault = Configuration->GetSprite(element, "Sprite")) )
				mDefault->resize(mSize, mSize);
			
			mSuitableBuilding = Configuration->GetSprite(element, "SuitableBuilding");
			mUnsuitableBuilding = Configuration->GetSprite(element, "UnsuitableBuilding");
		}
	}
	
	void CursorStruct::draw(sf::RenderTarget* const render, bool constructible, double width, double height)
	{
		if (width && height && constructible && mSuitableBuilding)
		{
			mSuitableBuilding->resize(width, height);
			mSuitableBuilding->setPosition(mRenderX, mRenderY);
			render->draw(*mSuitableBuilding);
		}
		else if (width && height && mUnsuitableBuilding)
		{
			mUnsuitableBuilding->resize(width, height);
			mUnsuitableBuilding->setPosition(mRenderX, mRenderY);
			render->draw(*mUnsuitableBuilding);
		}
		else if (mDefault)
		{
			mDefault->setPosition(mRenderX, mRenderY);
			render->draw(*mDefault);
		}
	}
	
	void CursorStruct::SetShift(double x, double y)
	{
		mShiftX = x;
		mShiftY = y;
	}
	
	void CursorStruct::setPosition(const int& x, const int& y, sf::RenderWindow* const render, sf::View View)
	{
		mRenderX = x + mShiftX;
		mRenderY = y + mShiftY;
		
		sf::Vector2f worldPosition = render->mapPixelToCoords(sf::Vector2i(mRenderX, mRenderY), View);
		mWorldX = worldPosition.x;
		mWorldY = worldPosition.y;
	}
}
