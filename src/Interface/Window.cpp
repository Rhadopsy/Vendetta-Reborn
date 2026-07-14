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
#include "Window.hpp"

#include <sstream>

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "../functions.hpp"

namespace InterfaceModule
{
	Window::Window() :
		mVisible(false),
		mDraging(false),
		mDragX(0),
		mDragY(0),
		mX(0),
		mY(0),
		mWidth(0),
		mHeight(0),
		mMargin(10),
		mTextSize(12),
		mTitle(""),
		mTitleSprite(new legacy::Text()),
		mTitleX(0),
		mTitleY(15),
		mPageIssue(0),
		mPageAmount(1),
		mLeftArrow(NULL),
		mRightArrow(NULL),
		mArrowSize(32),
		mShowKey(sf::Keyboard::A),
		mSprite(NULL),
		mBorderWidth(0),
		mBorderHeight(0)
	{
		mTitleSprite->setColor(sf::Color(0, 0, 0, 255));
	}
	
	Window::~Window()
	{
		delete mTitleSprite;
	}
	
	void Window::Configure(const TiXmlElement* const element)
	{
		if (element)
		{
			double decimal;
			if (LoadXmlValue(element, "TitleSize", decimal) && decimal >= 1)
				mTitleSprite->setSize(decimal);
			
			LoadXmlValue(element, "TextSize", mTextSize);
			LoadXmlValue(element, "TitleY", mTitleY);
			LoadXmlValue(element, "Margin", mMargin);
			
			// TODO
			unsigned int idImage;
			legacy::Image* image = NULL;
			if (LoadXmlValue(element, "Image", idImage) && (image = Configuration->getImage(idImage)) )
			{
				mSprite = new legacy::Sprite(*image);
				mWidth = image->getWidth();
				mHeight = image->getHeight();
			}
			
			LoadXmlValue(element, "BorderWidth", mBorderWidth);
			LoadXmlValue(element, "BorderHeight", mBorderHeight);
			
			LoadXmlValue(element, "Width", mWidth);
			LoadXmlValue(element, "Height", mHeight);
			
			LoadXmlValue(element, "ArrowSize", mArrowSize);
			
			
			// TODO
			if ( (mLeftArrow = Configuration->GetSprite(element, "LeftArrow")) )
				mLeftArrow->resize(mArrowSize, mArrowSize);
			
			if ( (mRightArrow = Configuration->GetSprite(element, "RightArrow")) )
				mRightArrow->resize(mArrowSize, mArrowSize);
		}
	}
	
	void Window::Load(const TiXmlElement* const element)
	{
		if (element)
		{
			LoadXmlValue(element, "Visible", mVisible);
			LoadXmlValue(element, "X", mX);
			LoadXmlValue(element, "Y", mY);
			LoadXmlValue(element, "Width", mWidth);
			LoadXmlValue(element, "Height", mHeight);
		}
	}
	
	TiXmlElement* Window::Save(TiXmlElement* const element) const
	{
		TiXmlElement* saveNode = element ? element : new TiXmlElement("Window");
		
		AddProperty(saveNode, "Visible", mVisible);
		AddProperty(saveNode, "X", mX);
		AddProperty(saveNode, "Y", mY);
		AddProperty(saveNode, "Width", mWidth);
		AddProperty(saveNode, "Height", mHeight);
		
		return saveNode;
	}
	
	void Window::draw(sf::RenderTarget* const render)
	{
		if (mVisible)
		{
			DrawFrame(render, mSprite, mX, mY, mWidth, mHeight, mBorderWidth, mBorderHeight);
			
			if (mTitle != "")
			{
				if (mPageAmount > 1)
				{
					std::ostringstream chaine;
					chaine << mTitle;
					chaine << "(";
					chaine << mPageIssue + 1;
					chaine << "/";
					chaine << mPageAmount;
					chaine << ")";
					mTitleSprite->setText(chaine.str());
				}
				else
					mTitleSprite->setText(mTitle);
				
				mTitleX = (mWidth / 2) - mTitleSprite->getRect().width / 2;
				mTitleSprite->setPosition(mX + mTitleX, mY + mTitleY);
				render->draw(*mTitleSprite);
			}
			
			if (mLeftArrow && mPageIssue > 0)
			{
				mLeftArrow->setPosition(mX + mMargin, mY + mHeight - mArrowSize - mMargin);
				render->draw(*mLeftArrow);
			}
			
			if (mRightArrow && mPageIssue < mPageAmount - 1)
			{
				mRightArrow->setPosition(mX + mWidth - mArrowSize - mMargin, mY + mHeight - mArrowSize - mMargin);
				render->draw(*mRightArrow);
			}
		}
	}
	
	void Window::setPosition(const int& x, const int& y)
	{
		mX = x;
		mY = y;
	}
	
	bool Window::Catch(const int& x, const int& y, const sf::Event& Event, InterfaceModule::DragNDrop* const)
	{
		if (Event.type == sf::Event::KeyPressed && Event.key.code == mShowKey)
		{
			mVisible = !mVisible;
			return true;
		}
		else if (Event.type == sf::Event::MouseMoved)
		{
			if (mDraging)
			{
				mX += x - mDragX;
				mY += y - mDragY;
				
				mDragX = x;
				mDragY = y;
			}
		}
		else if (mVisible)
		{
			if (mDraging && Event.type == sf::Event::MouseButtonReleased && Event.mouseButton.button == sf::Mouse::Left)
			{
				mDraging = false;
				return true;
			}
			else if (mX + mMargin <= x && x <= mX + mMargin + mArrowSize
				&& mY + mHeight - mArrowSize - mMargin <= y && y <= mY + mHeight - mMargin
				&& mPageIssue > 0)
			{
				if (Event.type == sf::Event::MouseButtonReleased)
					mPageIssue--;
				
				return true;
			}
			else if (mX + mWidth - mArrowSize - mMargin <= x && x <= mX + mWidth - mMargin
				&& mY + mHeight - mArrowSize - mMargin <= y && y <= mY + mHeight - mMargin
				&& mPageIssue < mPageAmount - 1)
			{
				if (Event.type == sf::Event::MouseButtonReleased)
					mPageIssue++;
				
				return true;
			}
			else if (Event.type == sf::Event::MouseButtonPressed && Event.mouseButton.button == sf::Mouse::Left)
			{
				if (mX <= x && x <= mX + mWidth && mY <= y && y <= mY + mHeight)
				{
					mDraging = true;
					mDragX = x;
					mDragY = y;
					return true;
				}
				
			}
		}
		
		return false;
	}
	
	std::string Window::getText(const int& x, const int& y) const
	{
		if (mVisible)
		{
			if (mX + mMargin <= x && x <= mX + mMargin + mArrowSize &&
				mY + mHeight - mArrowSize - mMargin <= y && y <= mY + mHeight - mMargin
				&& mPageIssue > 0)
			{
				return "Aller à la page précédente";
			}
			else if (mX + mWidth - mArrowSize - mMargin <= x && x <= mX + mWidth - mMargin &&
					mY + mHeight - mArrowSize - mMargin <= y && y <= mY + mHeight - mMargin
					&& mPageIssue < mPageAmount - 1 )
			{
				return "Aller à la page suivante";
			}
		}
		
		return "";
	}
	
	void Window::SetWidth(const unsigned int& width)
	{
		mWidth = width;
	}
	
	void Window::SetHeight(const unsigned int& height)
	{
		mHeight = height;
	}
}
