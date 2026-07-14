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
#include "WindowInventory.hpp"

#include <sstream>
#include <cmath>

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "../data.hpp"

namespace InterfaceModule
{
	WinInventory::WinInventory() : Window(),
		mCellSize(32),
		mGridX(10),
		mGridY(90),
		mGridMargin(10),
		mInventory(NULL),
		mColumnCount(0),
		mLineCount(0),
		mSpriteGrilleInventory(NULL)
	{
		mShowKey = sf::Keyboard::I;
		mTitle = "Inventory";
	}
	
	void WinInventory::Configure(const TiXmlElement* const element)
	{
		Window::Configure(element);
		
		if (element)
		{
			LoadXmlValue(element, "CellSize", mCellSize);
			mSpriteGrilleInventory = Configuration->GetSprite(element, "GridSprite");
			LoadXmlValue(element, "GridX", mGridX);
			LoadXmlValue(element, "GridY", mGridY);
			LoadXmlValue(element, "GridMargin", mGridMargin);
			LoadXmlValue(element, "ColumnCount", mColumnCount);
			LoadXmlValue(element, "LineCount", mLineCount);
		}
	}
	
	void WinInventory::Load(const TiXmlElement* const element)
	{
		this->Window::Load(element);
		
		LoadXmlValue(element, "PageIssue", mPageIssue);
	}
	
	TiXmlElement* WinInventory::Save() const
	{
		TiXmlElement* saveNode = new TiXmlElement("WinInventory");
		this->Window::Save(saveNode);
		
		AddProperty(saveNode, "PageIssue", mPageIssue);
		
		return saveNode;
	}
	
	void WinInventory::draw(sf::RenderTarget* const render)
	{
		mPageAmount = mPageAmount;
		Window::draw(render);
		
		if (mVisible && mInventory)
		{
			if (mSpriteGrilleInventory)
			{
				mSpriteGrilleInventory->setPosition(mX + mGridX, mY + mGridY);
				render->draw(*mSpriteGrilleInventory);
			}
			
			double shift = mCellSize / 2 - Prgm::Items::IconSize / 2;
			
			unsigned int i = 0;
			unsigned int itemCount = (*mInventory).size();
			unsigned int column = 0;
			unsigned int line = 0;
			unsigned int page = 0;
			while (i < itemCount)
			{
				if (page == mPageIssue && (*mInventory)[i])
				{
					(*mInventory)[i]->DrawIcon(render,
						mX + mGridX + shift + mCellSize * column,
						mY + mGridY + mGridMargin + shift + mCellSize * line);
				}
				
				column++;
				if (column >= mColumnCount)
				{
					column = 0;
					line++;
					if (column >= mLineCount)
					{
						line = 0;
						page++;
					}
				}
				
				i++;
			}
		}
	}
	
	bool WinInventory::Catch(const int& x, const int& y, const sf::Event& Event, InterfaceModule::DragNDrop* const drag)
	{
		if (mVisible && Event.type == sf::Event::MouseButtonReleased && Event.mouseButton.button == sf::Mouse::Left)
		{
			int boxColumn = floor( (x - mX - mGridX) / mCellSize );
			int boxLine = floor( (y - mY - mGridY - mGridMargin) / mCellSize );
			
			if (0 <= boxLine && 0 <= boxColumn)
			{
				unsigned int line = static_cast<unsigned int>(boxLine);
				unsigned int column = static_cast<unsigned int>(boxColumn);
				
				if (column < mColumnCount
					&& line < mLineCount
					&& (line != mLineCount - 1 || (column != 0 && column != mColumnCount - 1)) )
				{
					unsigned int i = (mPageIssue * mLineCount + line) * mColumnCount + column;
					
					if (i < (*mInventory).size() && (*mInventory)[i])
					{
						if (drag->object && (drag->object->GetClass() == GameSystem::Classes::Item))
						{
							/* TODO
							GameSystem::Item* item = static_cast<GameSystem::Item*>(drag->object);
							*drag->callback = (*mInventory)[i];
							(*mInventory)[i] = item;
							drag->object = NULL;
							return true;
							*/
						}
						else
						{
							/* TODO
							GameSystem::Object* item = static_cast<GameSystem::Object*>(drag->object);
							drag->object = Object;
							drag->callback = &( (*mInventory)[i] );
							(*mInventory)[i] = NULL;
							return true;
							*/
						}
					}
				}
			}
		}
		
		return Window::Catch(x, y, Event, drag);
	}
	
	void WinInventory::SetInventory(Vector <GameSystem::Item*>* const vector)
	{
		mInventory = vector;
	}
}
