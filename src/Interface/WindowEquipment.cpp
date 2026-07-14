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
#include "WindowEquipment.hpp"

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "../data.hpp"

namespace InterfaceModule
{
	WinEquipment::WinEquipment() :
		mPlayer(NULL),
		mEquipment(NULL),
		mEmplacementSprite(NULL),
		mPlayerIllustration(NULL)
	{
		mShowKey = sf::Keyboard::E;
		mTitle = "Équipement";
	}
	
	void WinEquipment::Configure(TiXmlElement* const element)
	{
		Window::Configure(element);
		mEmplacementSprite = Configuration->GetSprite(element, "EmplacementSprite");
	}
	
	TiXmlElement* WinEquipment::Save() const
	{
		TiXmlElement* saveNode = new TiXmlElement("WinEquipment");
		this->Window::Save(saveNode);
		return saveNode;
	}
	
	void WinEquipment::draw(sf::RenderTarget* const render)
	{
		Window::draw(render);
		
		if (mVisible)
		{
			if (mPlayerIllustration)
			{
				mPlayerIllustration->setPosition(mX + mWidth / 2 - mPlayerIllustration->getSize().x / 2, mY + mHeight / 2 - mPlayerIllustration->getSize().y / 2);
				render->draw(*mPlayerIllustration);
			}
			
			unsigned int i = 0;
			while (i < mEquipmentSpotCategories.size())
			{
				mEmplacementSprite->setPosition(mX + mEquipmentSpotAbscissas[i], mY + mEquipmentSpotOrdinates[i]);
				render->draw(*mEmplacementSprite);
				
				if (mEquipment && (*mEquipment)[i])
					(*mEquipment)[i]->DrawIcon(render, mX + mEquipmentSpotAbscissas[i] + mEmplacementSprite->getSize().x / 2 - Prgm::Items::IconSize / 2, mY + mEquipmentSpotOrdinates[i] + mEmplacementSprite->getSize().y / 2 - Prgm::Items::IconSize / 2);
				
				i++;
			}
		}
	}
	
	bool WinEquipment::Catch(const int& x, const int& y, const sf::Event& Event, InterfaceModule::DragNDrop* const drag)
	{
		if (mVisible)
		{
			if (drag->object && (drag->object->GetClass() == GameSystem::Classes::Item) && Event.type == sf::Event::MouseButtonReleased)
			{
				GameSystem::Item* item = static_cast<GameSystem::Item*>(drag->object);
				unsigned int categorieItem = item->GetType()->GetItemCategory()->GetId();
				unsigned int i = 0;
				while (i < mEquipmentSpotCategories.size())
				{
					if (mEquipmentSpotCategories[i] == categorieItem)
					{
						(*drag->callback) = (*mEquipment)[i];
						(*mEquipment)[i] = item;
						drag->object = NULL;
						return true;
					}
					
					i++;
				}
			}
			else
			{
				unsigned int i = 0;
				while (i < mEquipmentSpotCategories.size())
				{
					if (mEquipment && (*mEquipment)[i] && (*mEquipment)[i]->GetType()->Catch(x, y, mX + mEquipmentSpotAbscissas[i] + mEmplacementSprite->getSize().x / 2 - Prgm::Items::IconSize / 2, mY + mEquipmentSpotOrdinates[i] + mEmplacementSprite->getSize().y / 2 - Prgm::Items::IconSize / 2))
					{
						if (Event.type == sf::Event::MouseButtonPressed && Event.mouseButton.button == sf::Mouse::Left)
						{
							/* TODO
							drag->object = (*mEquipment)[i];
							drag->callback = &( (*mEquipment)[i] );
							(*mEquipment)[i] = NULL;
							return true;
							*/
						}
						else if (Event.type == sf::Event::MouseButtonReleased && Event.mouseButton.button == sf::Mouse::Right)
						{
							mPlayer->AddItemToStock((*mEquipment)[i]);
							(*mEquipment)[i] = NULL;
							return true;
						}
					}
					
					i++;
				}
			}
		}
		
		return Window::Catch(x, y, Event, drag);
	}
	
	void WinEquipment::SetPlayer(GameSystem::Character* const player)
	{
		mPlayer = player;
		if (mPlayer)
		{
			mEquipment = mPlayer->GetItemsEquipes();
			Types::Race* race = mPlayer->GetRace();
			if (race)
			{
				mPlayerIllustration = Configuration->GetSprite(race->GetSpriteEquipementIllustration());
				mEquipmentSpotCategories = race->GetEquipmentSpotCategories();
				mEquipmentSpotAbscissas = race->GetEquipmentSpotAbscissas();
				mEquipmentSpotOrdinates = race->GetEquipmentSpotOrdinates();
			}
		}
	}
}
