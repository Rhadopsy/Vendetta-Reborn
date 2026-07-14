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
#ifndef WINDOW_EQUIPMENT_HPP
#define WINDOW_EQUIPMENT_HPP

#include "Window.hpp"
#include "../GameSystem/Character.hpp"

namespace InterfaceModule
{
	class WinEquipment : public Window
	{
		public:
			WinEquipment();
			
			void Configure(TiXmlElement* const element = NULL);
			TiXmlElement* Save() const;
			
			void draw(sf::RenderTarget* const render);
			bool Catch(const int& x, const int& y, const sf::Event& Event, InterfaceModule::DragNDrop* const drag);
			
			void SetPlayer(GameSystem::Character* const player);
		
		private:
			GameSystem::Character* mPlayer;
			
			Vector <GameSystem::Item*>* mEquipment;
			Vector <unsigned int> mEquipmentSpotCategories;
			Vector <unsigned int> mEquipmentSpotAbscissas;
			Vector <unsigned int> mEquipmentSpotOrdinates;
			
			legacy::Sprite* mEmplacementSprite;
			legacy::Sprite* mPlayerIllustration;
	};
}

#endif
