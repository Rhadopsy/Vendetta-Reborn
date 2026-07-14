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
#ifndef WINDOW_INVENTORY_HPP
#define WINDOW_INVENTORY_HPP

#include "../compat/SfmlLegacy.hpp"
#include <map>
#include "../myvector.hpp"

#include "Window.hpp"
#include "../GameSystem/Character.hpp"

namespace InterfaceModule
{
	class WinInventory : public Window
	{
		public:
			WinInventory();
			
			void Configure(const TiXmlElement* const element = NULL);
			void Load(const TiXmlElement* const element = NULL);
			TiXmlElement* Save() const;
			
			void draw(sf::RenderTarget* const render);
			bool Catch(const int& x, const int& y, const sf::Event& Event, InterfaceModule::DragNDrop* const drag);
			
			void SetInventory(Vector <GameSystem::Item*>* const vector);
		
		private:
			double mCellSize;
			double mGridX;
			double mGridY;
			double mGridMargin;
			
			Vector <GameSystem::Item*>* mInventory;
			
			unsigned int mColumnCount;
			unsigned int mLineCount;
			legacy::Sprite* mSpriteGrilleInventory;
	};
}

#endif
