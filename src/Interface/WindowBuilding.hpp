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
#ifndef WINDOW_BUILDING_HPP
#define WINDOW_BUILDING_HPP

#include "../compat/SfmlLegacy.hpp"

#include "../GameSystem/Character.hpp"

#include "Window.hpp"
#include "Text.hpp"

namespace InterfaceModule
{
	class WinBuilding : public Window
	{
		public:
			WinBuilding();
			
			void Configure(const TiXmlElement* const element = NULL) override;
			void Load(const TiXmlElement* const element = NULL);
			TiXmlElement* Save() const;
			
			void draw(sf::RenderTarget* const render);
			bool Catch(const int& x, const int& y, const sf::Event& Event, InterfaceModule::DragNDrop* const drag);
			
			using Window::getText;
			std::string getText(const int& x, const int& y, const GameSystem::Character* const personnage) const;
			
			void LoadBuilding(GameSystem::Building* const building);
			void SetPlayer(GameSystem::Character* const player);
		
		private:
			GameSystem::Building* mBuilding;
			Types::Building* mType;
			
			double mNameSize;
			double mNameY;
			double mListX;
			double mListY;
			double mNamesWidth;
			double mLineSpacing;
			
			bool mItems;
			unsigned int mSwitchSize;
			legacy::Sprite* mMaterialsButton;
			legacy::Sprite* mItemsButton;
			double mSwitchX;
			double mSwitchY;
			
			unsigned int mMaterialsPerPage;
			unsigned int mItemsPerPage;
			unsigned int mMaterialPageCount;
			unsigned int mItemPageCount;
			
			int mAmount;
			double mAmountX;
			double mAmountY;
			
			GameSystem::Character* mPlayer;
	};
}

#endif
