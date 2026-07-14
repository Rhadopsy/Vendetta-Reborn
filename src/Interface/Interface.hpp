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
#ifndef INTERFACE_HPP
#define INTERFACE_HPP

#include "../compat/SfmlLegacy.hpp"
#include "../tinyxml/tinyxml.h"

#include "../GameSystem/World.hpp"

#include "Cursor.hpp"
#include "Prompt.hpp"
#include "Text.hpp"
#include "Dragndrop.hpp"

#include "WindowMenu.hpp"
#include "WindowBuilding.hpp"
#include "WindowInventory.hpp"
#include "WindowEquipment.hpp"

namespace InterfaceModule
{
	class Interface
	{
		public:
			Interface();
			~Interface();
			
			bool isOpen() const;
			
			void Configure(TiXmlElement* const element = NULL);
			void Load(const TiXmlElement* const element = NULL);
			TiXmlElement* Save() const;
			
			void draw();
			bool CatchEvent();
			void Process(const double& time);
		
		private:
			bool mLoaded;
			
			sf::RenderWindow* mRender;
			unsigned int mWidth;
			unsigned int mHeight;
			sf::View mWorldView;
			sf::View mUiView;
			double mZoom;
			legacy::Image* mIcon;
			
			legacy::Sprite* mLoadingImage;
			
			CursorStruct mCursor;
			InterfaceModule::DragNDrop* mDrag;
			legacy::Sprite* mSelectingSprite;
			legacy::Sprite* mTargetingSprite;
			
			Types::Building* mSelectedBuilding;
			void SelectBuilding(Types::Building* const building);
			bool mHideBuildPannel;
			unsigned int mBuildPannelColumnCount;
			double mBuildPannelX;
			double mBuildPannelXMin;
			
			double mStatusInterspace;
			double mStatusBottomSpace;
			double mStatusLeftSpace;
			
			Text* mPopup;
			
			GameSystem::World* mWorld;
			GameSystem::Character* mPlayer;
			InterfaceModule::Prompt* mPrompt; // dialog
			
			InterfaceModule::WinMenu* mWinMenu;
			InterfaceModule::WinBuilding* mWinBuilding;
			InterfaceModule::WinInventory* mWinInventory;
			InterfaceModule::WinEquipment* mWinEquipement;
	};
}

#endif
