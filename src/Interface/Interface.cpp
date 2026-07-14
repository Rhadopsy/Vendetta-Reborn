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
#include "Interface.hpp"

#include <cmath>
#include <filesystem>
#include <vector>

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "../functions.hpp"
#include "../data.hpp"

namespace InterfaceModule
{
	Interface::Interface() :
		mLoaded(false),
		mRender(NULL),
		mWidth(0),
		mHeight(0),
		mZoom(1),
		mIcon(new legacy::Image()),
		mLoadingImage(new legacy::Sprite()),
		mDrag(new DragNDrop()),
		mSelectingSprite(NULL),
		mTargetingSprite(NULL),
		mSelectedBuilding(NULL),
		mHideBuildPannel(false),
		mBuildPannelX(0),
		mBuildPannelXMin(0),
		mStatusInterspace(10),
		mStatusBottomSpace(5),
		mStatusLeftSpace(5),
		mPopup(new Text()),
		mWorld(new GameSystem::World()),
		mPlayer(NULL),
		mPrompt(new Prompt("> ")),
		mWinMenu(NULL),
		mWinBuilding(NULL),
		mWinInventory(NULL),
		mWinEquipement(NULL)
	{
		mRender = new sf::RenderWindow(sf::VideoMode(1280, 720), "Vendetta", sf::Style::Default);
		mRender->setFramerateLimit(0);
		mRender->setMouseCursorVisible(false);
		mRender->setVerticalSyncEnabled(true);
		
		mWorldView = mRender->getDefaultView();
		mUiView = mRender->getDefaultView();
		
		if (mIcon->loadFromFile("./Data/Images/icone.png"))
			mRender->setIcon(16, 16, mIcon->getPixelsPtr());
		
		mWidth = mRender->getSize().x;
		mHeight = mRender->getSize().y;
		
		mCursor.setPosition(mWidth / 2, mHeight / 2, mRender, mWorldView);
		
		// chargement de l'image d'attente
			std::vector<std::filesystem::path> imageList;
			const std::filesystem::path loadingDirectory("./Data/Images/Chargement/");
			if (std::filesystem::exists(loadingDirectory))
			{
				for (const auto& entry : std::filesystem::directory_iterator(loadingDirectory))
					if (entry.is_regular_file())
						imageList.push_back(entry.path());
			}

			legacy::Image image;
			if (!imageList.empty())
				image.loadFromFile(imageList[legacy::random(std::size_t{0}, imageList.size() - 1)].string());
			else
				image.create(1, 1, sf::Color::Black);
			mLoadingImage->setImage(image);
		
		mWorld->SetPlayer(&mPlayer);
		
		mWinMenu = new WinMenu();
		mWinBuilding = new WinBuilding();
		mWinInventory = new WinInventory();
		mWinEquipement = new WinEquipment();
	}
	
	Interface::~Interface()
	{
		delete mWorld;
		delete mDrag;
		delete mSelectingSprite;
		delete mTargetingSprite;
		
		delete mPopup;
		delete mPrompt;
		
		delete mWinMenu;
		delete mWinBuilding;
		delete mWinInventory;
		delete mWinEquipement;
		
		delete mLoadingImage;
		
		delete mRender;
		delete mIcon;
	}
	
	bool Interface::isOpen() const
	{
		return mRender->isOpen();
	}
	
	void Interface::Configure(TiXmlElement* const element)
	{
		if (element)
		{
			mPopup->Configure(element->FirstChildElement("Popup"));
			mCursor.Configure(element->FirstChildElement("Cursor"));
			
			unsigned int pannelHeight;
			LoadXmlValue(element, "BuildPannelHeight", pannelHeight);
			mBuildPannelColumnCount = Configuration->GetNombreTypesBuildings() / pannelHeight;
			mBuildPannelXMin = 2 - mBuildPannelColumnCount * Prgm::TypesBuildings::ButtonSize;
			mBuildPannelX = mBuildPannelXMin;
			
			unsigned int imageId;
			legacy::Image* image = NULL;
			
			// TODO
			if (LoadXmlValue(element, "SelectingSprite", imageId) && (image = Configuration->getImage(imageId)) )
				mSelectingSprite = new legacy::Sprite(*image);
			
			if (LoadXmlValue(element, "CadreTarget", imageId) && (image = Configuration->getImage(imageId)) )
				mTargetingSprite = new legacy::Sprite(*image);
			
			LoadXmlValue(element, "StatusInterspace", mStatusInterspace);
			LoadXmlValue(element, "StatusBottomSpace", mStatusBottomSpace);
			LoadXmlValue(element, "StatusLeftSpace", mStatusLeftSpace);
			
			TiXmlElement* selement = element->FirstChildElement("Windows");
			
			mWinMenu->Window::Configure(selement);
			mWinBuilding->Window::Configure(selement);
			mWinInventory->Window::Configure(selement);
			mWinEquipement->Window::Configure(selement);
			
			mWinMenu->Configure(element->FirstChildElement("WinMenu"));
			mWinBuilding->Configure(element->FirstChildElement("WinBuilding"));
			mWinInventory->Configure(element->FirstChildElement("WinInventory"));
			mWinEquipement->Configure(element->FirstChildElement("WinEquipment"));
			
			mWinMenu->SetWidth(mWidth);
		}
	}
	
	void Interface::Load(const TiXmlElement* const element)
	{
		mWorld->Load(element ? element->FirstChildElement("World") : NULL);
		mPrompt->ReturnTo(mPlayer);
		
		if (element)
		{
			LoadXmlValue(element, "zoom", mZoom);
			mWorldView.zoom(mZoom);
			
			mWinMenu->Load(element->FirstChildElement("WinMenu"));
			mWinBuilding->Load(element->FirstChildElement("WinBuilding"));
			mWinInventory->Load(element->FirstChildElement("WinInventory"));
			mWinEquipement->Load(element->FirstChildElement("WinEquipment"));
		}
		else
		{
			mWinMenu->setPosition(0, -10);
			mWinBuilding->setPosition(0, 0);
			mWinInventory->setPosition(0, 0);
			mWinEquipement->setPosition(0, 0);
		}
		
		mWinBuilding->SetPlayer(mPlayer);
		mWinInventory->SetInventory(mPlayer->GetInventory());
		mWinEquipement->SetPlayer(mPlayer);
		
		mCursor.setPosition(mCursor.mRenderX, mCursor.mRenderY, mRender, mWorldView);
		
		mRender->setMouseCursorVisible(false);
		mLoaded = true;
	}
	
	TiXmlElement* Interface::Save() const
	{
		TiXmlElement* saveNode = new TiXmlElement("Save");
		
		TiXmlElement* selement = new TiXmlElement("zoom");
		saveNode->LinkEndChild(selement);
		selement->SetDoubleAttribute("value", mZoom);
		
		saveNode->LinkEndChild( mWinMenu->Save() );
		saveNode->LinkEndChild( mWinBuilding->Save() );
		saveNode->LinkEndChild( mWinInventory->Save() );
		saveNode->LinkEndChild( mWinEquipement->Save() );
		
		saveNode->LinkEndChild( mWorld->Save() );
		
		return saveNode;
	}
	
	void Interface::draw()
	{
		mRender->clear();
		
		if (mLoaded)
		{
			mRender->setView(mWorldView);
			mWorld->draw(mRender);
			
			DrawFrame(mRender, mSelectingSprite, mPlayer->GetXSprite(), mPlayer->GetYSprite() + 2, mPlayer->getWidth(), mPlayer->getHeight(), 5, 5);
			if (GameSystem::Object* target = mPlayer->GetTarget())
				DrawFrame(mRender, mTargetingSprite, target->GetXSprite(), target->GetYSprite(), target->getWidth() + 2, target->getHeight(), 5, 5);
			
			mRender->setView(mUiView);
			
			mPrompt->draw(mRender);
			
			unsigned int statusCount = Configuration->GetNombreStatuses();
			unsigned int x = mStatusLeftSpace;
			unsigned int i = 0;
			Types::Status* statut;
			while (i < statusCount)
			{
				if ((statut = Configuration->GetStatus(i++)))
				{
					statut->setPosition(x, mHeight - Prgm::Statuses::Height - mStatusBottomSpace);
					x += mStatusInterspace + Prgm::Statuses::Width;
					statut->draw(mRender);
				}
			}
			
			Types::Building* typeBuilding;
			unsigned int typesBuildingCount = Configuration->GetNombreTypesBuildings();
			unsigned int column = 0;
			unsigned int line = 0;
			i = 0;
			while (i < typesBuildingCount)
			{
				if ( (typeBuilding = Configuration->GetTypeBuilding(i)) )
				{
					typeBuilding->SetPositionButton(Prgm::TypesBuildings::ButtonSize * column + mBuildPannelX, 50 + Prgm::TypesBuildings::ButtonSize * line);
					typeBuilding->DrawButton(mRender);
					
					column++;
					if (column >= mBuildPannelColumnCount)
					{
						column = 0;
						line++;
					}
				}
				
				i++;
			}
			
			mWinMenu->draw(mRender);
			mWinInventory->draw(mRender);
			mWinBuilding->LoadBuilding(mPlayer->GetCurrentBuilding());
			mWinBuilding->draw(mRender);
			mWinEquipement->draw(mRender);
			
			if (mDrag->object)
			{
				if (mDrag->object->GetClass() == GameSystem::Classes::Item)
				{
					GameSystem::Item* item = static_cast<GameSystem::Item*>(mDrag->object);
					item->DrawIcon(mRender, mCursor.mRenderX - Prgm::Items::IconSize / 2,  mCursor.mRenderY - Prgm::Items::IconSize / 2);
				}
			}
			
			std::string popupText = "";
			
			typeBuilding = NULL;
			i = 0;
			if (mHideBuildPannel)
			{
				while (i < typesBuildingCount && ( !(typeBuilding = Configuration->GetTypeBuilding(i++)) || !typeBuilding->Catch(mCursor.mRenderX, mCursor.mRenderY) ) );
				i--;
			}
			if (!mHideBuildPannel || ( typeBuilding && typeBuilding->Catch(mCursor.mRenderX, mCursor.mRenderY) ) )
			{
				popupText = mPlayer->GetTextTypeBuilding(Configuration->GetTypeBuilding(i));
				
				mBuildPannelX += 5;
				
				if (mBuildPannelX >= 0)
					mBuildPannelX = 0;
			}
			else
			{
				mBuildPannelX -= 5;
				
				if (mBuildPannelX < mBuildPannelXMin)
					mBuildPannelX = mBuildPannelXMin;
			}
			if (popupText == "" && (popupText = mWinBuilding->getText(mCursor.mRenderX, mCursor.mRenderY, mPlayer)) == "")
				popupText = mPlayer->GetTextObject(mWorld->GetObjectPoint(mCursor.mWorldX, mCursor.mWorldY));
			mPopup->setText(popupText);
			mPopup->setPosition(mCursor.mRenderX + mCursor.mSize, mCursor.mRenderY + mCursor.mSize);
			mPopup->draw(mRender);
			
			if (mSelectedBuilding)
				mCursor.draw(mRender, mWorld->Buildable(mCursor.mWorldX, mCursor.mWorldY, mSelectedBuilding), mSelectedBuilding->GetStructureWidth() * mZoom, mSelectedBuilding->GetStructureHeight() * mZoom);
			else
				mCursor.draw(mRender);
		}
		else
		{
			DrawGauge(mRender, 0, 0, mWidth, mHeight, Prgm::Progression);
			mLoadingImage->resize(mWidth, mHeight);
			mRender->draw(*mLoadingImage);
		}
		
		mRender->display();
	}
	
	bool Interface::CatchEvent()
	{
		sf::Event Event;
		if (!mRender->pollEvent(Event))
			return false;
		
		if ( (Event.type == sf::Event::KeyPressed && Event.key.code == sf::Keyboard::Escape) || Event.type == sf::Event::Closed)
		{
			if (mLoaded)
			{
				TiXmlDeclaration* const declaration = new TiXmlDeclaration("2.0", "utf-8", "");
				TiXmlElement* const root = this->Save();
				
				TiXmlDocument* const sauvegarde = new TiXmlDocument();
				sauvegarde->LinkEndChild(declaration);
				sauvegarde->LinkEndChild(root);
				sauvegarde->SaveFile("./Data/Saves/test.xml");
				delete sauvegarde;
			}
			mRender->close();
		}
		else if (Event.type == sf::Event::MouseMoved)
		{
			mCursor.setPosition(Event.mouseMove.x, Event.mouseMove.y, mRender, mWorldView);
		}
		else if (Event.type == sf::Event::Resized)
		{
			mWidth = Event.size.width;
			mHeight = Event.size.height;
			
			mUiView.reset(sf::FloatRect(0.f, 0.f, static_cast<float>(mWidth), static_cast<float>(mHeight)));
		}
		
		if (!mLoaded)
			return false;
		
		if (Event.type == sf::Event::MouseWheelMoved && sf::Keyboard::isKeyPressed(sf::Keyboard::LControl))
		{
			mZoom *= pow(1.1, Event.mouseWheel.delta);
			mWorldView.zoom(pow(1.1, Event.mouseWheel.delta));
		}
		else if (!mPrompt->Catch(Event) && 
			!mWinMenu->Catch(mCursor.mRenderX, mCursor.mRenderY, Event, mDrag) &&
			!mWinBuilding->Catch(mCursor.mRenderX, mCursor.mRenderY, Event, mDrag) &&
			!mWinInventory->Catch(mCursor.mRenderX, mCursor.mRenderY, Event, mDrag) &&
			!mWinEquipement->Catch(mCursor.mRenderX, mCursor.mRenderY, Event, mDrag) &&
			Event.type == sf::Event::MouseButtonReleased )
		{
			if (mDrag->object)
			{
				(*mDrag->callback) = mDrag->object;
				mDrag->object = NULL;
			}
			else if (Event.mouseButton.button == sf::Mouse::Left)
			{
				Types::Building* typeBuilding;
				unsigned int typesBuildingCount = Configuration->GetNombreTypesBuildings();
				unsigned int i = 0;
				while (i < typesBuildingCount)
				{
					if ( (typeBuilding = Configuration->GetTypeBuilding(i)) && typeBuilding->Catch(mCursor.mRenderX, mCursor.mRenderY) )
					{
						this->SelectBuilding(Configuration->GetTypeBuilding(i));
						return true;
					}
					
					i++;
				}
				
				if (mSelectedBuilding)
				{
					if (mWorld->SeatBuilding(mCursor.mWorldX, mCursor.mWorldY, mSelectedBuilding))
						this->SelectBuilding(NULL);
				}
				else
				{
					if (!sf::Keyboard::isKeyPressed(sf::Keyboard::RShift) && !sf::Keyboard::isKeyPressed(sf::Keyboard::LShift))
						mPlayer->ResetActions();
					
					if (GameSystem::Object* target = mWorld->GetObjectPoint(mCursor.mWorldX, mCursor.mWorldY))
						mPlayer->onTargetSelected(target);
					else
						mPlayer->GoCoords(mCursor.mWorldX, mCursor.mWorldY);
				}
			}
			else if (Event.mouseButton.button == sf::Mouse::Right && mSelectedBuilding)
				this->SelectBuilding(NULL);
		}
		
		return true;
	}
	
	void Interface::Process(const double& time)
	{
		if (mLoaded)
		{
			mWorld->Process(time);
			mWorldView.setCenter(mPlayer->GetX(), mPlayer->GetY());
		}
	}
	
	void Interface::SelectBuilding(Types::Building* const building)
	{
		mSelectedBuilding = building;
		
		if (mSelectedBuilding)
			mCursor.SetShift(-static_cast<float>(mSelectedBuilding->GetStructureWidth()) / 2, static_cast<float>(mSelectedBuilding->GetStructureHeight()) / 2);
		else
			mCursor.SetShift(0, 0);
	}
}
