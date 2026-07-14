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
#include "WindowBuilding.hpp"

#include <sstream>
#include <cmath>

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "../data.hpp"
#include "../functions.hpp"
#include "../GameSystem/TypeItem.hpp"
#include "../GameSystem/TypeBuilding.hpp"

namespace InterfaceModule
{
	WinBuilding::WinBuilding() : Window(),
		mBuilding(NULL),
		mType(NULL),
		mNameSize(14),
		mNameY(30),
		mListX(20),
		mListY(60),
		mNamesWidth(200),
		mLineSpacing(10),
		mItems(false),
		mSwitchSize(24),
		mMaterialsButton(NULL),
		mItemsButton(NULL),
		mSwitchX(240),
		mSwitchY(30),
		mMaterialsPerPage(0),
		mItemsPerPage(0),
		mMaterialPageCount(0),
		mItemPageCount(0),
		mAmount(0),
		mAmountX(10),
		mAmountY(30),
		mPlayer(NULL)
	{
		mShowKey = sf::Keyboard::B;
		mTitle = "Batiment";
	}
	
	void WinBuilding::Configure(const TiXmlElement* const element)
	{
		Window::Configure(element);
		
		if (element)
		{
			LoadXmlValue(element, "NameSize", mNameSize);
			LoadXmlValue(element, "NameY", mNameY);
			LoadXmlValue(element, "ListX", mListX);
			LoadXmlValue(element, "ListY", mListY);
			LoadXmlValue(element, "NamesWidth", mNamesWidth);
			LoadXmlValue(element, "LineSpacing", mLineSpacing);
			
			if ( (mMaterialsButton = Configuration->GetSprite(element, "MaterialsButton")) )
				mMaterialsButton->resize(mSwitchSize, mSwitchSize);
			
			if ( (mItemsButton = Configuration->GetSprite(element, "ItemsButton")) )
				mItemsButton->resize(mSwitchSize, mSwitchSize);
			
			LoadXmlValue(element, "ButtonSize", mSwitchSize);
			LoadXmlValue(element, "ButtonX", mSwitchX);
			LoadXmlValue(element, "ButtonY", mSwitchY);
			
			LoadXmlValue(element, "AmountX", mAmountX);
			LoadXmlValue(element, "AmountY", mAmountY);
		}
		
		mMaterialsPerPage = static_cast<unsigned int>( (mHeight - mListX - mArrowSize) / (Prgm::Materials::IconSize + mLineSpacing) );
		mItemsPerPage = static_cast<unsigned int>( (mHeight - mListX - mArrowSize) / (Prgm::Items::IconSize + mLineSpacing) );
	}
	
	void WinBuilding::Load(const TiXmlElement* const element)
	{
		this->Window::Load(element);
		
		LoadXmlValue(element, "FabriqueItems", mItems);
		LoadXmlValue(element, "QuantiteATransformer", mAmount);
		LoadXmlValue(element, "PageIssue", mPageIssue);
	}
	
	TiXmlElement* WinBuilding::Save() const
	{
		TiXmlElement* saveNode = new TiXmlElement("WinBuilding");
		this->Window::Save(saveNode);
		
		AddProperty(saveNode, "FabriqueItems", mItems);
		AddProperty(saveNode, "QuantiteATransformer", mAmount);
		AddProperty(saveNode, "PageIssue", mPageIssue);
		
		return saveNode;
	}
	
	void WinBuilding::draw(sf::RenderTarget* const render)
	{
		Window::draw(render);
		
		if (mVisible && mBuilding && mBuilding->IsBuilt())
		{
			legacy::Text* tempString = new legacy::Text();
			
			// building type
			tempString->setText(mType->GetName());
			tempString->setSize(mNameSize);
			tempString->setPosition(mX + mWidth / 2 - tempString->getRect().width / 2, mY + mNameY);
			render->draw(*tempString);
			
			unsigned int i = 0;
			std::ostringstream progressText;
			double progress;
			if (mItems)
			{
				Types::Item* typeItem;
				unsigned int itemCount = mType->GetAvailableItemCount();
				while (i < mItemsPerPage && i + mPageIssue * mItemsPerPage < itemCount)
				{
					typeItem = mType->GetItemFromList(i + mPageIssue * mItemsPerPage);
					
					// item type
					tempString->setText(typeItem->GetName());
					tempString->setSize(mTextSize);
					if (tempString->getRect().width > mNamesWidth)
						tempString->setScaleX(mNamesWidth / tempString->getRect().width);
					tempString->setPosition(mX + mListX, mY + (Prgm::Items::IconSize + mLineSpacing) * i + mListY);
					render->draw(*tempString);
					
					typeItem->DrawIcon(render, mX + mListX + mNamesWidth, mY + (Prgm::Items::IconSize + mLineSpacing) * i + mListY);
					
					// progress (and count)
					progressText.clear();
					progressText.str("");
					progressText << "x" << mBuilding->GetStockItemCount(i + mPageIssue * mItemsPerPage);
					
					if ( (progress = mBuilding->GetItemManufacturingProgress(i + mPageIssue * mItemsPerPage)) )
						progressText << " (" << floor(progress * 100) << "%)";
					
					tempString->setText(progressText.str());
					tempString->setSize(mTextSize);
					tempString->setPosition(mX + mListX + mNamesWidth + Prgm::Items::IconSize, mY + (Prgm::Items::IconSize + mLineSpacing) * i + mListY);
					render->draw(*tempString);
					
					i++;
				}
			}
			else
			{
				std::ostringstream text;
				text << "Quantité: " << mAmount;
				
				Text* amountText = new Text();
				amountText->setText(text.str());
				amountText->setPosition(mX + mAmountX, mY + mAmountY);
				amountText->draw(render);
				
				Types::Material* typeMaterial;
				unsigned int materialCount = mType->GetAvailableMaterialCount();
				
				while (i < mMaterialsPerPage && i + mPageIssue * mMaterialsPerPage < materialCount)
				{
					typeMaterial = Configuration->GetTypeMaterial(mType->GetMaterialIdFromList(i + mPageIssue * mMaterialsPerPage));
					
					tempString->setText(typeMaterial->GetName());
					tempString->setSize(mTextSize);
					
					if (tempString->getRect().width > mNamesWidth)
						tempString->setScaleX(mNamesWidth / tempString->getRect().width);
					tempString->setPosition(mX + mListX, mY + (Prgm::Materials::IconSize + mLineSpacing) * i + mListY);
					render->draw(*tempString);
					
					typeMaterial->DrawIcon(render, mX + mListX + mNamesWidth, mY + (Prgm::Materials::IconSize + mLineSpacing) * i + mListY);
					
					i++;
				}
			}
			
			GameSystem::WorkElement* element = mBuilding->GetWorkList();
			double maxItems = (mWidth - 2 * (mMargin + mArrowSize)) / Prgm::Items::IconSize - 1;
			
			i = 0;
			while (element && i < maxItems)
			{
				if (element->material)
					Configuration->GetTypeMaterial(mType->GetMaterialIdFromList(element->id))->DrawIcon(render, mX + mMargin + mArrowSize + i * Prgm::Materials::IconSize, mY + mHeight - Prgm::Materials::IconSize - mMargin);
				else
					mType->GetItemFromList(element->id)->DrawIcon(render, mX + mMargin + mArrowSize + i * Prgm::Items::IconSize, mY + mHeight - Prgm::Items::IconSize - mMargin);
				
				element = element->nextElement;
				i++;
			}
			
			if (mItems)
			{
				if (mMaterialsButton && mMaterialPageCount > 0)
				{
					mMaterialsButton->setPosition(mX + mSwitchX, mY + mSwitchY);
					render->draw(*mMaterialsButton);
				}
			}
			else
			{
				if (mItemsButton && mItemPageCount > 0)
				{
					mItemsButton->setPosition(mX + mSwitchX, mY + mSwitchY);
					render->draw(*mItemsButton);
				}
			}
			
			delete tempString;
		}
	}
	
	bool WinBuilding::Catch(const int& x, const int& y, const sf::Event& Event, InterfaceModule::DragNDrop* const drag)
	{
		if (mVisible
			&& mX <= x && x <= mX + mWidth
			&& mY <= y && y <= mY + mHeight)
		{/* TODO
			if (Event.type == sf::Event::MouseButtonReleased && Event.mouseButton.button == sf::Mouse::Left && drag->Item)
			{
				if (!mBuilding->AddItemToStock(drag->Item))
					(*drag->RetourItem) = drag->Item;
				
				drag->object = NULL;
				
				return true;
			}
			else */if (!mItems && Event.type == sf::Event::MouseWheelMoved)
			{
				mAmount += Event.mouseWheel.delta;
				
				if (mAmount < 0)
					mAmount = 0;
				
				return true;
			}
			else if (mX + mSwitchX <= x && x <= mX + mSwitchX + mSwitchSize &&
				mY + mSwitchY <= y && y <= mY + mSwitchY + mSwitchSize &&
				mMaterialPageCount > 0 && mItemPageCount > 0)
			{
				if (Event.type == sf::Event::MouseButtonReleased)
				{
					mItems = !mItems;
					mPageAmount = (mItems ? mItemPageCount : mMaterialPageCount);
				}
				
				return true;
			}
			else if (mBuilding)
			{
				unsigned int i = 0;
				
				if (mItems)
				{
					unsigned int itemCount = mType->GetAvailableItemCount();
					while (i < mItemsPerPage && i + mPageIssue * mItemsPerPage < itemCount)
					{
						if (mType->GetItemFromList(i + mPageIssue * mItemsPerPage)->Catch(x, y, mX + mListX + mNamesWidth, mY + (Prgm::Items::IconSize + mLineSpacing) * i + mListY))
						{
							if (Event.type == sf::Event::MouseButtonReleased && Event.mouseButton.button == sf::Mouse::Left)
							{
								mBuilding->AddItemToWorkList(i + mPageIssue * mItemsPerPage);
							}
							else if (Event.type == sf::Event::MouseButtonReleased && Event.mouseButton.button == sf::Mouse::Right)
							{
								GameSystem::Item* tookItem = mBuilding->RemoveFromStock(i + mPageIssue * mItemsPerPage);
								
								if (tookItem)
									mPlayer->AddItemToStock(tookItem);
							}
							
							return true;
						}
						
						i++;
					}
				}
				else
				{
					unsigned int materialCount = mType->GetAvailableMaterialCount();
					while (i < mMaterialsPerPage && i + mPageIssue * mMaterialsPerPage < materialCount)
					{
						if (Configuration->GetTypeMaterial(mType->GetMaterialIdFromList(i + mPageIssue * mItemsPerPage))->Catch(x, y, mX + mListX + mNamesWidth, mY + (Prgm::Items::IconSize + mLineSpacing) * i + mListY))
						{
							if (Event.type == sf::Event::MouseButtonReleased && Event.mouseButton.button == sf::Mouse::Left && mAmount)
							{
								mBuilding->AddMaterialToWorlList(i + mPageIssue * mMaterialsPerPage, mAmount);
								mAmount = 0;
							}
							
							return true;
						}
						
						i++;
					}
				}
				
				GameSystem::WorkElement* element = mBuilding->GetWorkList();
				double shownElements = (mWidth - 2 * (mMargin + mArrowSize)) / Prgm::Items::IconSize - 1;
				
				i = 0;
				while (element && i < shownElements)
				{
					if (mX + mMargin + mArrowSize + i * Prgm::Items::IconSize <= x && x <= mX + mMargin + mArrowSize + i * Prgm::Items::IconSize + Prgm::Items::IconSize
						&& mY + mHeight - Prgm::Items::IconSize - mMargin <= y && y <= mY + mHeight - Prgm::Items::IconSize - mMargin + Prgm::Items::IconSize)
					{
						if (Event.type == sf::Event::MouseButtonReleased)
							mBuilding->RemoveFromWorkList(i);
						
						return true;
					}
					
					element = element->nextElement;
					i++;
				}
			}
		}
		
		return Window::Catch(x, y, Event, drag);
	}
	
	std::string WinBuilding::getText(const int& x, const int& y, const GameSystem::Character* const personnage) const
	{
		if (mVisible && mBuilding)
		{
			unsigned int i = 0;
			unsigned int itemCount = mType->GetAvailableItemCount();
			while (i < mItemsPerPage && i + mPageIssue * mItemsPerPage < itemCount)
			{
				if (mType->GetItemFromList(i + mPageIssue * mItemsPerPage)->Catch(x, y, mX + mListX + mNamesWidth, mY + (Prgm::Items::IconSize + mLineSpacing) * i + mListY))
				{
					std::ostringstream text;
					Types::Item* typeItem = mType->GetItemFromList(i + mPageIssue * mItemsPerPage);
					
					text << typeItem->GetName();
					text << "\nRessources necessaires:";
					
					unsigned int j = 0;
					unsigned int materialCount = typeItem->GetRequisiteMaterialCount();
					while (j < materialCount)
					{
						text << "\n" << NormalizeMaterial( personnage->GetMaterialStockAmount(typeItem->GetRequisiteMaterialId(j)) ) << "/" << typeItem->GetRequisiteMaterialAmount(j) << " " << Configuration->GetTypeMaterial(typeItem->GetRequisiteMaterialId(j))->GetName();
						j++;
					}
					
					return text.str();
				}
				
				i++;
			}
			
			GameSystem::WorkElement* element = mBuilding->GetWorkList();
			double shownElements = (mWidth - 2 * (mMargin + mArrowSize)) / Prgm::Items::IconSize - 1;
			i = 0;
			while (element && i < shownElements)
			{
				if (mX + mMargin + mArrowSize + i * Prgm::Items::IconSize <= x && x <= mX + mMargin + mArrowSize + i * Prgm::Items::IconSize + Prgm::Items::IconSize
					&& mY + mHeight - Prgm::Items::IconSize - mMargin <= y && y <= mY + mHeight - Prgm::Items::IconSize - mMargin + Prgm::Items::IconSize)
				{
					std::ostringstream text;
					
					if (element->material)
					{
						Types::Material* typeMaterial = Configuration->GetTypeMaterial(mType->GetMaterialIdFromList(element->id));
						
						text << typeMaterial->GetName() << " (" << floor(10 * element->amount) / 10 << ")";
						text << "\nRessources necessaires:";
						
						unsigned int i = 0;
						unsigned int materialCount = typeMaterial->GetRequisiteMaterialCount();
						while (i < materialCount)
						{
							double requisiteAmount = element->amount * typeMaterial->GetRequisiteMaterialAmount(i);
							if (requisiteAmount)
								text << "\n" << NormalizeMaterial(requisiteAmount) << " " << Configuration->GetTypeMaterial(typeMaterial->GetRequisiteMaterialId(i))->GetName();
							
							i++;
						}
					}
					else
					{
						Types::Item* typeItem = mType->GetItemFromList(element->id);
						double avancement = mBuilding->GetItemManufacturingProgress(element->id);
						
						text << typeItem->GetName() << " (" << floor(100 * avancement) << "%)";
						text << "\nRessources necessaires:";
						
						unsigned int i = 0;
						unsigned int materialCount = typeItem->GetRequisiteMaterialCount();
						double amount;
						while (i < materialCount)
						{
							amount = (1 - avancement) * typeItem->GetRequisiteMaterialAmount(i);
							if (amount)
								text << "\n" << NormalizeMaterial(amount) << "/" << typeItem->GetRequisiteMaterialAmount(i) << " " << Configuration->GetTypeMaterial(typeItem->GetRequisiteMaterialId(i))->GetName();
							
							i++;
						}
					}
					
					return text.str();
				}
				
				element = element->nextElement;
				i++;
			}
		}
		
		return Window::getText(x, y);
	}
	
	void WinBuilding::LoadBuilding(GameSystem::Building* const building)
	{
		if ( (mBuilding != building) && (mBuilding = building) )
		{
			mType = mBuilding->GetType();
			
			mMaterialPageCount = mType->GetAvailableMaterialCount() / mMaterialsPerPage;
			if (mMaterialPageCount * mMaterialsPerPage < mType->GetAvailableMaterialCount())
				mMaterialPageCount++;
			
			mItemPageCount = mType->GetAvailableItemCount() / mItemsPerPage;
			if (mItemPageCount * mItemsPerPage < mType->GetAvailableItemCount())
				mItemPageCount++;
			
			mItems = (mMaterialPageCount == 0);
			mPageAmount = (mItems ? mItemPageCount : mMaterialPageCount);
		}
	}
	
	void WinBuilding::SetPlayer(GameSystem::Character* const player)
	{
		mPlayer = player;
	}
}
