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
#include "Character.hpp"

#include <iostream>
#include <sstream>

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "../data.hpp"
#include "../functions.hpp"
#include "TypeItemCategory.hpp"
#include "World.hpp"

namespace GameSystem
{
	Character::Character() : GameSystem::Object(),
		mVisible(true),
		mRace(Configuration->GetRace(1)), // TODO
		mNation(0),
		mSpriteWidth(24),
		mSpriteHeight(32),
		mSpeed(7),
		mName("Yoha"),
		mDirection(2),
		mStep(1),
		mActionList(NULL),
		mCurrentBuilding(NULL),
		mText(new InterfaceModule::Text()),
		mImage(NULL),
		mSprite(NULL)
	{
		mClass = GameSystem::Classes::Character;
		
		mImage = mRace->getImage();
		mWidth = mRace->getWidth();
		mHeight = mRace->getHeight();
		mSpeed = mRace->GetSpeed();
	}
	
	Character::~Character()
	{
		//delete mSprite;
		//delete mText;
	}
	
	void Character::Load()
	{
		if (mXml)
		{
			LoadXmlValue(mXml, "Nation", mNation);
			LoadXmlValue(mXml, "Name", mName);
			LoadXmlValue(mXml, "X", mX);
			LoadXmlValue(mXml, "Y", mY);
			LoadXmlValue(mXml, "Direction", mDirection);
			LoadXmlValue(mXml, "Pas", mStep);
			LoadXmlValue(mXml, "Visible", mVisible);
			
			unsigned int currentBuildingId;
			if (LoadXmlValue(mXml, "CurrentBuilding", currentBuildingId))
			{
				GameSystem::Object* object = mWorld->GetObject(currentBuildingId);
				if (object && object->GetClass() == GameSystem::Classes::Building)
					mCurrentBuilding = static_cast<GameSystem::Building*>(object);
			}
			
			const TiXmlElement* actionXml = mXml->FirstChildElement("ActionList");
			if (actionXml && (actionXml = actionXml->FirstChildElement()))
			{
				while (actionXml)
				{
					std::string type;
					LoadXmlValue(actionXml, "type", type);
					if (type == "Wait")
					{
						double duration;
						LoadXmlValue(actionXml, "duration", duration);
						this->Wait(duration);
					}
					else if (type == "Say")
					{
						/* TODO */
					}
					else if (type == "GoCoords")
					{
						double x;
						double y;
						LoadXmlValue(actionXml, "x", x);
						LoadXmlValue(actionXml, "y", y);
						this->GoCoords(x, y);
					}
					else if (type == "GoObject")
					{
						unsigned int target;
						LoadXmlValue(actionXml, "target", target);
						this->GoObject(mWorld->GetObject(target));
					}
					else if (type == "Enter")
					{
						unsigned int building;
						LoadXmlValue(actionXml, "building", building);
						GameSystem::Object* object = mWorld->GetObject(building);
						if (object && object->GetClass() == GameSystem::Classes::Building)
							this->Enter(static_cast<GameSystem::Building*>(object));
					}
					else if (type == "Work")
					{
						unsigned int target;
						double amount = 0;
						LoadXmlValue(actionXml, "target", target);
						LoadXmlValue(actionXml, "amount", amount);
						this->Work(mWorld->GetObject(target), amount);
					}
					else if (type == "Attack")
					{
						unsigned int target;
						LoadXmlValue(actionXml, "target", target);
						this->Attack(mWorld->GetObject(target));
					}
					
					actionXml = actionXml->NextSiblingElement();
				}
			}
			
			LoadCompactFromXml(mXml, "Materials", mMaterialAmounts);
			LoadCompactFromXml(mXml, "Inventory", mInventory, mWorld);
			LoadCompactFromXml(mXml, "Equipment", mEquipment, mWorld);
		}
		else
			mNation = mWorld->AssignNation(this);
		
		//mSkillLevels[GameSystem::Skills::Velocity] = 100;
		
		if (mImage)
		{
			mSprite = new legacy::Sprite();
			mSprite->setImage(*mImage);
		}
	}
	
	TiXmlElement* Character::Save() const
	{
		TiXmlElement* saveNode = new TiXmlElement("Character");
		
		saveNode->SetAttribute("id", mId);
		
		AddProperty(saveNode, "Nation", mNation);
		AddProperty(saveNode, "Name", mName);
		AddProperty(saveNode, "Direction", mDirection);
		AddProperty(saveNode, "X", mX);
		AddProperty(saveNode, "Y", mY);
		AddProperty(saveNode, "Pas", mStep);
		AddProperty(saveNode, "Visible", mVisible);
		
		if (mActionList)
		{
			TiXmlElement* selement = new TiXmlElement("ActionList");
			saveNode->LinkEndChild(selement);
			
			GameSystem::ActionElement* action = mActionList;
			while (action)
			{
				TiXmlElement* sselement = new TiXmlElement("Action");
				selement->LinkEndChild(sselement);
				if (action->type == GameSystem::TypesActions::Wait)
				{
					sselement->SetAttribute("duration", "Wait");
					sselement->SetDoubleAttribute("duration", *static_cast<double*>(action->params[0]));
				}
				else if (action->type == GameSystem::TypesActions::Say)
				{
					/* TODO
					sselement->SetAttribute("type", "Say");
					sselement->SetAttribute("text", *static_cast<sf::String*>(action->params[0]));
					*/
				}
				else if (action->type == GameSystem::TypesActions::GoCoords)
				{
					sselement->SetAttribute("type", "GoCoords");
					sselement->SetDoubleAttribute("x", *static_cast<double*>(action->params[0]));
					sselement->SetDoubleAttribute("y", *static_cast<double*>(action->params[1]));
				}
				else if (action->type == GameSystem::TypesActions::GoObject)
				{
					sselement->SetAttribute("type", "GoObject");
					sselement->SetAttribute("target", static_cast<GameSystem::Object*>(action->params[0])->GetId());
				}
				else if (action->type == GameSystem::TypesActions::Enter)
				{
					sselement->SetAttribute("type", "Enter");
					GameSystem::Building* building = static_cast<GameSystem::Building*>(action->params[0]);
					if (building)
						sselement->SetAttribute("building", building->GetId());
				}
				else if (action->type == GameSystem::TypesActions::Work)
				{
					sselement->SetAttribute("type", "Work");
					GameSystem::Object* target = static_cast<GameSystem::Object*>(action->params[0]);
					sselement->SetAttribute("target", target->GetId());
					double* amount = static_cast<double*>(action->params[1]);
					if (amount)
						sselement->SetDoubleAttribute("amount", *amount);
				}
				else if (action->type == GameSystem::TypesActions::Attack)
				{
					sselement->SetAttribute("type", "Attack");
					GameSystem::Object* target = static_cast<GameSystem::Object*>(action->params[0]);
					sselement->SetAttribute("target", target->GetId());
				}
				
				action = action->nextAction;
			}
		}
		
		if (mCurrentBuilding)
			AddProperty(saveNode, "CurrentBuilding", mCurrentBuilding->GetId());
		
		SaveCompactToXml(saveNode, "Materials", mMaterialAmounts);
		SaveCompactToXml(saveNode, "Inventory", mInventory);
		SaveCompactToXml(saveNode, "Equipment", mEquipment);
		
		return saveNode;
	}
	
	bool Character::Catch(const double& x, const double& y) const
	{
		unsigned int X = this->GetXSprite();
		unsigned int Y = this->GetYSprite();
		return	X <= x && x <= X + mWidth &&
			Y <= y && y <= Y + mHeight;
	}
	
	void Character::ReceiveString(const sf::String& txt)
	{
		mText->setText(txt);
	}
	
	void Character::Process(const double& time)
	{
		this->onWhenever();
		
		unsigned int i = 0;
		while (i < mSkillLevels.size())
		{
			double bonus = 0;
			unsigned int j = 0;
			while (j < mEquipment.size())
			{
				if (mEquipment[j])
				{
					Types::Item* typeItem = mEquipment[j]->GetType();
					Types::Skill* skill = typeItem->GetSkill();
					bonus += typeItem->GetBonus(i) * skill->GetRatio(mSkillLevels[skill->GetId()]);
				}
				j++;
			}
			
			mAlteredSkills[i] = mSkillLevels[i] + bonus;
			i++;
		}
		
		if (mActionList)
		{
			if (mActionList->type == GameSystem::TypesActions::Wait)
			{
				double* duration = static_cast<double*>(mActionList->params[0]);
				*duration -= time;
			}
			else if (mActionList->type == GameSystem::TypesActions::Say)
			{
				sf::String text = *static_cast<sf::String*>(mActionList->params[0]);
				this->ReceiveString(text);
				this->NextAction();
			}
			else if (mActionList->type == GameSystem::TypesActions::GoCoords || mActionList->type == GameSystem::TypesActions::GoObject)
			{
				mVisible = true;
				mCurrentBuilding = NULL;
				
				double x;
				double y;
				if (mActionList->type == GameSystem::TypesActions::GoObject)
				{
					GameSystem::Object* target = static_cast<GameSystem::Object*>(mActionList->params[0]);
					x = target->GetX();
					y = target->GetY();
				}
				else
				{
					x = *static_cast<double*>(mActionList->params[0]);
					y = *static_cast<double*>(mActionList->params[1]);
				}
				
				//double traveledDistance = mAlteredSkills[GameSystem::Skills::Velocity] * time;
				double traveledDistance = 100 * time * Configuration->GetTypeField(mWorld->GetIdField(mX + mWidth / 2, mY + mHeight))->GetSpeedRatio();
				double targetDistance = Range(mX, mY, x, y);
				double targetX = x - mX;
				double targetY = y - mY;
				
				if (traveledDistance > targetDistance)
				{
					mX = x;
					mY = y;
					this->NextAction();
				}
				else
				{
					mX += targetX * traveledDistance / targetDistance;
					mY += targetY * traveledDistance / targetDistance;
				}
				
				if (Absolute(targetX) < Absolute(targetY)) // vertical
				{
					mDirection = targetY < 0 ? 0 : 2;
				}
				else
				{
					mDirection = targetX < 0 ? 3 : 1;
				}
				
				mStep += mSpeed * time;
				if (mStep >= 4)
					mStep = 0;
			}
			else if (mActionList->type == GameSystem::TypesActions::Enter)
			{
				GameSystem::Building* building = static_cast<GameSystem::Building*>(mActionList->params[0]);
				
				if (Range(building, this) < 1)
				{
					mCurrentBuilding = building;
					mVisible = false;
				}
				
				this->NextAction();
			}
			else if (mActionList->type == GameSystem::TypesActions::Work)
			{
				GameSystem::Object* target = static_cast<GameSystem::Object*>(mActionList->params[0]);
				double* amount = static_cast<double*>(mActionList->params[1]);
				
				double doneWork = 0;
				if (target->GetClass() == GameSystem::Classes::Building)
					doneWork = static_cast<GameSystem::Building*>(target)->Work(this, time);
				else if (target->GetClass() == GameSystem::Classes::Ressource)
					doneWork = static_cast<GameSystem::Ressource*>(target)->Harvest(this, time);
				
				if (amount)
				{
					*amount -= doneWork;
					if (*amount <= 0)
						this->NextAction();
				}
			}
			else if (mActionList->type == GameSystem::TypesActions::Attack)
			{
				GameSystem::Object* target = static_cast<GameSystem::Object*>(mActionList->params[0]);
				
				if (Range(this, target) < 100) // TODO (traveledDistance)
				{
					target->PerdHP(5 * time);
					
					if (target->GetClass() == GameSystem::Classes::Character)
						static_cast<GameSystem::Character*>(target)->onAttacked(this);
				}
				else
					this->NextAction();
			}
		}
	}
	
	void Character::draw(sf::RenderTarget* const render)
	{
		if (mVisible && mSprite)
		{
			double x = this->GetXSprite();
			double y = this->GetYSprite();
			
			int stepToDraw = floor(mStep);
			if (stepToDraw == 3)
				stepToDraw = 1;
			
			mSprite->setSubRect(sf::IntRect(mSpriteWidth * stepToDraw,		mSpriteHeight * mDirection,
							mSpriteWidth * (stepToDraw + 1),	mSpriteHeight * (mDirection + 1)));
			mSprite->resize(mWidth, mHeight);
			mSprite->setPosition(x, y);
			render->draw(*mSprite);
			
			DrawGauge(render, x, y + mHeight + 2, mWidth, 5, mHP / mMaxHP);
		}
		
		mText->setPosition(mX + mWidth / 2 - mText->getRect().width / 2, mY + Prgm::Fields::Height - mHeight - mText->getRect().height);
		mText->draw(render);
	}
	
	GameSystem::Ressource* Character::GetRessource(const unsigned int& id)
	{
		return mWorld->FindRessource(this, Configuration->GetTypeRessource(id));
	}
	
	double Character::GetXSprite() const
	{
		return mX + Prgm::Fields::Width / 2 - mWidth / 2;
	}
	
	double Character::GetYSprite() const
	{
		return mY + Prgm::Fields::Height - mHeight;
	}
	
	unsigned int Character::GetNation() const
	{
		return mNation;
	}
	
	void Character::AddAction(GameSystem::ActionElement* action)
	{
		action->nextAction = NULL;
		
		if (mActionList)
		{
			GameSystem::ActionElement* lastAction = mActionList;
			while (lastAction->nextAction)
				lastAction = lastAction->nextAction;
			
			lastAction->nextAction = action;
		}
		else
			mActionList = action;
	}
	
	void Character::NextAction()
	{
		if (mActionList)
			mActionList = mActionList->nextAction;
	}
	
	void Character::ResetActions()
	{
		mActionList = NULL;
	}
	
	void Character::SetSay(const char* txt)
	{
		this->ReceiveString(sf::String(txt));
	}
	
	void Character::Wait(double time)
	{
		GameSystem::ActionElement* action = new GameSystem::ActionElement();
		action->type = GameSystem::TypesActions::Wait;
		action->params = new void*[1];
		action->params[0] = new double(time);
		
		this->AddAction(action);
	}
	
	void Character::Say(const char* txt)
	{
		GameSystem::ActionElement* action = new GameSystem::ActionElement();
		action->type = GameSystem::TypesActions::Say;
		action->params = new void*[1];
		action->params[0] = static_cast<void*>(new sf::String(txt));
		
		this->AddAction(action);
	}
	
	void Character::GoCoords(const double& x, const double& y)
	{
		GameSystem::ActionElement* action = new GameSystem::ActionElement();
		action->type = GameSystem::TypesActions::GoCoords;
		action->params = new void*[2];
		action->params[0] = new double(x);
		action->params[1] = new double(y);
		
		this->AddAction(action);
		
		this->onGoCoords(x, y);
	}
	
	void Character::GoObject(GameSystem::Object* const target)
	{
		if (target && target != this)
		{
			GameSystem::ActionElement* action = new GameSystem::ActionElement();
			action->type = GameSystem::TypesActions::GoObject;
			action->params = new void*[1];
			action->params[0] = static_cast<void*>(target);
			
			this->AddAction(action);
			
			this->onGoObject(target);
		}
	}
	
	void Character::Enter(GameSystem::Building* const building)
	{
		if (building)
		{
			GameSystem::ActionElement* action = new GameSystem::ActionElement();
			action->type = GameSystem::TypesActions::Enter;
			action->params = new void*[1];
			action->params[0] = static_cast<void*>(building);
			
			this->AddAction(action);
		}
	}
	
	void Character::Work(GameSystem::Object* const target, const double& amount)
	{
		if (target)
		{
			GameSystem::ActionElement* action = new GameSystem::ActionElement();
			action->type = GameSystem::TypesActions::Work;
			action->params = new void*[2];
			action->params[0] = static_cast<void*>(target);
			if (amount)
				action->params[1] = new double(amount);
			else
				action->params[1] = NULL;
			
			this->AddAction(action);
		}
	}
	
	void Character::Attack(GameSystem::Object* const target)
	{
		GameSystem::ActionElement* action = new GameSystem::ActionElement();
		action->type = GameSystem::TypesActions::Attack;
		action->params = new void*[1];
		action->params[0] = static_cast<void*>(target);
		
		this->AddAction(action);
	}
	
	GameSystem::Building* Character::SeatBuilding(const unsigned int& id)
	{
		Types::Building* typeBuilding = Configuration->GetTypeBuilding(id);
		double maxX = mWorld->getWidth();
		double maxY = mWorld->getHeight();
		
		unsigned int i = 0;
		while (i < 10)
		{
			GameSystem::Building* building;
			if ( (building = mWorld->SeatBuilding(legacy::random(0.f, maxX), legacy::random(0.f, maxY), typeBuilding)) )
				return building;
			
			i++;
		}
		
		return NULL;
	}
	
	bool Character::isActive() const
	{
		return (mActionList != NULL);
	}
	
	bool Character::isGoing(GameSystem::Object* const target) const
	{
		return (mActionList && mActionList->type == GameSystem::TypesActions::GoObject) && (mActionList->params[0] == target);
	}
	
	bool Character::isAttacking(GameSystem::Object* const target) const
	{
		return (mActionList && mActionList->type == GameSystem::TypesActions::Attack) && (mActionList->params[0] == target);
	}
	
	GameSystem::Object* Character::GetTarget() const
	{
		return (mActionList && mActionList->type == GameSystem::TypesActions::Attack) ? static_cast<GameSystem::Object*>(mActionList->params[0]) : NULL;
	}
	
	Types::Race* Character::GetRace() const
	{
		return mRace;
	}
	
	std::string Character::GetName() const
	{
		return mName;
	}
	
	GameSystem::Building* Character::GetCurrentBuilding() const
	{
		return mCurrentBuilding;
	}
	
	std::string Character::GetTextTypeBuilding(const Types::Building* const typeBuilding) const
	{
		if (typeBuilding)
		{
			std::ostringstream text;
			
			text << typeBuilding->GetName();
			text << "\n" << typeBuilding->GetDescription();
			
			unsigned int length = typeBuilding->GetRequisiteMaterialCount();
			unsigned int i = 0;
			while (i < length)
			{
				unsigned int materialId = typeBuilding->GetRequisiteMaterialId(i);
				double amount = mMaterialAmounts[materialId];
				text << "\n" << NormalizeMaterial(amount) << "/" << typeBuilding->GetRequisiteMaterialAmount(i) << " " << Configuration->GetTypeMaterial(materialId)->GetName();
				i++;
			}
			
			return text.str();
		}
		
		return "";
	}
	
	std::string Character::GetTextObject(GameSystem::Object* const object) const
	{
		if (object)
		{
			if (object->GetClass() == GameSystem::Classes::Character)
			{
				GameSystem::Character* character = static_cast<GameSystem::Character*>(object);
				if (character == this)
					return "Moi";
				else if (!mNation || character->GetNation() != mNation)
					return "Ennemi";
				else
					return character->GetName();
			}
			else if (object->GetClass() == GameSystem::Classes::Building) // TODO (et re-do)
			{
				GameSystem::Building* building = static_cast<GameSystem::Building*>(object);
				Types::Building* typeBuilding = building->GetType();
				
				std::ostringstream text;
				
				text << typeBuilding->GetName();
				
				if (!building->IsBuilt())
					text << " (" << floor(100 * building->GetHP() / typeBuilding->GetMaxHP()) << "%)";
				
				if (!building->IsBuilt())
				{
					text << "\nMaterials nécessaires:";
					
					unsigned int length = typeBuilding->GetRequisiteMaterialCount();
					unsigned int i = 0;
					double requisiteAmount;
					while (i < length)
					{
						requisiteAmount = typeBuilding->GetRequisiteMaterialAmount(i) - building->GetRequisiteMaterialAmount(i);
						
						if (requisiteAmount)
							text << "\n" << NormalizeMaterial(requisiteAmount) << " " << Configuration->GetTypeMaterial(typeBuilding->GetRequisiteMaterialId(i))->GetName();
						
						i++;
					}
				}
				
				return text.str();
			}
			else if (object->GetClass() == GameSystem::Classes::Ressource)
			{
				Types::Material* typeMaterial = static_cast<GameSystem::Ressource*>(object)->GetTypeRessource()->GetTypeMaterial();
				unsigned int materialId = typeMaterial->GetId();
				
				std::ostringstream text;
				text << typeMaterial->GetName() << "(";
				text << ( materialId < mMaterialAmounts.size() ? NormalizeMaterial( mMaterialAmounts[materialId] ) : 0 );
				text << ")";
				return text.str();
			}
			else if (object->GetClass() == GameSystem::Classes::Item)
			{
				return static_cast<GameSystem::Item*>(object)->GetType()->GetName();
			}
		}
		
		return "";
	}
	
	double Character::GetMaterialStockAmount(const unsigned int& materialId) const
	{
		if (materialId < mMaterialAmounts.size())
			return mMaterialAmounts[materialId];
		else
			return 0;
	}
	
	double Character::AddMaterial(const unsigned int& materialId, const double& amount)
	{
		double max = 20 - mMaterialAmounts[materialId];
		double addedAmount = amount < max ? amount : max;
		mMaterialAmounts[materialId] += addedAmount;
		
		return addedAmount;
	}
	
	double Character::RemoveMaterial(const unsigned int& materialId, const double& amount)
	{
		double removedAmount = amount < mMaterialAmounts[materialId] ? amount : mMaterialAmounts[materialId];
		mMaterialAmounts[materialId] -= removedAmount;
		
		return removedAmount;
	}
	
	Vector <GameSystem::Item*>* Character::GetInventory()
	{
		return &mInventory;
	}
	
	void Character::AddItemToStock(GameSystem::Item* const item)
	{
		unsigned int i = 0;
		while (mInventory[i])
			i++;
		mInventory[i] = item;
	}
	
	Vector <GameSystem::Item*>* Character::GetItemsEquipes()
	{
		return &mEquipment;
	}
}
