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
#include "Building.hpp"

#include <sstream>
#include <cmath>

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "TypeBuilding.hpp"
#include "World.hpp"
#include "Character.hpp"

namespace GameSystem
{
	Building::Building(Types::Building* const type) : GameSystem::Object(),
		mType(type),
		mBuilt(false),
		mWorkList(NULL)
	{
		mHP = 0;
		mClass = GameSystem::Classes::Building;
		
		if (mType)
		{
			mWidth = mType->GetStructureWidth();
			mHeight = mType->GetStructureHeight();
		}
	}
	
	Building::~Building()
	{
		while (mWorkList)
		{
			GameSystem::WorkElement* nextElement = mWorkList->nextElement;
			delete mWorkList;
			mWorkList = nextElement;
		}
	}
	
	void Building::Load()
	{
		if (mXml)
		{
			LoadXmlValue(mXml, "x", mX);
			LoadXmlValue(mXml, "y", mY);
			
			unsigned int type = 0;
			LoadXmlValue(mXml, "Type", type);
			mType = Configuration->GetTypeBuilding(type);
			mWidth = mType->GetStructureWidth();
			mHeight = mType->GetStructureHeight();
			
			LoadXmlValue(mXml, "Built", mBuilt);
			LoadXmlValue(mXml, "HP", mHP);
			
			const TiXmlElement* requisiteMaterials = mXml->FirstChildElement("Requisite");
			if (requisiteMaterials)
			{
				unsigned int id;
				unsigned int requisiteMaterialCount = mType->GetRequisiteMaterialCount();
				
				requisiteMaterials = requisiteMaterials->FirstChildElement("Material");
				while (requisiteMaterials)
				{
					if (	LoadXmlValue(requisiteMaterials, "id", id) && id < requisiteMaterialCount )
						LoadXmlValue(requisiteMaterials, "amount", mRequisiteMaterial[id]);
					
					requisiteMaterials = requisiteMaterials->NextSiblingElement("Material");
				}
			}
			
			LoadCompactFromXml(mXml, "ItemMaking", mItemMaking);
			LoadCompactFromXml(mXml, "ItemStock", mItemStock);
			
			const TiXmlElement* listeElementXml = mXml->FirstChildElement("WorkList");
			if (listeElementXml)
			{
				const TiXmlElement* xmlElement = listeElementXml->FirstChildElement();
				GameSystem::WorkElement* currentElement = NULL;
				while (xmlElement)
				{
					GameSystem::WorkElement* element = new GameSystem::WorkElement();
					
					element->material = strcmp(xmlElement->Value(), "Item");
					LoadXmlValue(xmlElement, "id", element->id);
					if (element->material)
						LoadXmlValue(xmlElement, "amount", element->amount);
					element->nextElement = NULL;
					
					if (!mWorkList)
						mWorkList = element;
					else
						currentElement->nextElement = element;
					
					currentElement = element;
					xmlElement = xmlElement->NextSiblingElement();
				}
			}
		}
	}
	
	TiXmlElement* Building::Save() const
	{
		TiXmlElement* saveNode = new TiXmlElement("Building");
		
		saveNode->SetAttribute("id", mId);
		
		AddProperty(saveNode, "x", mX);
		AddProperty(saveNode, "y", mY);
		AddProperty(saveNode, "Type", mType->GetId());
		AddProperty(saveNode, "Built", mBuilt);
		AddProperty(saveNode, "HP", mHP);
		
		unsigned int count = mRequisiteMaterial.size();
		if (count > 0)
		{
			TiXmlElement* selement = new TiXmlElement("Requisite");
			saveNode->LinkEndChild(selement);
			unsigned int i = 0;
			while (i < count)
			{
				TiXmlElement* sselement = new TiXmlElement("Material");
				sselement->SetAttribute("id", i);
				sselement->SetDoubleAttribute("amount", mRequisiteMaterial[i]);
				selement->LinkEndChild(sselement);
			
				i++;
			}
		}
		
		SaveCompactToXml(saveNode, "ItemMaking", mItemMaking);
		SaveCompactToXml(saveNode, "ItemStock", mItemStock);
		
		if (mWorkList)
		{
			TiXmlElement* selement = new TiXmlElement("WorkList");
			saveNode->LinkEndChild(selement);
			
			GameSystem::WorkElement* element = mWorkList;
			while (element)
			{
				if (element->material)
				{
					TiXmlElement* sselement = new TiXmlElement("Material");
					selement->LinkEndChild(sselement);
					sselement->SetAttribute("id", element->id);
					sselement->SetDoubleAttribute("amount", element->amount);
				}
				else
				{
					TiXmlElement* sselement = new TiXmlElement("Item");
					selement->LinkEndChild(sselement);
					sselement->SetAttribute("id", element->id);
				}
				
				element = element->nextElement;
			}
		}
		
		return saveNode;
	}
	
	std::string Building::getText() const
	{
		std::ostringstream text;
		text << mType->GetName();
		
		if (!mBuilt)
		{
			text << 100 * mHP / mType->GetMaxHP() << "%";
			text << "\nManquent:";
			
			unsigned int length = mType->GetRequisiteMaterialCount();
			unsigned int i = 0;
			while (i < length)
			{
				text << "\n" << mType->GetRequisiteMaterialAmount(i) - floor(mRequisiteMaterial[i]) << " " << Configuration->GetTypeMaterial(mType->GetRequisiteMaterialId(i))->GetName();
				
				i++;
			}
		}
		
		return text.str();
	}
	
	void Building::draw(sf::RenderTarget* const render)
	{
		if (mType)
			mType->DrawStructure(render, mX, mY, mHP);
	}
	
	double Building::Work(GameSystem::Object* const worker_object, const double& time)
	{
		GameSystem::Character* const worker = static_cast<GameSystem::Character*>(worker_object);
		
		if (!mBuilt)
		{
			double HPBonus = 0;
			
			double maxHP = mType->GetMaxHP();
			double ratio = time * 100 / maxHP;
			if (ratio > 1 - mHP / maxHP)
				ratio = 1 - mHP / maxHP;
			
			unsigned int i = 0;
			unsigned int materialCount = mType->GetRequisiteMaterialCount();
			while (i < materialCount)
			{
				double maxRatio = worker->GetMaterialStockAmount(mType->GetRequisiteMaterialId(i)) / mType->GetRequisiteMaterialAmount(i);
				if (ratio > maxRatio)
					ratio = maxRatio;
				
				maxRatio = 1 - mRequisiteMaterial[i] / mType->GetRequisiteMaterialAmount(i);
				if (ratio > maxRatio)
					ratio = maxRatio;
				
				i++;
			}
			
			i = 0;
			while (i < materialCount)
			{
				double requisiteAmount = mType->GetRequisiteMaterialAmount(i);
				double broughtAmount = worker->RemoveMaterial(mType->GetRequisiteMaterialId(i), ratio * requisiteAmount);
				mRequisiteMaterial[i] += broughtAmount;
				HPBonus += mType->GetMaxHP() * broughtAmount / requisiteAmount / materialCount;
				
				i++;
			}
			
			mHP += HPBonus;
			if (mType->GetMaxHP() - mHP <= 0.01) // due to double precision
			{
				mBuilt = true;
				mHP = mType->GetMaxHP();
			}
			
			return HPBonus;
		}
		else if (mWorkList)
		{
			if (mWorkList->material)
			{
				Types::Material* typeMaterial = Configuration->GetTypeMaterial(mType->GetMaterialIdFromList(mWorkList->id));
				
				double ratio = time * typeMaterial->GetProductionRate();
				if (ratio > mWorkList->amount)
					ratio = mWorkList->amount;
				
				unsigned int i = 0;
				unsigned int materialCount = typeMaterial->GetRequisiteMaterialCount();
				while (i < materialCount)
				{
					double maxRatio = worker->GetMaterialStockAmount(typeMaterial->GetRequisiteMaterialId(i)) / typeMaterial->GetRequisiteMaterialAmount(i);
					if (ratio > maxRatio)
						ratio = maxRatio;
					
					i++;
				}
				
				i = 0;
				while (i < materialCount)
				{
					worker->RemoveMaterial(typeMaterial->GetRequisiteMaterialId(i), ratio * typeMaterial->GetRequisiteMaterialAmount(i));
					i++;
				}
				
				worker->AddMaterial(mType->GetMaterialIdFromList(mWorkList->id), ratio);
				mWorkList->amount -= ratio;
				
				if (mWorkList->amount <= 0)
					mWorkList = mWorkList->nextElement;
				
				return 0;
			}
			else
			{
				Types::Item* typeItem = mType->GetItemFromList(mWorkList->id);
				
				double ratio = time / typeItem->GetManufacturingDuration();
				if (ratio > 1 - mItemMaking[mWorkList->id])
					ratio = 1 - mItemMaking[mWorkList->id];
				
				unsigned int i = 0;
				unsigned int materialCount = typeItem->GetRequisiteMaterialCount();
				while (i < materialCount)
				{
					double maxRatio = worker->GetMaterialStockAmount(typeItem->GetRequisiteMaterialId(i)) / typeItem->GetRequisiteMaterialAmount(i);
					if (ratio > maxRatio)
						ratio = maxRatio;
					
					i++;
				}
				
				i = 0;
				while (i < materialCount)
				{
					worker->RemoveMaterial(typeItem->GetRequisiteMaterialId(i), ratio * typeItem->GetRequisiteMaterialAmount(i));
					
					i++;
				}
				
				mItemMaking[mWorkList->id] += ratio;
				
				if (mItemMaking[mWorkList->id] >= 1)
				{
					mItemMaking[mWorkList->id] = 0;
					mItemStock[mWorkList->id]++;
					mWorkList = mWorkList->nextElement;
				}
				
				return 0;
			}
		}
		
		return 0;
	}
	
	void Building::AddMaterialToWorlList(const unsigned int& materialId, const unsigned int& amount)
	{
		GameSystem::WorkElement* element;
		
		if (mWorkList)
		{
			element = mWorkList;
			
			while (element->nextElement)
			{
				element = element->nextElement;
			}
			
			element->nextElement = new GameSystem::WorkElement();
			element = element->nextElement;
		}
		else
		{
			mWorkList = new GameSystem::WorkElement();
			element = mWorkList;
		}
		
		element->id = materialId;
		element->material = true;
		element->amount = amount;
		element->nextElement = NULL;
	}
	
	void Building::AddItemToWorkList(const unsigned int& itemId)
	{
		GameSystem::WorkElement* element;
		
		if (mWorkList)
		{
			element = mWorkList;
		 	
			while (element->nextElement)
			{
				element = element->nextElement;
			}
			
			element->nextElement = new GameSystem::WorkElement();
			element = element->nextElement;
		}
		else
		{
			mWorkList = new GameSystem::WorkElement();
			element = mWorkList;
		}
		
		element->id = itemId;
		element->material = false;
		element->amount = 0;
		element->nextElement = NULL;
	}
	
	void Building::RemoveFromWorkList(const unsigned int& elementId)
	{
		if (elementId)
		{
			GameSystem::WorkElement* element = mWorkList;
			unsigned int i = 1;
			while (i < elementId && element)
			{
				element = element->nextElement;
				i++;
			}
			if (element && element->nextElement)
				element->nextElement = element->nextElement->nextElement;
		}
		else
			mWorkList = mWorkList->nextElement;
	}
	
	bool Building::AddItemToStock(GameSystem::Item* const item)
	{
		Types::Item* type = item->GetType();
		
		unsigned int i = 0;
		unsigned int availableItemCount = mType->GetAvailableItemCount();
		while (i < availableItemCount)
		{
			if (mType->GetItemFromList(i) == type)
			{
				mItemStock[i]++;				
				mWorld->RemoveObject(item->GetId());
				
				return true;
			}
			
			i++;
		}
		
		return false;
	}
	
	GameSystem::Item* Building::RemoveFromStock(const unsigned int& itemId)
	{
		if (mItemStock[itemId])
		{
			mItemStock[itemId]--;
			GameSystem::Item* item = mWorld->CreateItem(mType->GetItemFromList(itemId));
			return item;
		}
		else
			return NULL;
	}
	
	double Building::GetHP() const
	{
		return mHP;
	}
	
	Types::Building* Building::GetType() const
	{
		return mType;
	}
	
	bool Building::IsBuilt() const
	{
		return mBuilt;
	}
	
	double Building::GetRequisiteMaterialAmount(const unsigned int& materialId) const
	{
		return materialId < mRequisiteMaterial.size() ? mRequisiteMaterial[materialId] : 0;
	}
	
	unsigned int Building::GetStockItemCount(const unsigned int& typeItemId) const
	{
		return typeItemId < mItemStock.size() ? mItemStock[typeItemId] : 0;
	}
	
	double Building::GetItemManufacturingProgress(const unsigned int& typeItemId) const
	{
		return typeItemId < mItemMaking.size() ? mItemMaking[typeItemId] : 0;
	}
	
	GameSystem::WorkElement* Building::GetWorkList() const
	{
		return mWorkList;
	}
	
	bool Building::Buildable(const double& x, const double& y, const unsigned int& width, const unsigned int& height) const
	{
		return mX <= x + width && mY <= y + height && x <= mX + mWidth && y <= mY + mHeight;
	}
}
