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
#ifndef CONFIGURATION_HPP
#define CONFIGURATION_HPP

#include "compat/SfmlLegacy.hpp"
#include "tinyxml/tinyxml.h"
#include "myvector.hpp"

#include <memory>
#include <vector>

#include "GameSystem/TypeField.hpp"
#include "GameSystem/TypeRessource.hpp"
#include "GameSystem/TypeBuilding.hpp"
#include "GameSystem/TypeItem.hpp"
#include "GameSystem/TypeRace.hpp"
#include "GameSystem/TypeStatus.hpp"
#include "GameSystem/TypeSkill.hpp"

namespace Sprites
{
	enum Names
	{
		FoundationsSprite
	};
}

// TODO
class ConfigurationClass
{
	public:
		ConfigurationClass();
		~ConfigurationClass();
		
		TiXmlElement* Configure();
		
		legacy::Image* getImage(const unsigned int& idImage);
		legacy::Image* getImage(const TiXmlElement* const element, const char* name);
		legacy::Sprite* GetSprite(const unsigned int& idSprite);
		legacy::Sprite* GetSprite(const Sprites::Names& idSprite);
		legacy::Sprite* GetSprite(const TiXmlElement* const element, const char* name);
		
		unsigned int GetNombreTypesFields() const;
		Types::Field* GetTypeField(const unsigned int& idField);
		unsigned int GetRandomTypeField() const;
		
		unsigned int GetNombreTypesMaterials() const;
		Types::Material* GetTypeMaterial(const unsigned int& materialId);
		
		unsigned int GetNombreTypesRessources() const;
		Types::Ressource* GetTypeRessource(const unsigned int& idTypeRessource);
		Types::Ressource* GetRandomType() const;
		
		unsigned int GetNombreItemCategories() const;
		Types::ItemCategory* GetItemCategory(const unsigned int& idCategory);
		
		unsigned int GetNombreTypesItems() const;
		Types::Item* GetTypeItem(const unsigned int& idTypeItem);
		
		unsigned int GetNombreTypesBuildings() const;
		Types::Building* GetTypeBuilding(const unsigned int& idTypeBuilding);
		
		unsigned int GetNombreRaces() const;
		Types::Race* GetRace(const unsigned int& idRace);
		Types::Race* GetRandomRace() const;
		
		Types::Skill* GetSkill(const unsigned int skillId);
		
		unsigned int GetNombreStatuses() const;
		Types::Status* GetStatus(const unsigned int& idStatus);
	
	private:
		std::vector<std::unique_ptr<TiXmlDocument>> mDocuments;

		Vector <legacy::Image*> mImages;
		Vector <std::string*> mImagePaths;
		
		Vector <legacy::Sprite*> mSprites;
		Vector <int> mSpriteIds; // TODO
		Vector <legacy::Sprite*> mCopiesSprites;
		
		Vector <Types::Field*> mFieldArray;
		Vector <Types::Material*> mMaterialArray;
		Vector <Types::Ressource*> mRessourceArray;
		Vector <Types::ItemCategory*> mItemCategoryArray;
		Vector <Types::Item*> mItemArray;
		Vector <Types::Building*> mBuildingArray;
		Vector <Types::Race*> mRaceArray;
		Vector <Types::Skill*> mSkillArray;
		Vector <Types::Status*> mStatusArray;
};

extern ConfigurationClass* Configuration;

#endif
