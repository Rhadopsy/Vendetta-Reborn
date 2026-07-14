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
#include "Configuration.hpp"

#include <iostream>
#include <algorithm>
#include <filesystem>
#include <memory>
#include <vector>

#include "xml.hpp"
#include "GameSystem/Object.hpp"
#include "data.hpp"

namespace
{
	template <typename T>
	void DeletePointers(Vector<T*>& values)
	{
		for (unsigned int i = 0; i < values.size(); ++i)
			delete values[i];
		values.clear();
	}
}

#define XML_PROCESSING(bloc, whattodo) \
	i = 0; \
	while (i < documents.size()) \
	{ \
		if ( (element = documents[i++]->RootElement()) && (element = element->FirstChildElement(#bloc)) ) \
			whattodo \
	}
#define FOREACH_XML_ELEMENT(name, whattodo) \
	element = element->FirstChildElement(#name); \
	while (element) \
	{ \
		whattodo \
		element = element->NextSiblingElement(#name); \
	}
#define LOAD_EACH(name) \
	FOREACH_XML_ELEMENT(name,\
		if (LoadXmlValue(element, "id", id)) \
		{ \
			delete m##name##Array[id]; \
			m##name##Array[id] = new Types::name(element); \
		} \
	)

ConfigurationClass::ConfigurationClass()
{
}

ConfigurationClass::~ConfigurationClass()
{
	DeletePointers(mRaceArray);
	DeletePointers(mBuildingArray);
	DeletePointers(mItemArray);
	DeletePointers(mItemCategoryArray);
	DeletePointers(mSkillArray);
	DeletePointers(mStatusArray);
	DeletePointers(mRessourceArray);
	DeletePointers(mMaterialArray);
	DeletePointers(mFieldArray);
	// Copies are historically owned by a mix of game types and interface widgets.
	// Their ownership will be unified in a later refactor; deleting them here can
	// double-free the copies already released by type destructors.
	mCopiesSprites.clear();
	DeletePointers(mSprites);
	DeletePointers(mImages);
	DeletePointers(mImagePaths);
}

TiXmlElement* ConfigurationClass::Configure()
{
	std::vector<std::unique_ptr<TiXmlDocument>> documents;
	unsigned int i = 0;
	std::vector<std::filesystem::path> configFiles;
	const std::filesystem::path configDirectory("./Data/Config/");
	if (!std::filesystem::exists(configDirectory))
	{
		std::cerr << "Repertoire de configuration introuvable: " << configDirectory << "\n";
		return NULL;
	}
	for (const auto& entry : std::filesystem::directory_iterator(configDirectory))
		if (entry.is_regular_file() && entry.path().extension() == ".xml")
			configFiles.push_back(entry.path());
	std::sort(configFiles.begin(), configFiles.end());
	for (const auto& filename : configFiles)
	{
		documents.push_back(std::make_unique<TiXmlDocument>());
		documents[i]->LoadFile(filename.string().c_str());
		i++;
	}
	
	TiXmlElement* element;
	unsigned int id;
	std::string string = "";
	unsigned int integer = 0;
	
	XML_PROCESSING(Images,
	{
		FOREACH_XML_ELEMENT(Image,
		{
			if (LoadXmlValue(element, "id", id) && LoadXmlValue(element, "File", string))
			{
				delete mImagePaths[id];
				mImagePaths[id] = new std::string(string);
			}
		})
	})
	Prgm::Progression = 0.1;
	
	XML_PROCESSING(Sprites,
	{
		FOREACH_XML_ELEMENT(Sprite,
		{
			if (LoadXmlValue(element, "id", id))
			{
				legacy::Image* image = NULL;
				if (LoadXmlValue(element, "Image", integer) && (image = this->getImage(integer)) )
				{
					delete mSprites[id];
					mSprites[id] = new legacy::Sprite();
					mSprites[id]->setImage(*image);
					
					int x = 0;
					int y = 0;
					int w = 0;
					int h = 0;
					
					if (!LoadXmlValue(element, "x", x))
						x = 0;
					if (!LoadXmlValue(element, "y", y))
						y = 0;
					if (!LoadXmlValue(element, "width", w))
						w = image->getWidth();
					if (!LoadXmlValue(element, "height", h))
						h = image->getHeight();
					
					mSprites[id]->setSubRect( sf::IntRect(x, y, x + w, y + h) );
				}
				else
				{
					mSprites[id] = new legacy::Sprite(legacy::Image(1, 1, sf::Color::Transparent));
					std::cerr << "Erreur lors du chargement du sprite " << id <<"\n";
				}
			}
		})
	})
	Prgm::Progression = 0.2;
	
	XML_PROCESSING(Fields,
	{
		LoadXmlValue(element, "Width", Prgm::Fields::Width);
		LoadXmlValue(element, "Height", Prgm::Fields::Height);
		
		LOAD_EACH(Field)
	})
	Prgm::Progression = 0.3;
	
	XML_PROCESSING(Materials,
	{
		LoadXmlValue(element, "IconSize", Prgm::Materials::IconSize);
		
		LOAD_EACH(Material)
	})
	
	XML_PROCESSING(Ressources, LOAD_EACH(Ressource) )
	Prgm::Progression = 0.4;
	
	XML_PROCESSING(ItemCategories, LOAD_EACH(ItemCategory) )
	Prgm::Progression = 0.5;
	
	XML_PROCESSING(Items,
	{
		LoadXmlValue(element, "IconSize", Prgm::Items::IconSize);
		
		LOAD_EACH(Item)
	})
	Prgm::Progression = 0.6;
	
	XML_PROCESSING(Buildings,
	{
		LoadXmlValue(element, "SpriteButtonSize", Prgm::TypesBuildings::SpriteButtonSize);
		LoadXmlValue(element, "ButtonSize", Prgm::TypesBuildings::ButtonSize);
		LoadXmlValue(element, "FoundationsSprite", mSpriteIds[Sprites::FoundationsSprite]);
		
		LOAD_EACH(Building)
	})
	
	XML_PROCESSING(Races, LOAD_EACH(Race) )
	Prgm::Progression = 0.7;
	
	XML_PROCESSING(Skills, LOAD_EACH(Skill) )
	Prgm::Progression = 0.8;
	
	XML_PROCESSING(Statuses,
	{
		LoadXmlValue(element, "Width", Prgm::Statuses::Width);
		LoadXmlValue(element, "Height", Prgm::Statuses::Height);
		LoadXmlValue(element, "Border", Prgm::Statuses::Border);
		LoadXmlValue(element, "TextSize", Prgm::Statuses::TextSize);
		
		LOAD_EACH(Status)
	})
	
	Prgm::Progression = 0.9;
	
	XML_PROCESSING(Interface,
	{
		return element;
	})
	
	return NULL;
}

legacy::Image* ConfigurationClass::getImage(const unsigned int& imageId)
{
	if (imageId < mImagePaths.size())
	{
		if (!mImages[imageId])
		{
			if (!mImagePaths[imageId] || !(mImages[imageId] = new legacy::Image()) || !mImages[imageId]->loadFromFile(R_IMAGES + *mImagePaths[imageId] ))
				std::cerr << "Erreur lors du chargement de l'image " << imageId << "\n";
			
			mImages[imageId]->setSmooth(false);
		}
		return mImages[imageId];
	}
	else
		return NULL;
}

legacy::Image* ConfigurationClass::getImage(const TiXmlElement* const element, const char* name)
{
	unsigned int imageId;
	return element && LoadXmlValue(element, name, imageId) ? this->getImage(imageId) : NULL;
}

legacy::Sprite* ConfigurationClass::GetSprite(const unsigned int& spriteId)
{
	if (spriteId < mSprites.size())
	{
		legacy::Sprite* newSprite = new legacy::Sprite(*mSprites[spriteId]);
		mCopiesSprites.push_back(newSprite);
		return newSprite;
	}
	else
	{
		std::cerr << "Demande du sprite " << spriteId << " non charge\n";
		legacy::Image newImage(1, 1, sf::Color::Transparent);
		newImage.setPixel(0, 0, sf::Color(0, 255, 0, 255) );
		legacy::Sprite* newSprite = new legacy::Sprite(newImage);
		mCopiesSprites.push_back(newSprite);
		return newSprite;
	}
}

legacy::Sprite* ConfigurationClass::GetSprite(const Sprites::Names& spriteId)
{
	return this->GetSprite(mSpriteIds[spriteId]);
}

legacy::Sprite* ConfigurationClass::GetSprite(const TiXmlElement* const element, const char* name)
{
	unsigned int spriteId;
	return element && LoadXmlValue(element, name, spriteId) ? this->GetSprite(spriteId) : NULL;
}

unsigned int ConfigurationClass::GetNombreTypesFields() const
{
	return mFieldArray.size();
}

Types::Field* ConfigurationClass::GetTypeField(const unsigned int& fieldId)
{
	return fieldId < mFieldArray.size() ? mFieldArray[fieldId] : NULL;
}

unsigned int ConfigurationClass::GetRandomTypeField() const
{
	bool continuer = true;
	unsigned int fieldId = 0;
	unsigned int nbTypesFields = mFieldArray.size();
	Types::Field* typeField = NULL;
	while (continuer && fieldId < nbTypesFields) /* TODO I failed a lot */
	{
		typeField = mFieldArray[fieldId];
		if (typeField && legacy::random(0.f, 1.f) <= typeField->GetProbability())
			continuer = false;
		else
			fieldId++;
	}
	return fieldId;
}

unsigned int ConfigurationClass::GetNombreTypesMaterials() const
{
	return mMaterialArray.size();
}

Types::Material* ConfigurationClass::GetTypeMaterial(const unsigned int& materialId)
{
	return materialId < mMaterialArray.size() ? mMaterialArray[materialId] : NULL;
}


unsigned int ConfigurationClass::GetNombreTypesRessources() const
{
	return mRessourceArray.size();
}

Types::Ressource* ConfigurationClass::GetTypeRessource(const unsigned int& idTypeRessource)
{
	return idTypeRessource < mRessourceArray.size() ? mRessourceArray[idTypeRessource] : NULL;
}

Types::Ressource* ConfigurationClass::GetRandomType() const
{
	Types::Ressource* typeRessource = NULL;
	unsigned int nbTypesRessources = mRessourceArray.size();
	while (!(typeRessource = mRessourceArray[legacy::random(0, nbTypesRessources - 1)]));
	return typeRessource;
}

unsigned int ConfigurationClass::GetNombreItemCategories() const
{
	return mItemCategoryArray.size();
}

Types::ItemCategory* ConfigurationClass::GetItemCategory(const unsigned int& idCategory)
{
	return idCategory < mItemCategoryArray.size() ? mItemCategoryArray[idCategory] : NULL;
}

unsigned int ConfigurationClass::GetNombreTypesItems() const
{
	return mItemArray.size();
}

Types::Item* ConfigurationClass::GetTypeItem(const unsigned int& idTypeItem)
{
	return idTypeItem < mItemArray.size() ? mItemArray[idTypeItem] :  NULL;
}

unsigned int ConfigurationClass::GetNombreTypesBuildings() const
{
	return mBuildingArray.size();
}

Types::Building* ConfigurationClass::GetTypeBuilding(const unsigned int& idTypeBuilding)
{
	return idTypeBuilding < mBuildingArray.size() ? mBuildingArray[idTypeBuilding] : NULL;
}

unsigned int ConfigurationClass::GetNombreRaces() const
{
	return mRaceArray.size();
}

Types::Race* ConfigurationClass::GetRace(const unsigned int& idRace)
{
	return idRace < mRaceArray.size() ? mRaceArray[idRace] : NULL;
}

Types::Race* ConfigurationClass::GetRandomRace() const
{
	return mRaceArray[legacy::random(0, mRaceArray.size()) - 1];
}

Types::Skill* ConfigurationClass::GetSkill(const unsigned int skillId)
{
	return skillId < mSkillArray.size() ? mSkillArray[skillId] : NULL;
}

unsigned int ConfigurationClass::GetNombreStatuses() const
{
	return mStatusArray.size();
}

Types::Status* ConfigurationClass::GetStatus(const unsigned int& idStatus)
{
	return idStatus < mStatusArray.size() ? mStatusArray[idStatus] : NULL;
}
