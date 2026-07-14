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
#include "World.hpp"

#include <iostream>
#include <cmath>
#include <sstream>

#include "../xml.hpp"
#include "../Configuration.hpp"
#include "../data.hpp"
#include "../functions.hpp"

#include "Character.hpp"
#include "TypeField.hpp"

#ifndef VENDETTA_ENABLE_PYTHON
#define VENDETTA_ENABLE_PYTHON 0
#endif

#if VENDETTA_ENABLE_PYTHON
#include "scripts.hpp"
#endif

#define NOMBRE_PERSONNAGES 5

namespace GameSystem
{
	World::World() :
		mWidthEnFields(0),
		mHeightEnFields(0),
		mGrilleFields(NULL),
		mImage(NULL),
		mSprite(NULL),
		mCharacterJoueur(NULL)
	{
	}
	
	World::~World()
	{
		for (unsigned int i = 0; i < mObjects.size(); ++i)
			delete mObjects[i];
		delete mSprite;
		delete mImage;
		delete[] mGrilleFields;
	}
	
	void World::Load(const TiXmlElement* const element)
	{
		if (element)
		{
			LoadXmlValue(element, "WidthCarte", mWidthEnFields);
			LoadXmlValue(element, "HeightCarte", mHeightEnFields);
		}
		else
		{
			mWidthEnFields = legacy::random(64, 256);
			mHeightEnFields = legacy::random(64, 256);
		}
		
		mWidth = mWidthEnFields * Prgm::Fields::Width;
		mHeight = mHeightEnFields * Prgm::Fields::Height;
		
		// creation du field
		mGrilleFields = new unsigned int[mWidthEnFields * mHeightEnFields];
		unsigned int nbTypesFields = Configuration->GetNombreTypesFields();
		mImage = new legacy::Image(static_cast<int>(mWidth), static_cast<int>(mHeight));
		mSprite = new legacy::Sprite(*mImage);
		
		std::string stringXml;
		unsigned int rang = 0;
		unsigned int length = 0;
		if (element)
		{
			LoadXmlValue(element, "GrilleFields", stringXml);
			length = stringXml.length();
		}
		
		unsigned int line = 0;
		while (line < mHeightEnFields)
		{
			unsigned int column = 0;
			while (column < mWidthEnFields)
			{
				unsigned int idField = 0;
				Types::Field* typeField = NULL;
				if (element)
				{
					char caractere = 0;
					std::stringstream stream;
					
					while (rang < length && (caractere = stringXml.c_str()[rang++]) != '/')
						stream << caractere;
					
					stream >> idField;
					
					if (idField >= nbTypesFields)
						idField = 0;
				}
				else
					idField = Configuration->GetRandomTypeField();
				
				mGrilleFields[line * mWidthEnFields + column] = idField;
				
				typeField = Configuration->GetTypeField(idField);
				
				legacy::Sprite* spriteTer;
				if (typeField && (spriteTer = typeField->GetSprite()) )
				{
					unsigned int k = 0;
					while (k < Prgm::Fields::Width)
					{
						unsigned int l = 0;
						while (l < Prgm::Fields::Height)
						{
							sf::Color color = spriteTer->getPixel(k, l);
							mImage->setPixel(column * Prgm::Fields::Width + k, line * Prgm::Fields::Height + l, color);
							
							l++;
						}
						k++;
					}
				}
				column++;
			}
			line++;
		}
		
		
#if VENDETTA_ENABLE_PYTHON
		bool pyClasseCharacterLoade = false;
		boost::python::object pyClasseCharacter;
		try
		{
			PyImport_AppendInittab("Vendetta", &PyInit_Vendetta);
			Py_Initialize();
			boost::python::object main = boost::python::import("__main__");
			boost::python::object global = main.attr("__dict__");
			boost::python::exec_file(R_SCRIPT_BASE, global, global);
			boost::python::exec_file(R_SCRIPT_PERSO, global, global);
			pyClasseCharacter = global["Character"];
			pyClasseCharacterLoade = true;
		}
		catch (...)
		{
			PyErr_Print();
		}
#endif
		
		// creation des objects
		if (element)
		{
			const TiXmlElement* xmlObjects = element->FirstChildElement("Objects");
			while (xmlObjects)
			{
				const TiXmlElement* xmlTerRessouces = xmlObjects->FirstChildElement("Ressource");
				while (xmlTerRessouces)
				{
					GameSystem::Ressource* nouvelleRessource = new GameSystem::Ressource();
					this->AjouteObject(nouvelleRessource, xmlTerRessouces);
					mRessources.push_back(nouvelleRessource);
					
					xmlTerRessouces = xmlTerRessouces->NextSiblingElement("Ressource");
				}
				
				const TiXmlElement* xmlCharacters = xmlObjects->FirstChildElement("Character");
				while (xmlCharacters)
				{
					GameSystem::Character* nouveauCharacter;
#if VENDETTA_ENABLE_PYTHON
					if (pyClasseCharacterLoade)
					{
						boost::python::object* pyCharacter = new boost::python::object(pyClasseCharacter());
						nouveauCharacter = boost::python::extract<GameSystem::Character*>(*pyCharacter) BOOST_EXTRACT_WORKAROUND;
					}
					else
						nouveauCharacter = new GameSystem::Character();
#else
					nouveauCharacter = new GameSystem::Character();
#endif
					
					this->AjouteObject(nouveauCharacter, xmlCharacters);
					mCharacters.push_back(nouveauCharacter);
					
					xmlCharacters = xmlCharacters->NextSiblingElement("Character");
				}
				
				const TiXmlElement* xmlBuildings = xmlObjects->FirstChildElement("Building");
				while (xmlBuildings)
				{
					GameSystem::Building* nouveauBuilding = new GameSystem::Building();
					this->AjouteObject(nouveauBuilding, xmlBuildings);
					mBuildings.push_back(nouveauBuilding);
					
					xmlBuildings = xmlBuildings->NextSiblingElement("Building");
				}
				
				const TiXmlElement* xmlItems = xmlObjects->FirstChildElement("Item");
				while (xmlItems)
				{
					GameSystem::Item* nouvelItem = new GameSystem::Item();
					this->AjouteObject(nouvelItem, xmlItems);
					mItems.push_back(nouvelItem);
					
					xmlItems = xmlItems->NextSiblingElement("Item");
				}
				
				xmlObjects = xmlObjects->NextSiblingElement("Objects");
			}
		}
		else
		{
			unsigned int nombreRessources = legacy::random(1, mWidthEnFields * mHeightEnFields / 500);
			while (mRessources.size() < nombreRessources)
			{
				GameSystem::Ressource* nouvelleRessource = new GameSystem::Ressource();
				this->AjouteObject(nouvelleRessource);
				mRessources.push_back(nouvelleRessource);
				
				double x;
				double y;
				bool continuer = true;
				while (continuer)
				{
					x = legacy::random(0.0, static_cast<float>(mWidth) );
					y = legacy::random(0.0, static_cast<float>(mHeight) );
					nouvelleRessource->SetType( Configuration->GetRandomType() );
					
					continuer = !this->RessourceFits(nouvelleRessource, x, y);
				}
				
				nouvelleRessource->setPosition(x, y);
			}
			
			while (mCharacters.size() < NOMBRE_PERSONNAGES)
			{
				GameSystem::Character* nouveauCharacter;
#if VENDETTA_ENABLE_PYTHON
				if (pyClasseCharacterLoade)
				{
					boost::python::object* pyCharacter = new boost::python::object(pyClasseCharacter());
					nouveauCharacter = boost::python::extract<GameSystem::Character*>(*pyCharacter) BOOST_EXTRACT_WORKAROUND;
				}
				else
					nouveauCharacter = new GameSystem::Character();
#else
				nouveauCharacter = new GameSystem::Character();
#endif
				
				this->AjouteObject(nouveauCharacter);
				mCharacters.push_back(nouveauCharacter);
				
				nouveauCharacter->onSpawn();
			}
		}
		
		unsigned int i = 0;
		while (i < mObjects.size())
		{
			if (mObjects[i])
				mObjects[i]->Load();
			i++;
		}
		
		*mCharacterJoueur = mCharacters[0];
	}
	
	TiXmlElement* World::Save() const
	{
		TiXmlElement* saveNode = new TiXmlElement("World");
		TiXmlElement* selement;
		TiXmlElement* sselement;
		
		AddProperty(saveNode, "WidthCarte", mWidthEnFields);
		AddProperty(saveNode, "HeightCarte", mHeightEnFields);
		
		selement = new TiXmlElement("GrilleFields");
		saveNode->LinkEndChild(selement);
		std::ostringstream grilleFields;
		unsigned int i = 0;
		unsigned long int nb = mWidthEnFields * mHeightEnFields;
		while (i < nb)
		{
			if (mGrilleFields[i])
				grilleFields << mGrilleFields[i];
			
			grilleFields << "/";
			
			i++;
		}
		selement->SetAttribute("value", grilleFields.str().c_str());
		
		selement = new TiXmlElement("Objects");
		saveNode->LinkEndChild(selement);
		i = 0;
		while (i < mObjects.size())
		{
			if (mObjects[i])
			{
				sselement = mObjects[i]->Save();
				if (sselement)
					selement->LinkEndChild(sselement);
			}
			i++;
		}
		
		return saveNode;
	}
	
	void World::SetPlayer(GameSystem::Character** const personnage)
	{
		mCharacterJoueur = personnage;
	}
	
	void World::draw(sf::RenderTarget* const render)
	{
		render->draw(*mSprite);
		
		unsigned int i = 0;
		while (i < mRessources.size())
		{
			mRessources[i]->draw(render);
			i++;
		}
		
		i = 0;
		while (i < mBuildings.size())
		{
			mBuildings[i]->draw(render);
			i++;
		}
		
		i = 0;
		while (i < mCharacters.size())
		{
			mCharacters[i]->draw(render);
			i++;
		}
	}
	
	void World::Process(const double& time)
	{
		unsigned int i = 0;
		while (i < mCharacters.size())
		{
			mCharacters[i]->Process(time);
			i++;
		}
	}
	
	double World::getWidth() const
	{
		return mWidth;
	}
	
	double World::getHeight() const
	{
		return mHeight;
	}
	
	GameSystem::Object* World::GetObjectPoint(const double& x, const double& y) const
	{
		if (0 <= x && x < mWidth && 0 <= y && y < mHeight)
		{
			unsigned int i = 0;
			while (i < mObjects.size())
			{
				if (mObjects[i] && mObjects[i]->Catch(x, y))
					return mObjects[i];
				
				i++;
			}
		}
		
		return NULL;
	}
	
	int World::GetIdField(const double& x, const double& y) const
	{
		if (0 <= x && x < mWidth && 0 <= y && y < mHeight)
		{
			int fieldLine = static_cast<int>(x / Prgm::Fields::Width);
			int fieldColumn = static_cast<int>(y / Prgm::Fields::Height);
			//std::cout << mGrilleFields[fieldLine * mWidthEnFields + fieldColumn] << "\n"; //TODO
			return mGrilleFields[fieldLine * mWidthEnFields + fieldColumn];
		}
		else
			return -1;
	}
	
	void World::AjouteObject(GameSystem::Object* const object, const TiXmlElement* const element)
	{
		object->Initialise(this, element);
		mObjects[object->GetId()] = object;
	}

	GameSystem::Object* World::GetObject(const unsigned int& objectId) const
	{
		return objectId < mObjects.size() ? mObjects[objectId] : NULL;
	}
	
	void World::RemoveObject(const unsigned int& objectId)
	{
		delete mObjects[objectId];
		mObjects[objectId] = NULL;
	}
	
	GameSystem::Ressource* World::FindRessource(const GameSystem::Object* const origin, Types::Ressource* const type) const
	{
		GameSystem::Ressource* nearestRessource = NULL;
		double shorterDistance = 0;
		
		unsigned int i = 0;
		while (i < mRessources.size())
		{
			GameSystem::Ressource* ressource = mRessources[i++];
			if (ressource->GetTypeRessource() == type)
			{
				double distance = Range(origin, ressource);
				if (!nearestRessource || distance < shorterDistance)
				{
					nearestRessource = ressource;
					shorterDistance = distance;
				}
			}
		}
		
		return nearestRessource;
	}
	
	unsigned int World::AssignNation(GameSystem::Character*)
	{
		return legacy::random(0, 2);
	}
	
	bool World::RessourceFits(GameSystem::Ressource* ressource, const double& x, const double& y) const
	{
		unsigned int MinX = x < 0 ? 0 : static_cast<unsigned int>(x / Prgm::Fields::Width);
		unsigned int MinY = y < 0 ? 0 : static_cast<unsigned int>(y / Prgm::Fields::Height);
		unsigned int MaxX = MinX + static_cast<unsigned int>(ressource->getWidth() / Prgm::Fields::Width);
		unsigned int MaxY = MinY + static_cast<unsigned int>(ressource->getHeight() / Prgm::Fields::Height);
		
		if (MaxX > mWidthEnFields)
			MaxX = mWidthEnFields;
		if (MaxY > mHeightEnFields)
			MaxY = mHeightEnFields;
		
		unsigned int column = MinX;
		while (column < MaxX)
		{
			unsigned int line = MinY;
			while (line < MaxY)
			{
				if (!Configuration->GetTypeField(mGrilleFields[line++ * mWidthEnFields + column])->AcceptRessource())
					return false;
			}
			column++;
		}
		
		unsigned int i = 0;
		while (i < mRessources.size())
		{
			if (mRessources[i++]->Catch(x, y))
				return false;
		}
		
		return true;
	}
	
	bool World::Buildable(const double& x, const double& y, Types::Building* const building) const
	{
		unsigned int width = building->GetStructureWidth();
		unsigned int height = building->GetStructureHeight();
		
		if (x < 0 || x + width > mWidth || y < 0 || y + height > mHeight)
			return false;
		
		int MinX = static_cast<unsigned int>(x / Prgm::Fields::Width);
		int MinY = static_cast<unsigned int>(y / Prgm::Fields::Height);
		int MaxX = MinX + width / Prgm::Fields::Width;
		int MaxY = MinY + height / Prgm::Fields::Height;
		
		if (MinX < 0)
			MinX = 0;
		if (MinY < 0)
			MinY = 0;
		if (MaxX > static_cast<int>(mWidthEnFields))
			MaxX = mWidthEnFields;
		if (MaxY > static_cast<int>(mHeightEnFields))
			MaxY = mHeightEnFields;
		
		int fieldLine = MinY;
		while (fieldLine < MaxY)
		{
			int fieldColumn = MinX;
			while (fieldColumn < MaxX)
			{
				if (!Configuration->GetTypeField(mGrilleFields[fieldLine * mWidthEnFields + fieldColumn++])->Buildable())
					return false;
			}
			fieldLine++;
		}
		
		unsigned int i = 0;
		while (i < mRessources.size())
		{
			if (mRessources[i++]->Buildable(x, y, width, height))
				return false;
		}
		
		i = 0;
		while (i < mBuildings.size())
		{
			if (mBuildings[i++]->Buildable(x, y, width, height))
				return false;
		}
		
		return true;
	}
	
	GameSystem::Building* World::SeatBuilding(const double& x, const double& y, Types::Building* const type)
	{
		if (this->Buildable(x, y, type))
		{
			GameSystem::Building* building = new GameSystem::Building(type);
			this->AjouteObject(building);
			building->Load();
			building->setPosition(x, y);
			mBuildings.push_back(building);
			return building;
		}
		else
			return NULL;
	}
	
	GameSystem::Item* World::CreateItem(Types::Item* const type)
	{
		GameSystem::Item* nouvelItem = new GameSystem::Item(type);
		this->AjouteObject(nouvelItem);
		mItems.push_back(nouvelItem);
		
		return nouvelItem;
	}
}
