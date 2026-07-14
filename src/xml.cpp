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
#include "xml.hpp"

#include <sstream>

#include "GameSystem/World.hpp"

void AddProperty(TiXmlElement* const element, const char* const property, const int& value)
{
	TiXmlElement* selement = new TiXmlElement(property);
	element->LinkEndChild(selement);
	selement->SetAttribute("value", value);
}

void AddProperty(TiXmlElement* const element, const char* const property, const unsigned int& value)
{
	TiXmlElement* selement = new TiXmlElement(property);
	selement->SetAttribute("value", value);
	element->LinkEndChild(selement);
}

void AddProperty(TiXmlElement* const element, const char* const property, const double& value)
{
	TiXmlElement* selement = new TiXmlElement(property);
	selement->SetDoubleAttribute("value", value);
	element->LinkEndChild(selement);
}

void AddProperty(TiXmlElement* const element, const char* const property, const std::string& value)
{
	TiXmlElement* selement = new TiXmlElement(property);
	selement->SetAttribute("value", value.c_str());
	element->LinkEndChild(selement);
}

bool LoadXmlValue(const TiXmlElement* const element, const char* const property, bool& callback)
{
	int value;
	
	if (element->QueryIntAttribute(property, &value) == TIXML_SUCCESS)
	{
		callback = (value != 0);
		return true;
	}
	
	const TiXmlElement* selement = element->FirstChildElement(property);
	while (selement)
	{
		if (selement->QueryIntAttribute("value", &value) == TIXML_SUCCESS)
		{
			callback = (value != 0);
			return true;
		}
		
		selement = selement->NextSiblingElement(property);
	}
	
	return false;
}

bool LoadXmlValue(const TiXmlElement* const element, const char* const property, int& callback)
{
	if (element->QueryIntAttribute(property, &callback) == TIXML_SUCCESS)
		return true;
	
	const TiXmlElement* selement = element->FirstChildElement(property);
	while (selement)
	{
		if (selement->QueryIntAttribute("value", &callback) == TIXML_SUCCESS)
			return true;
		
		selement = selement->NextSiblingElement(property);
	}
	
	return false;
}

bool LoadXmlValue(const TiXmlElement* const element, const char* const property, unsigned int& callback)
{
	int value;
	if (LoadXmlValue(element, property, value) && value >= 0)
	{
		callback = static_cast<unsigned int>(value);
		return true;
	}
	
	return false;
}

bool LoadXmlValue(const TiXmlElement* const element, const char* const property, double& callback)
{
	if (element->QueryDoubleAttribute(property, &callback) == TIXML_SUCCESS)
		return true;
	
	const TiXmlElement* selement = element->FirstChildElement(property);
	while (selement)
	{
		if (selement->QueryDoubleAttribute("value", &callback) == TIXML_SUCCESS)
			return true;
		
		selement = selement->NextSiblingElement(property);
	}
	
	return false;
}

bool LoadXmlValue(const TiXmlElement* const element, const char* const property, std::string& callback)
{
	if (element->Attribute(property))
	{
		callback = element->Attribute(property);
		return true;
	}
	
	const TiXmlElement* selement = element->FirstChildElement(property);
	while (selement)
	{
		if (selement->Attribute("value"))
		{
			callback = selement->Attribute("value");
			return true;
		}
		
		selement = selement->NextSiblingElement(property);
	}
	
	return false;
}

template <class T> void LoadCompactFromXml_Template(const TiXmlElement* const element, const char* const property, Vector<T>& arrayToFill)
{
	std::string str;
	LoadXmlValue(element, property, str);
	
	const char* compactString = str.c_str();
	unsigned int length = strlen(compactString);
	unsigned int rang = 0;
	std::stringstream stream;
	
	unsigned int i = 0;
	while (rang < length)
	{
		char car = compactString[rang++];
		
		if (car == '/')
		{
			stream >> arrayToFill[i++];
			stream.clear();
			stream.str("");
		}
		else
			stream << car;
	}
}

void LoadCompactFromXml(const TiXmlElement* const element, const char* const property, Vector<unsigned int>&	arrayToFill)	{ LoadCompactFromXml_Template(element, property, arrayToFill); }
void LoadCompactFromXml(const TiXmlElement* const element, const char* const property, Vector<double>&		arrayToFill)	{ LoadCompactFromXml_Template(element, property, arrayToFill); }

void LoadCompactFromXml(const TiXmlElement* const element, const char* const property, Vector<GameSystem::Item*>&	arrayToFill, const GameSystem::World* const world)
{
	std::string str;
	LoadXmlValue(element, property, str);
	
	const char* compactString = str.c_str();
	unsigned int length = strlen(compactString);
	unsigned int rang = 0;
	std::stringstream stream;
	
	unsigned int i = 0;
	while (rang < length)
	{
		char car = compactString[rang++];
		
		if (car == '/')
		{
			unsigned int id = 0;
			GameSystem::Object* object;
			
			stream >> id;
			if (id && (object = world->GetObject(id)) && object->GetClass() == GameSystem::Classes::Item)
				arrayToFill[i] = dynamic_cast<GameSystem::Item*>(object);
			
			i++;
			stream.clear();
			stream.str("");
		}
		else
			stream << car;
	}
}

/* TODO
void LoadCompactFromXml(const TiXmlElement* const element, const char* const property, Vector<GameSystem::Item*>& arrayToFill, const GameSystem::World* const world)
{
	LoadCompactFromXml(element, property, arrayToFill.reinterpret<GameSystem::Object*>(), world);
}
*/

template <class T> void SaveCompactToXml_Template(TiXmlElement* const element, const char* const property, const Vector<T>& arrayToSave)
{
	unsigned int length = arrayToSave.size();
	if (length > 0)
	{
		TiXmlElement* selement = new TiXmlElement(property);
		element->LinkEndChild(selement);
		
		std::ostringstream compactString;
		unsigned int i = 0;
		while (i < length)
		{
			if (arrayToSave[i])
				compactString << arrayToSave[i];
			
			compactString << "/";
			
			i++;
		}
		
		selement->SetAttribute("value", compactString.str().c_str());
	}
}

void SaveCompactToXml(TiXmlElement* const element, const char* const property, const Vector<unsigned int>& arrayToSave)	{ SaveCompactToXml_Template(element, property, arrayToSave); }
void SaveCompactToXml(TiXmlElement* const element, const char* const property, const Vector<double>& arrayToSave)	{ SaveCompactToXml_Template(element, property, arrayToSave); }

void SaveCompactToXml(TiXmlElement* const element, const char* const property, const Vector<GameSystem::Item*>& arrayToSave)
{
	unsigned int length = arrayToSave.size();
	if (length > 0)
	{
		TiXmlElement* selement = new TiXmlElement(property);
		element->LinkEndChild(selement);
		
		std::ostringstream compactString;
		unsigned int i = 0;
		while (i < length)
		{
			if (arrayToSave[i])
				compactString << arrayToSave[i]->GetId();
			
			compactString << "/";
			
			i++;
		}
		
		selement->SetAttribute("value", compactString.str().c_str());
	}
}
