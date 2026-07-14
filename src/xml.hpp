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
#ifndef XML_HPP
#define XML_HPP

//#define TIXML_USE_STL
#include "tinyxml/tinyxml.h"
#include "myvector.hpp"

#include "GameSystem/Item.hpp"

void AddProperty(TiXmlElement* const element, const char* const property, const int&		value);
void AddProperty(TiXmlElement* const element, const char* const property, const unsigned int&	value);
void AddProperty(TiXmlElement* const element, const char* const property, const double&		value);
void AddProperty(TiXmlElement* const element, const char* const property, const std::string&	value);

bool LoadXmlValue(const TiXmlElement* const element, const char* const property, bool&		callback);
bool LoadXmlValue(const TiXmlElement* const element, const char* const property, int&		callback);
bool LoadXmlValue(const TiXmlElement* const element, const char* const property, unsigned int&	callback);
bool LoadXmlValue(const TiXmlElement* const element, const char* const property, double&	callback);
bool LoadXmlValue(const TiXmlElement* const element, const char* const property, std::string&	callback);

void LoadCompactFromXml(const TiXmlElement* const element, const char* const property, Vector<unsigned int>&		arrayToFill);
void LoadCompactFromXml(const TiXmlElement* const element, const char* const property, Vector<double>&			arrayToFill);
void LoadCompactFromXml(const TiXmlElement* const element, const char* const property, Vector<GameSystem::Item*>&	arrayToFill, const GameSystem::World* const world);

void SaveCompactToXml(TiXmlElement* const element, const char* const property, const Vector<unsigned int>&		arrayToSave);
void SaveCompactToXml(TiXmlElement* const element, const char* const property, const Vector<double>&			arrayToSave);
void SaveCompactToXml(TiXmlElement* const element, const char* const property, const Vector<GameSystem::Item*>&	arrayToSave);

#endif
