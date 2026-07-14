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
#ifndef OBJECT_HPP
#define OBJECT_HPP

#include "../compat/SfmlLegacy.hpp"
#include "../tinyxml/tinyxml.h"

namespace GameSystem
{
	namespace Classes
	{
		enum Names
		{
			Object,
			Character,
			Ressource,
			Building,
			Item
		};
	}
	
	class World;
	
	class Object
	{
		public:
			Object();
			virtual ~Object();
			
			void Initialise(GameSystem::World* const monde, const TiXmlElement* const element);
			virtual void Load();
			virtual TiXmlElement* Save() const;
			
			void draw(sf::RenderTarget* const render) const;
			
			virtual bool Catch(const double& x, const double& y) const;
			void setPosition(const double& x, const double& y);
			
			// Déclencheurs
			virtual void onWhenever() {};
			virtual void onSpawn() {};
			virtual void onAttacked(GameSystem::Object*) {};
			
			unsigned int GetId() const;
			double GetX() const;
			double GetY() const;
			virtual double GetXSprite() const;
			virtual double GetYSprite() const;
			double getWidth() const;
			double getHeight() const;
			GameSystem::Classes::Names GetClass() const;
			
			void PerdHP(const double& amount);
		
		protected:
			GameSystem::World* mWorld;
			// Save de l'object
			const TiXmlElement* mXml;
			
			unsigned int mId;
			GameSystem::Classes::Names mClass;
			double mX;
			double mY;
			double mWidth;
			double mHeight;
			double mMaxHP;
			double mHP;
	};
}

#endif
