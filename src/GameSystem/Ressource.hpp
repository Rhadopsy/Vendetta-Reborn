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
#ifndef RESSOURCE_HPP
#define RESSOURCE_HPP

#include "Object.hpp"
#include "TypeRessource.hpp"

namespace GameSystem
{
	// Material de field
	class Ressource : public GameSystem::Object
	{
		public:
			Ressource();
			
			virtual void Load();
			virtual TiXmlElement* Save() const;
			
			double Harvest(GameSystem::Object* const collecteur, const double& time);
			
			std::string getText() const;
			
			void draw(sf::RenderTarget* const render);
			
			// Type de material de field
			void SetType(Types::Ressource* const type);
			Types::Ressource* GetTypeRessource() const;
			
			bool Buildable(const double& x, const double& y, const unsigned int& width, const unsigned int& height) const;
			
		private:
			Types::Ressource* mType;
	};
}

#endif
