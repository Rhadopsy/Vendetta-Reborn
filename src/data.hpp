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
#ifndef DATA_HPP
#define DATA_HPP

// Repertoires
#define R_IMAGES "./Data/Images/"
#define R_IMAGES_BATIMENTS_STRUCTURES "./Data/Images/Buildings/Structures/"
#define R_IMAGES_BATIMENTS_BOUTONS "./Data/Images/Buildings/Buttons/"
#define R_IMAGES_OBJETS "./Data/Images/Items/"
#define R_IMAGES_RESSOURCES "./Data/Images/Materials/"
#define R_MUSIQUES "./Data/Musiques/"
#define R_CONFIG "./Data/Config/"
#define R_SCRIPTS "./Data/Bots/"
#define R_SCRIPT_BASE "./Data/Bots/base.py"
#define R_SCRIPT_PERSO "./Data/Bots/script.py"

namespace Prgm
{
	extern double Progression;
	
	namespace Fields
	{
		extern unsigned int Width;
		extern unsigned int Height;
	}
	
	namespace Materials
	{
		extern unsigned int IconSize;
	}
	
	namespace Items
	{
		extern unsigned int IconSize;
	}
	
	namespace TypesBuildings
	{
		extern unsigned int SpriteButtonSize;
		extern unsigned int ButtonSize;
	}
	
	namespace Statuses
	{
		extern double Width;
		extern double Height;
		extern double Border;
		extern double TextSize;
	}
}

#endif
