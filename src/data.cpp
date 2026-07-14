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
#include "data.hpp"

namespace Prgm
{
	double Progression = 0;
	
	namespace World
	{
		double Pesanteur = 9.81;
	}
	
	namespace Fields
	{
		unsigned int Width = 16;
		unsigned int Height = 16;
	}
	
	namespace Materials
	{
		unsigned int IconSize = 24;
	}
	
	namespace Items
	{
		unsigned int IconSize = 24;
	}
	
	namespace TypesBuildings
	{
		unsigned int SpriteButtonSize = 28;
		unsigned int ButtonSize = 28;
	}
	
	namespace WinInventory
	{
		unsigned int LineCount = 8;
		unsigned int ColumnCount = 9;
	}
	
	namespace Statuses
	{
		double Width = 100;
		double Height = 20;
		double Border = 2;
		double TextSize = 15;
	}
}
