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
#ifndef FONCTIONS_HPP
#define FONCTIONS_HPP

#include "GameSystem/Object.hpp"

double Absolute(const double& x);
double Range(const double& x1, const double& y1, const double& x2, const double& y2);
double Range(const GameSystem::Object* const object1, const GameSystem::Object* const object2);
double NormalizeMaterial(const double& amount);
void DrawGauge(sf::RenderTarget* const render, const double& x, const double& y, const unsigned int& width, const unsigned int& height, const double& progress);
void DrawFrame(sf::RenderTarget* const render, legacy::Sprite* const sprite, const double& x, const double& y, const unsigned int& width, const unsigned int& height, const unsigned int& borderWidth, const unsigned int& borderHeight);

#endif
