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
#include "functions.hpp"

#include <sstream>
#include <cmath>

double Absolute(const double& x)
{
	return x > 0 ? x : -x;
}

double Range(const double& x1, const double& y1, const double& x2, const double& y2)
{
	return sqrt(pow(x1 - x2, 2) + pow(y1 - y2, 2));

}

double Range(const GameSystem::Object* const object1, const GameSystem::Object* const object2)
{
	return Range(object1->GetX(), object1->GetY(), object2->GetX(), object2->GetY());
}

double NormalizeMaterial(const double& amount)
{
	return static_cast<double>( floor(amount * 10) / 10 );
}

void DrawGauge(sf::RenderTarget* const render, const double& x, const double& y, const unsigned int& width, const unsigned int& height, const double& progress)
{
	static legacy::Image backgroundImage(1, 1, sf::Color(0, 0, 0, 255));
	static legacy::Sprite backgroundSprite(backgroundImage);
	static legacy::Image gaugeImage(1, 1, sf::Color(255, 255, 255, 255));
	static legacy::Sprite gaugeSprite(gaugeImage);
	
	backgroundSprite.resize(width, height);
	backgroundSprite.setPosition(x, y);
	render->draw(backgroundSprite);
	
	if (progress)
	{
		gaugeImage.setPixel(0, 0, sf::Color( static_cast<sf::Uint8>( 127 - 127 * progress ), static_cast<sf::Uint8>( 255 * progress ), 0, 255) );
		gaugeSprite.resize(width * progress, height);
		gaugeSprite.setPosition(x, y);
		render->draw(gaugeSprite);
	}
}

void DrawFrame(sf::RenderTarget* const render, legacy::Sprite* const sprite, const double& x, const double& y, const unsigned int& width, const unsigned int& height, const unsigned int& borderWidth, const unsigned int& borderHeight)
{
	if (!sprite)
		return;
	
	const legacy::Image* image = sprite->getImage();
	unsigned int imageWidth = image->getWidth();
	unsigned int imageHeight = image->getHeight();
	
	// Top-left corner
	sprite->setSubRect( sf::IntRect(0, 0, borderWidth, borderHeight) );
	sprite->resize(borderWidth, borderHeight);
	sprite->setPosition(x, y);
	render->draw(*sprite);
	
	// Top side
	sprite->setSubRect( sf::IntRect(borderWidth, 0, imageWidth - borderWidth, borderHeight) );
	sprite->resize(width - 2 * borderWidth, borderHeight);
	sprite->setPosition(x + borderWidth, y);
	render->draw(*sprite);
	
	// Top-right corner
	sprite->setSubRect( sf::IntRect(imageWidth - borderWidth, 0, imageWidth, borderHeight) );
	sprite->resize(borderWidth, borderHeight);
	sprite->setPosition(x + width - borderWidth, y);
	render->draw(*sprite);
	
	// Left side
	sprite->setSubRect( sf::IntRect(0, borderHeight, borderWidth, imageHeight - borderHeight) );
	sprite->resize(borderWidth, height - 2 * borderHeight);
	sprite->setPosition(x, y + borderHeight);
	render->draw(*sprite);
	
	// Main
	sprite->setSubRect( sf::IntRect(borderWidth, borderHeight, imageWidth - borderWidth, imageHeight - borderHeight) );
	sprite->setPosition(x + borderWidth, y + borderHeight);
	sprite->resize(width - 2 * borderWidth, height - 2 * borderHeight);
	render->draw(*sprite);
	
	// Right side
	sprite->setSubRect( sf::IntRect(imageWidth - borderWidth, borderHeight, imageWidth, imageHeight - borderHeight) );
	sprite->resize(borderWidth, height - 2 * borderHeight);
	sprite->setPosition(x + width - borderWidth, y + borderHeight);
	render->draw(*sprite);
	
	// Bottom-left corner
	sprite->setSubRect( sf::IntRect(0, imageHeight - borderHeight, borderWidth, imageHeight) );
	sprite->resize(borderWidth, borderHeight);
	sprite->setPosition(x, y + height - borderHeight);
	render->draw(*sprite);
	
	// Bottom side
	sprite->setSubRect( sf::IntRect(borderWidth, imageHeight - borderHeight, imageWidth - borderWidth, imageHeight) );
	sprite->resize(width - 2 * borderWidth, borderHeight);
	sprite->setPosition(x + borderWidth, y + height - borderHeight);
	render->draw(*sprite);
	
	// Bottom-right corner
	sprite->setSubRect( sf::IntRect(imageWidth - borderWidth, imageHeight - borderHeight, imageWidth, imageHeight) );
	sprite->resize(borderWidth, borderHeight);
	sprite->setPosition(x + width - borderWidth, y + height - borderHeight);
	render->draw(*sprite);
}
