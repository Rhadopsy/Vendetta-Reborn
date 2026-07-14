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
#include "Object.hpp"

#include "../xml.hpp"

namespace GameSystem
{
	Object::Object() :
		mXml(NULL),
		mId(0),
		mClass(GameSystem::Classes::Object),
		mX(0),
		mY(0),
		mWidth(0),
		mHeight(0),
		mMaxHP(100),
		mHP(mMaxHP)
	{
	}
	
	Object::~Object()
	{
	}
	
	void Object::Initialise(GameSystem::World* const monde, const TiXmlElement* const element)
	{
		static unsigned int id = 1;
		
		mWorld = monde;
		mXml = element;
		
		if (mXml)
			LoadXmlValue(mXml, "id", mId);
		else
			mId = id++;
		
		if (mId >= id)
			id = mId + 1;
	}
	
	void Object::Load()
	{
	}
	
	TiXmlElement* Object::Save() const
	{
		TiXmlElement* saveNode = new TiXmlElement("Object");
		TiXmlElement* selement;
		
		saveNode->SetAttribute("id", mId);
		
		selement = new TiXmlElement("x");
		saveNode->LinkEndChild(selement);
		selement->SetDoubleAttribute("value", mX);
		
		selement = new TiXmlElement("y");
		saveNode->LinkEndChild(selement);
		selement->SetDoubleAttribute("value", mY);
		
		selement = new TiXmlElement("Width");
		saveNode->LinkEndChild(selement);
		selement->SetDoubleAttribute("value", mWidth);
		
		selement = new TiXmlElement("Height");
		saveNode->LinkEndChild(selement);
		selement->SetDoubleAttribute("value", mHeight);
		
		return saveNode;
	}
	
	bool Object::Catch(const double& x, const double& y) const
	{
		return mX <= x && x < mX + mWidth && mY <= y && y <= mY + mHeight;
	}
	
	void Object::setPosition(const double& x, const double& y)
	{
		mX = x;
		mY = y;
	}
	
	unsigned int Object::GetId() const
	{
		return mId;
	}
	
	double Object::GetX() const
	{
		return mX;
	}
	
	double Object::GetY() const
	{
		return mY;
	}
	
	double Object::GetXSprite() const
	{
		return mX;
	}
	
	double Object::GetYSprite() const
	{
		return mY;
	}
	
	double Object::getWidth() const
	{
		return mWidth;
	}
	
	double Object::getHeight() const
	{
		return mHeight;
	}
	
	GameSystem::Classes::Names Object::GetClass() const
	{
		return mClass;
	}
	
	void Object::PerdHP(const double& amount)
	{
		mHP -= amount;
		if (mHP < 0)
			mHP = 0;
	}
}
