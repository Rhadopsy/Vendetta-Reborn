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
#ifndef SCRIPTS_HPP
#define SCRIPTS_HPP

#include <boost/python.hpp>

#include "../functions.hpp"
#include "Object.hpp"
#include "Building.hpp"
#include "Ressource.hpp"
#include "TypeBuilding.hpp"

#define PTR boost::python::return_value_policy<boost::python::reference_existing_object>()

#define WRAPER(CLASS) struct CLASS##Wrap : CLASS, boost::python::wrapper<CLASS> {
#define REGISTER(fun, def, param) void fun def { if (boost::python::override fct = this->get_override(#fun)) { try{fct param;}catch(...){PyErr_Print();}} }
#define END_WRAPER };

namespace GameSystem
{
	WRAPER(Character)
		REGISTER(onWhenever, 		(), 				())
		REGISTER(onSpwan, 		(), 				())
		REGISTER(onAttacked,		(GameSystem::Object* object), 	(boost::python::ptr(object)))
		REGISTER(onTargetSelected,	(GameSystem::Object* object),	(boost::python::ptr(object)))
		REGISTER(onGoCoords,		(double x, double y), 		(x, y))
		REGISTER(onGoObject,		(GameSystem::Object* object), 	(boost::python::ptr(object)))
	END_WRAPER
}

BOOST_PYTHON_MODULE(Vendetta)
{
	boost::python::enum_<GameSystem::Classes::Names>("Types")
		.value("Object",	GameSystem::Classes::Object)
		.value("Character",	GameSystem::Classes::Character)
		.value("Ressource",	GameSystem::Classes::Ressource)
		.value("Building",	GameSystem::Classes::Building)
		.value("Item",		GameSystem::Classes::Item)
	;
	boost::python::class_<Types::Building>("Building")
		.def("GetMaxHP",			&Types::Building::GetMaxHP)
		.def("GetRequisiteMaterialCount",	&Types::Building::GetRequisiteMaterialCount)
		.def("GetRequisiteMaterialId",		&Types::Building::GetRequisiteMaterialId)
		.def("GetRequisiteMaterialAmount",	&Types::Building::GetRequisiteMaterialAmount)
	;
	boost::python::class_<GameSystem::Object>("Object")
		.def("GetId",		&GameSystem::Object::GetId)
		.def("GetClass",	&GameSystem::Object::GetClass)
	;
	boost::python::class_<GameSystem::Ressource, boost::python::bases<GameSystem::Object> >("Ressource");
	boost::python::class_<GameSystem::CharacterWrap, /*std::auto_ptr<GameSystem::CharacterWrap>, */boost::noncopyable, boost::python::bases<GameSystem::Object> >("Character_Base")
		.def("ResetActions",	&GameSystem::Character::ResetActions)
		.def("GetRessource",	&GameSystem::Character::GetRessource, PTR)
		.def("SetSay",		&GameSystem::Character::SetSay)
		.def("Say",		&GameSystem::Character::Say)
		.def("Wait",		&GameSystem::Character::Wait)
		.def("Say",		&GameSystem::Character::Say)
		.def("GoCoords",	&GameSystem::Character::GoCoords)
		.def("GoObject",	&GameSystem::Character::GoObject)
		.def("Enter",		&GameSystem::Character::Enter)
		.def("Work",		&GameSystem::Character::Work, (boost::python::arg("target"), boost::python::arg("amount")=0) )
		.def("Attack",		&GameSystem::Character::Attack)
		.def("SeatBuilding",	&GameSystem::Character::SeatBuilding, PTR)
		.def("isActive",	&GameSystem::Character::isActive)
		.def("isGoing",		&GameSystem::Character::isGoing)
		.def("isAttacking",	&GameSystem::Character::isAttacking)
	;
	boost::python::class_<GameSystem::Building, boost::python::bases<GameSystem::Object> >("Building")
		.def("GetType",		&GameSystem::Building::GetType, PTR)
	;
	
	boost::python::def("Range", (double(*)(GameSystem::Object const* const, GameSystem::Object const* const))&Range);
}

#endif
