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
#include <filesystem>
#include <iostream>

#include "Configuration.hpp"
#include "Interface/Interface.hpp"

ConfigurationClass* Configuration;

void Load(void* ptr)
{
	InterfaceModule::Interface* interface = static_cast<InterfaceModule::Interface*>(ptr);
	Configuration = new ConfigurationClass();
	interface->Configure( Configuration->Configure() );
	
	TiXmlDocument* sauvegarde = new TiXmlDocument();
	interface->Load(sauvegarde->LoadFile("./Data/Saves/test.xml", TIXML_ENCODING_UTF8) ? sauvegarde->FirstChildElement("Save") : NULL);
}

int main(int, char** argv)
{
	try
	{
		const std::filesystem::path executable = std::filesystem::absolute(argv[0]);
		if (executable.has_parent_path())
			std::filesystem::current_path(executable.parent_path());
	}
	catch (const std::filesystem::filesystem_error& error)
	{
		std::cerr << "Impossible d'initialiser le repertoire du jeu: " << error.what() << "\n";
		return EXIT_FAILURE;
	}
	
	InterfaceModule::Interface* interface = new InterfaceModule::Interface();
	//sf::Thread threadLoadInterface(&Load, interface);
	//threadLoadInterface.Launch();
	Load(interface);
	
	sf::Clock frameClock;
	while (interface->isOpen())
	{
		frameClock.restart();
		interface->draw();
		while (interface->CatchEvent());
		interface->Process(frameClock.getElapsedTime().asSeconds());
	}
	
	//threadLoadInterface.Terminate();
	
	delete interface;
	delete Configuration;
	
	return EXIT_SUCCESS;
}
