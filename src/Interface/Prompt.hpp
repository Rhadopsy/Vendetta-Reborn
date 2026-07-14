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
#ifndef PROMPT_HPP
#define PROMPT_HPP

#include "../compat/SfmlLegacy.hpp"

#include "../GameSystem/Character.hpp"

namespace InterfaceModule
{
	class Prompt
	{
		public:
			// txt = Text d'invite
			Prompt(const std::string& txt);
			~Prompt();
			
			void draw(sf::RenderTarget* const render);
			bool Catch(const sf::Event& Event);
			void AddCharacter(const sf::Uint32& text);
			
			void setPosition(const double& x, const double& y);
			
			void ReturnTo(GameSystem::Character* const objet);
		
		private:
			bool mShow;
			bool mActive;
			
			double mX;
			double mY;
			
			legacy::Text* mPromptText;
			legacy::Text* mCurrentText;
			
			bool mInsertionActive;
			unsigned int mCurrentCaracter;
			legacy::Text* mSelectCaracter;
			
			GameSystem::Character* mTarget;
	};
}

#endif
