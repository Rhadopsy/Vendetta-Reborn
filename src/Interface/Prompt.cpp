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
#include "Prompt.hpp"

namespace InterfaceModule
{
	Prompt::Prompt(const std::string& txt) :
		mShow(false),
		mActive(false),
		mX(0),
		mY(0),
		mPromptText(new legacy::Text(txt)),
		mCurrentText(new legacy::Text("")),
		mInsertionActive(true),
		mCurrentCaracter(0),
		mSelectCaracter(new legacy::Text("|")),
		mTarget(NULL)
	{
		mPromptText->setSize(20);
		mCurrentText->setSize(20);
		mSelectCaracter->setSize(20);
	}
	
	Prompt::~Prompt()
	{
		delete mPromptText;
		delete mCurrentText;
		delete mSelectCaracter;
	}
	
	void Prompt::draw(sf::RenderTarget* const render)
	{
		if (mShow)
		{
			mPromptText->setPosition(mX, mY);
			render->draw(*mPromptText);
			
			mCurrentText->setPosition(mX + mPromptText->getRect().width, mY);
			render->draw(*mCurrentText);
			
			if (mActive)
			{
				legacy::Text precedingText(mCurrentText->getText().substring(0, mCurrentCaracter));
				precedingText.setSize(20);
				mSelectCaracter->setPosition(mX + mPromptText->getRect().width + precedingText.getRect().width, mY);
				render->draw(*mSelectCaracter);
			}
		}
	}
	
	bool Prompt::Catch(const sf::Event& Event)
	{
		if (mActive)
		{
			sf::String text = mCurrentText->getText();
			
			if (Event.type == sf::Event::TextEntered)
			{
				if (Event.text.unicode == '\r') // Return
				{
					if (mTarget)
						mTarget->ReceiveString(text);
					
					mActive = false;
					mShow = false;
					mCurrentText->setText("");
					mCurrentCaracter = 0;
				}
				else if (Event.text.unicode == '\b') // Back
				{
					if (mCurrentCaracter > 0)
					{
						mCurrentText->setText(text.substring(0, mCurrentCaracter - 1) + text.substring(mCurrentCaracter, text.getSize() - mCurrentCaracter));
						mCurrentCaracter--;
					}
				}
				else if (Event.text.unicode == 127) // Suppr
				{
					if (mCurrentCaracter < text.getSize())
						mCurrentText->setText(text.substring(0, mCurrentCaracter) + text.substring(mCurrentCaracter + 1, text.getSize() - mCurrentCaracter));
				}
				else
					AddCharacter(Event.text.unicode);
			}
			else if (Event.type == sf::Event::KeyPressed)
			{
				if (Event.key.code == sf::Keyboard::Insert)
				{
					if (mInsertionActive)
					{
						mSelectCaracter->setText("_");
						mInsertionActive = false;
					}
					else
					{
						mSelectCaracter->setText("|");
						mInsertionActive = true;
					}
				}
				else if (Event.key.code == sf::Keyboard::Up)
					mCurrentCaracter = 0;
				else if (Event.key.code == sf::Keyboard::Down)
					mCurrentCaracter = text.getSize();
				else if (Event.key.code == sf::Keyboard::Left)
				{
					if (mCurrentCaracter > 0)
						mCurrentCaracter--;
				}
				else if (Event.key.code == sf::Keyboard::Right)
				{
					if (mCurrentCaracter < text.getSize())
						mCurrentCaracter++;
				}
			}
			
			return true;
		}
		else if (Event.type == sf::Event::TextEntered && Event.text.unicode == '\r')
		{
			mShow = true;
			mActive = true;
			return true;
		}
		
		return false;
	}
	
	void Prompt::AddCharacter(const sf::Uint32& car)
	{
		sf::String ancienText = mCurrentText->getText();
		mCurrentText->setText(ancienText + car);
		
		if (mCurrentCaracter < ancienText.getSize())
			mCurrentText->setText(ancienText.substring(0, mCurrentCaracter) + car + ancienText.substring(mCurrentCaracter + static_cast<int>(!mInsertionActive), ancienText.getSize() - mCurrentCaracter));
		else
			mCurrentText->setText(ancienText + car);
		
		mCurrentCaracter++;
	}
	
	void Prompt::setPosition(const double& x, const double& y)
	{
		mX = x;
		mY = y;
	}
	
	void Prompt::ReturnTo(GameSystem::Character* const objet)
	{
		mTarget = objet;
	}
}
