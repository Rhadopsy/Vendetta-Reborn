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
#ifndef MYVECTOR_HPP
#define MYVECTOR_HPP

#include <iostream>
#include <vector>
#include <stdexcept>

template <class T>
struct Vector
{
	Vector()
	{
	}
	
	unsigned int size() const
	{
		return vector.size();
	}
	
	T operator[](const unsigned int& id) const
	{
		return vector.at(id);
	}
	
	T& operator[](const unsigned int& id)
	{
		try
		{
			return vector.at(id);
		}
		catch (std::out_of_range& e)
		{
			unsigned int size = id + 1;
			//vector.reserve(2*size);
			vector.resize(size);
			return (*this)[id];
		}
	}
	
	void clear()
	{
		vector.clear();
	}
	
	void push_back(T element)
	{
		vector.push_back(element);
	}
	
	std::vector<T> vector;
};

#endif
