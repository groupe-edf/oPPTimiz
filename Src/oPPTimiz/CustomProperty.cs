//oPPTimiz is a powerpoint addin that allow users to reduce presentations size.
//Copyright (C) 2025 EDF
//This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
//This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
//You should have received a copy of the GNU General Public License along with this program. If not, see https://www.gnu.org/licenses/.  

using Microsoft.Office.Core;
using System;

namespace oPPTimiz
{
    class CustomProperty
    {
        public string Name;
        public object Value;
        public MsoDocProperties Type;

        public CustomProperty(string propertyName, DateTime value)
        {
            Name = propertyName;
            Value = value;
            Type = MsoDocProperties.msoPropertyTypeDate;
        }
        public CustomProperty(string propertyName, long value)
        {
            Name = propertyName;
            Value = value;
            Type = MsoDocProperties.msoPropertyTypeNumber;
        }
        public CustomProperty(string propertyName, double value)
        {
            Name = propertyName;
            Value = value;
            Type = MsoDocProperties.msoPropertyTypeNumber;
        }
    }
}
