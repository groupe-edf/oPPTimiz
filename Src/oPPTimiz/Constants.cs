//oPPTimiz is a powerpoint addin that allow users to reduce presentations size.
//Copyright (C) 2025 EDF
//This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
//This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
//You should have received a copy of the GNU General Public License along with this program. If not, see https://www.gnu.org/licenses/.  

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace oPPTimiz
{
    internal static class Constants
    {
        public const string MsoPictureCompress = "PicturesCompress";
        public const string MsoAccessibilityChecker = "AccessibilityChecker";

        public const string KeysAltA = "%{a}";
        public const string KeysAltW = "%{w}";
        public const string KeysAltM = "%{m}";
        public const string KeysEnter = "{ENTER}";

        public const string IntermediateOptimizedFilenameSuffix = "_oPPTimiz";
        public const string MaximumOptimizedFilenameSuffix = "_oPPTimiz";

        public const string Office16Version = "16.0";

        public const string UnitsGigaOctets = "Go";
        public const string UnitsMegaOctets = "Mo";
        public const string UnitsKiloOctets = "Ko";
        public const string UnitsOctets = "octets";

        public const string RegKeyOpptimiz = @"Software\oPPTimiz";
        public const string RegValueThreshold = "OptimizationThreshold";

        public const string PathToImageForCompression = @"Resources\220ppi.png";

        public const string pptPropertyDate = "oPPTimizDate";
        public const string pptPropertyGain = "oPPTimizGain";
        public const string pptPropertyRatio = "oPPTimizRatio";
        public const string pptPropertyMethodAdd = "Add";
        public const string pptPropertyMethodUpdate = "Item";
    }
}
