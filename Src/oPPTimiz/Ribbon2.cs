//oPPTimiz is a powerpoint addin that allow users to reduce presentations size.
//Copyright (C) 2025 EDF
//This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
//This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
//You should have received a copy of the GNU General Public License along with this program. If not, see https://www.gnu.org/licenses/.  
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Tools.Ribbon;

namespace oPPTimiz
{
    public partial class Ribbon2
    {
        public static CultureInfo GetCulture()
        {
            return CultureInfo.CurrentCulture;
        }

        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {
            LanguageResources resources = new LanguageResources(GetCulture());

            group1.Label = resources.ribbonGroupTitle;
            Optimiser.Label = resources.ribbonButtonOpptimize;
            Optimiser.ScreenTip = resources.ribbonButtonOpptimizeTip;
            button1.Label = resources.ribbonButtonAccessibility;
        }

        private void ClickOptimize(object sender, RibbonControlEventArgs e)
        {
            (new Thread(() =>
            {
                Form1 dialog = new Form1(new LanguageResources(GetCulture()));
                dialog.ShowDialog();
            })).Start();
        }

        private void StartGroupAccessibility(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExecuteGroupAccessibilityMSO();
        }
    }
}
