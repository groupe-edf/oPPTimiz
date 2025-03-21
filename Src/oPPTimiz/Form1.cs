//oPPTimiz is a powerpoint addin that allow users to reduce presentations size.
//Copyright (C) 2025 EDF
//This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
//This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
//You should have received a copy of the GNU General Public License along with this program. If not, see https://www.gnu.org/licenses/.  

using System;
using System.Linq;
using System.Windows.Forms;

namespace oPPTimiz
{
    public partial class Form1 : Form
    {
        public Form1(LanguageResources ressources)
        {
            InitializeComponent();
            Text = ressources.formWindowTitle;
            labelVersion.Text = ressources.formCopyrigth;
            labelFormDescription.Text = ressources.formDescription;
            groupBoxLevel.Text = ressources.formGroupTitle;
            radioMaximal.Text = ressources.formRadioOptimized;
            radioIntermediate.Text = ressources.formRadioStandard;
            validateOverrideButton.Text = ressources.formButtonOverride;
            validateKeepButton.Text = ressources.formButtonKeep;
            cancelButton.Text = ressources.formButtonCancel;
        }

        private void startOptimization(object sender, EventArgs e)
        {
            RadioButton selected = this.groupBoxLevel.Controls.OfType<RadioButton>().FirstOrDefault(radioButton1 => radioButton1.Checked);

            PictureCompressionLevel compressionLevel;
            bool keepFile = false;
            switch (selected.Name)
            {
                case "radioIntermediaire":
                    compressionLevel = PictureCompressionLevel.Intermediate;
                    break;
                default:
                    compressionLevel = PictureCompressionLevel.Maximal;
                    break;
            }
            if((Button)sender == validateKeepButton)
            {
                keepFile = true;
            }
            this.Hide();
            this.Close();

            Globals.ThisAddIn.OptimizeFile(compressionLevel, keepFile);
        }

        private void CancelOptimization(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
