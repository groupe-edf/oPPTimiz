//oPPTimiz is a powerpoint addin that allow users to reduce presentations size.
//Copyright (C) 2025 EDF
//This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
//This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
//You should have received a copy of the GNU General Public License along with this program. If not, see https://www.gnu.org/licenses/.  

using System;
using System.Globalization;
using System.Reflection;

namespace oPPTimiz
{
    public class LanguageResources
    {
        public CultureInfo cultureInfo;
        public string strCulture;

        #region GENERAL TEXTS
        public readonly string formWindowTitle;
        public readonly string formDescription;
        public readonly string formGroupTitle;
        public readonly string formRadioOptimized;
        public readonly string formRadioStandard;
        public readonly string formCopyrigth;
        public readonly string formButtonOverride;
        public readonly string formButtonKeep;
        public readonly string formButtonCancel;
        public readonly string ribbonGroupTitle;
        public readonly string ribbonButtonOpptimize;
        public readonly string ribbonButtonOpptimizeTip;
        public readonly string ribbonButtonAccessibility;
        public readonly string shortkeyMaxCompression;
        public readonly string messageboxTitle;
        public readonly string messageSuccessOptimization;
        public readonly string messageFileGenerated;
        public readonly string messageInitialSize;
        public readonly string messageOptimizedSize;
        public readonly string messageOptimizationPercentage;
        public readonly string messageFileAlreadyOptimized;
        public readonly string messageError;
        #endregion

        public LanguageResources(CultureInfo culture)
        {
            cultureInfo = culture;

            switch (cultureInfo.TwoLetterISOLanguageName)
            {
                case "fr":
                    #region French culture elements
                    strCulture = "fr";

                    formWindowTitle = "oPPTimiz - Optimisation de fichier PowerPoint";
                    formDescription = $"Cet outil optimise un fichier PowerPoint en supprimant les {Environment.NewLine}masques inutilisés et en compressant les images contenues{Environment.NewLine}dans les diapositives et les masques.{Environment.NewLine}{Environment.NewLine}La présentation en cours est sauvegardée avant optimisation.";
                    formGroupTitle = "Niveau d'optimisation";
                    formRadioOptimized = $"Optimisé{Environment.NewLine}(Pour tous les usages standards)";
                    formRadioStandard = $"Standard{Environment.NewLine}(Pour la projection sur grand écran)";
                    formCopyrigth = $"v{Assembly.GetExecutingAssembly().GetName().Version}";
                    formButtonOverride = "Optimiser et remplacer";
                    formButtonKeep = "Optimiser et conserver";
                    formButtonCancel = "Annuler";
                    ribbonGroupTitle = "Numérique Responsable";
                    ribbonButtonOpptimize = "Optimiser";
                    ribbonButtonOpptimizeTip = "Optimiser la taille de la présentation";
                    ribbonButtonAccessibility = "Vérifier l'accessibilité";
                    shortkeyMaxCompression = "%{c}";
                    messageboxTitle = "Optimisation de la présentation";
                    messageSuccessOptimization = $"L'optimisation s'est terminée avec succès.";
                    messageFileGenerated = "Le fichier {0} a été généré.";
                    messageInitialSize = "Taille initiale";
                    messageOptimizedSize = "Taille optimisée";
                    messageOptimizationPercentage = "Soit une optimisation de";
                    messageFileAlreadyOptimized = "Félicitations, le fichier original était déjà correctement optimisé.";
                    messageError = "Erreur";
                    #endregion
                    break;

                default:
                    #region English culture elements
                    cultureInfo = new CultureInfo("en-US");
                    strCulture = "en";

                    formWindowTitle = "oPPTimiz - PowerPoint file optimization";
                    formDescription = $"This tool optimizes a PowerPoint file by deleting {Environment.NewLine}unused masks and by compressing pictures contained{Environment.NewLine}in slides and masks{Environment.NewLine}{Environment.NewLine}The current presentation is saved before optimization.";
                    formGroupTitle = "Optimization level";
                    formRadioOptimized = $"Optimized{Environment.NewLine}(For all standard uses)";
                    formRadioStandard = $"Standard{Environment.NewLine}(For large screen projection)";
                    formCopyrigth = $"v{Assembly.GetExecutingAssembly().GetName().Version}";
                    formButtonOverride = "Optimize && replace";
                    formButtonKeep = "Optimize && preserve";
                    formButtonCancel = "Cancel";
                    ribbonGroupTitle = "Digital Responsibility";
                    ribbonButtonOpptimize = "Optimize";
                    ribbonButtonOpptimizeTip = "Optimize presentation size";
                    ribbonButtonAccessibility = "Check accessibility";
                    shortkeyMaxCompression = "%{e}";
                    messageboxTitle = "Presentation optimization";
                    messageSuccessOptimization = $"Optimization was done successfully.{Environment.NewLine}";
                    messageFileGenerated = "File {0} has been created.";
                    messageInitialSize = "Initial size";
                    messageOptimizedSize = "Optimized size";
                    messageOptimizationPercentage = "This represents an optimization of";
                    messageFileAlreadyOptimized = "Congratulations, the original file was already properly optimized.";
                    messageError = "Error";
                    #endregion
                    break;
            }
        }
    }
}
