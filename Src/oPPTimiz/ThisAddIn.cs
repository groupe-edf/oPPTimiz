//oPPTimiz is a powerpoint addin that allow users to reduce presentations size.
//Copyright (C) 2025 EDF
//This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
//This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
//You should have received a copy of the GNU General Public License along with this program. If not, see https://www.gnu.org/licenses/.  

using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Threading;
using Microsoft.Win32;
using Microsoft.VisualBasic.FileIO;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;

namespace oPPTimiz
{
    /// <summary>
    /// Compression level for pictures
    /// </summary>
    public enum PictureCompressionLevel
    {
        Intermediate,
        Maximal
    }

    public partial class ThisAddIn
    {
        LanguageResources Resources;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Resources = new LanguageResources(Ribbon2.GetCulture());
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// Compress pictures through PictureCmpress MSO with selected compression level
        /// </summary>
        /// <param name="compressionLevel">Compression level for pictures</param>
        private void PictureCompress(PictureCompressionLevel compressionLevel)
        {
            Application.CommandBars.ExecuteMso(Constants.MsoPictureCompress);
            SendKeys.SendWait(Constants.KeysAltA);

            switch (compressionLevel)
            {
                case PictureCompressionLevel.Intermediate:
                    SendKeys.SendWait(Constants.KeysAltW);
                    break;
                case PictureCompressionLevel.Maximal:
                    //For M365, the shortcut depends on the configured language
                    if (Application.Version == Constants.Office16Version)
                        SendKeys.SendWait(Resources.shortkeyMaxCompression); 
                    SendKeys.SendWait(Constants.KeysAltM);
                    break;
            }

            SendKeys.SendWait(Constants.KeysEnter);
        }

        /// <summary>
        /// Returns file optimization summary
        /// </summary>
        /// <param name="initialSize">Initial size of file before optimization</param>
        /// <param name="newSize">New size of file after optimization</param>
        /// <param name="gain">Optimization gain</param>
        /// <param name="percentage">Gain percentage</param>
        /// <param name="threshold">Optimization threshold (optimization is saved only if percentage gain is greater than the threshold)</param>
        /// <param name="fileName">File generated name</param>
        /// <param name="keepFile">Indicates if initial file has been kept or overwritten</param>
        /// <returns></returns>
        private string GetEndMessage(long initialSize, long newSize, long gain, double percentage, int threshold, string fileName, bool keepFile)
        {
            string result = string.Empty;
            result = $"{Resources.messageSuccessOptimization}{Environment.NewLine}";

            if (percentage >= threshold)
            {
                if (keepFile)
                {
                    result += $"{string.Format(Resources.messageFileGenerated, fileName)}{Environment.NewLine}{Environment.NewLine}";
                }
                result += $"{Resources.messageInitialSize} : {GetSavedSize(initialSize)}{Environment.NewLine}";
                result += $"{Resources.messageOptimizedSize} : {GetSavedSize(newSize)}{Environment.NewLine}";
                result += $"{Resources.messageOptimizationPercentage} : {GetSavedSize(gain)} ({percentage}%)";
            }
            else
            {
                result += Resources.messageFileAlreadyOptimized;
            }

            return result;
        }

        /// <summary>
        /// Returns file optimization status
        /// </summary>
        /// <param name="fileName">File generated name</param>
        /// <param name="keepFile">Indicates if initial file has been kept or overwritten</param>
        /// <returns></returns>
        private string GetEndMessage(string fileName, bool keepFile)
        {
            string result = string.Empty;
            result = Resources.messageSuccessOptimization;
            if (keepFile)
            {
                result += $"{Environment.NewLine}{string.Format(Resources.messageFileGenerated, fileName)}{Environment.NewLine}{Environment.NewLine}";
            }
            return result;
        }

        /// <summary>
        /// Returns displayed string for saved size in octets
        /// </summary>
        /// <param name="octetsNb">octets saved (int)</param>
        /// <returns></returns>
        private string GetSavedSize(long octetsNb)
        {
            int threshold = 1024;
            long intermediate = 0;
            string result = string.Empty;

            if (octetsNb / threshold >= 1)
            {
                intermediate = octetsNb / threshold;
                if (intermediate / threshold >= 1)
                {
                    intermediate = intermediate / threshold;
                    if (intermediate / threshold >= 1)
                    {
                        intermediate = intermediate / threshold;
                        result = $"{intermediate} {Constants.UnitsGigaOctets}";
                    }
                    else
                    {
                        result = $"{intermediate} {Constants.UnitsMegaOctets}";
                    }
                }
                else
                {
                    result = $"{intermediate} {Constants.UnitsKiloOctets}";
                }
            }
            else
            {
                result = $"{octetsNb} {Constants.UnitsOctets}";
            }

            return result;
        }

        /// <summary>
        /// Retrives optimization percentage threshold from registry. Default value returned is 5%.
        /// </summary>
        /// <returns></returns>
        private int GetOptimizationThreshold()
        {
            string oPPTimizKey = Constants.RegKeyOpptimiz;
            int threshold = 5;

            try
            {
                RegistryKey regKeyoPPTimiz = Registry.CurrentUser.OpenSubKey(oPPTimizKey, false);
                threshold = (int)regKeyoPPTimiz.GetValue(Constants.RegValueThreshold);
            }
            catch
            {
                threshold = 5;
            }

            return threshold;
        }

        /// <summary>
        /// Optimizes file by removing unused masks & dispositions, compressing pictures and saving files
        /// </summary>
        /// <param name="compressionLevel">Compression level for pictures</param>
        /// <param name="keepOriginalFile">Indicates if initial file has been kept or overwritten</param>
        public void OptimizeFile(PictureCompressionLevel compressionLevel, bool keepOriginalFile = false)
        {
            try
            {
                PowerPoint.Presentation presentation = Application.ActivePresentation;

                #region Check if presentation is set to final and temporary disable it if needed
                bool bIsFinal = false;
                if (presentation.Final)
                {
                    presentation.Final = false;
                    bIsFinal = true;
                }
                #endregion

                presentation.Save();

                #region Retrieve file infos before optimization
                string fullPath = presentation.FullName;

                FileInfo fileinfoInitial = null;
                try
                {
                    fileinfoInitial = new FileInfo(fullPath);
                }
                catch { }
                long initialSize = fileinfoInitial.Length;
                #endregion

                #region Removal of unused masks and dispositions
                bool isUsed = false;
                for (int i = presentation.Designs.Count; i > 0; i--)
                {
                    for (int j = presentation.Designs[i].SlideMaster.CustomLayouts.Count; j > 0; j--)
                    {
                        try
                        {
                            presentation.Designs[i].SlideMaster.CustomLayouts[j].Delete();
                        }
                        catch
                        {
                            isUsed = true;
                        }
                    }

                    if (!isUsed)
                    {
                        try
                        {
                            presentation.Designs[i].SlideMaster.Delete();
                        }
                        catch
                        {

                        }
                    }
                    else
                    {
                        isUsed = false;
                    }
                }
                #endregion

                #region Picture compression
                //Add an image to allow compression
                PowerPoint.Shape shape = presentation.Slides[1].Shapes.AddPicture2($@"{AppDomain.CurrentDomain.BaseDirectory}\{Constants.PathToImageForCompression}", Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, 0, 0, -1, -1);
                Application.ActiveWindow.View.GotoSlide(1);
                shape.Select();

                PictureCompress(compressionLevel);
                #endregion

                (new Thread(() =>
                    {
                        string newPath = string.Empty;

                        if (keepOriginalFile)
                        {
                            #region Creation of a new document
                            string fileNameSuffix = string.Empty;

                            switch (compressionLevel)
                            {
                                case PictureCompressionLevel.Intermediate:
                                    fileNameSuffix = Constants.IntermediateOptimizedFilenameSuffix;
                                    break;
                                case PictureCompressionLevel.Maximal:
                                    fileNameSuffix = Constants.MaximumOptimizedFilenameSuffix;
                                    break;
                            }
                            string fileDir = Path.GetDirectoryName(fullPath);
                            string fileName = Path.GetFileNameWithoutExtension(fullPath);
                            string fileExt = Path.GetExtension(fullPath);
                            newPath = Path.Combine(fileDir, string.Concat(fileName, fileNameSuffix, fileExt));
                            #endregion
                        }
                        else
                        {
                            newPath = fullPath;
                        }

                        #region Restoring the PictureCompress interface (checkboxes)
                        Application.ActiveWindow.View.GotoSlide(1);
                        shape.Select();

                        PictureCompress(compressionLevel);
                        #endregion

                        #region Cleaning of the file and saving
                        shape.Delete();
                        if (keepOriginalFile)
                        {
                            presentation.SaveAs(newPath);
                        }
                        else
                        {
                            presentation.Save();
                        }
                        #endregion

                        #region Retrieve file infos after opptimization
                        FileInfo fileinfoNew = null;
                        try
                        {
                            fileinfoNew = new FileInfo(newPath);
                        }
                        catch { }
                        #endregion

                        #region Adding custom properties to the document
                        object customProperty;
                        customProperty = presentation.CustomDocumentProperties;

                        Type typeDocCustomProps = customProperty.GetType();

                        DateTime date = DateTime.Now;
                        long newSize = fileinfoNew.Length;
                        long gain = initialSize - newSize;
                        double percentage = Math.Ceiling((gain / (double)fileinfoInitial.Length) * 100);

                        List<CustomProperty> properties = new List<CustomProperty>
                        {
                            new CustomProperty(Constants.pptPropertyDate, date),
                            new CustomProperty(Constants.pptPropertyGain, gain),
                            new CustomProperty(Constants.pptPropertyRatio, percentage)
                        };

                        bool isSuccess = true;

                        foreach (CustomProperty oArg in properties)
                        {
                            try
                            {
                                //Add new property
                                object[] customPropertyArgs = { oArg.Name, false, oArg.Type, oArg.Value };
                                typeDocCustomProps.InvokeMember(Constants.pptPropertyMethodAdd, BindingFlags.Default | BindingFlags.InvokeMethod, null, customProperty, customPropertyArgs);
                            }
                            catch
                            {
                                isSuccess = false;
                            }
                            if (!isSuccess)
                            {
                                try
                                {
                                    //Update existing property
                                    object[] customPropertyArgs = { oArg.Name, oArg.Value };
                                    typeDocCustomProps.InvokeMember(Constants.pptPropertyMethodUpdate, BindingFlags.Default | BindingFlags.SetProperty, null, customProperty, customPropertyArgs);
                                }
                                catch
                                {
                                    isSuccess = false;
                                }
                            }

                            isSuccess = true;
                        }
                        #endregion

                        #region Reconfigure final state if needed
                        if (bIsFinal)
                        {
                            presentation.Final = true;
                        }
                        #endregion

                        #region Save file properties and final state
                        if (keepOriginalFile)
                        {
                            presentation.SaveAs(newPath);
                        }
                        else
                        {
                            presentation.SaveAs(fullPath);
                        }
                        #endregion

                        #region Show opptimization results to user
                        try
                        {
                            int thresholdOptimization = GetOptimizationThreshold();
                            if (percentage < thresholdOptimization)
                            {
                                FileSystem.DeleteFile(newPath);
                            }

                            MessageBox.Show(GetEndMessage(fileinfoInitial.Length, fileinfoNew.Length, gain, percentage, thresholdOptimization, newPath, keepOriginalFile), 
                                Resources.messageboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch
                        {
                            //Powerpoint is not present locally (sharepoint, onedrive ...)
                            MessageBox.Show(GetEndMessage(newPath, keepOriginalFile), Resources.messageboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        #endregion

                        Dispose();
                    })).Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), Resources.messageboxTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Executes the MSO for the accessibility checker by Microsoft
        /// </summary>
        public void ExecuteGroupAccessibilityMSO()
        {
            try
            {
                (new Thread(() =>
                {
                    Application.CommandBars.ExecuteMso(Constants.MsoAccessibilityChecker);
                })).Start();
            }
            catch
            {
                MessageBox.Show(Resources.messageError);
            }
        }

        #region Code généré par VSTO

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
