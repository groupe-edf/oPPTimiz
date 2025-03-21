#oPPTimiz is a powerpoint addin that allow users to reduce presentations size.
#Copyright (C) 2025 EDF
#This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
#This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
#You should have received a copy of the GNU General Public License along with this program. If not, see https://www.gnu.org/licenses/.  

param (
    [Parameter(Mandatory = $true)]
    [string]$pptFile,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Maximal", "Intermediate")]
    [string]$compressionLevel="Maximal",

    [int]$keepFile=0
)

#region Constants 
$sScriptName = "oPPTimiz"
$sScriptVersion = "4.0"
$sLogDirectory = [string]::Format("{0}\Souche\Logs\", $env:PROGRAMDATA)
$script:sLogFile = ""
$ErrorActionPreference = "SilentlyContinue"
$iMaxLogSize = 1

$RegKeyOpptimiz = "HKCU:\SOFTWARE\oPPTimiz"
$RegValueThreshold = "OptimizationThreshold"
$DefaultOptimizationThreshold = 5

$pptPropertyDate = "oPPTimizDate"
$pptPropertyGain= "oPPTimizGain"
$pptPropertyRatio = "oPPTimizRatio"
$pptPropertyMethodAdd = "Add"
$pptPropertyMethodUpdate= "Item"

$pptProcessName = "POWERPNT"

$PathToImageForCompression = "Resources\220ppi.png"
#endregion

#region Localization
$messageUnableToOptimizeEN = "Unable to optimize this presentation at the moment.`n`nPlease try again later."
$messageUnableToOptimizeFR = "Impossible d'optimiser cette présentation actuellement.`n`nVeuillez réessayer ultérieurement."
$messagePowerPointRunningEN = "PowerPoint is currently running.`n`nClick ""OK"" to save the current presentation(s) and start optimizing or click ""Cancel"" to cancel the operation."
$messagePowerPointRunningFR = "PowerPoint est actuellement en cours d'utilisation.`n`nCliquez sur ""OK"" pour sauvegarder la/les présentation(s) en cours et lancer l'optimisation ou cliquez sur ""Annuler"" pour annuler l'opération."
$keysMaxCompressionM365EN = "%{e}"
$keysMaxCompressionM365FR = "%{c}"
#endregion

#region Helper Logs
#Writes informations of execution context of the script to the log file
function Start-Log
{
	$sLogCustomDir = ""
	if ($env:USERNAME -match $env:COMPUTERNAME)
	{
		$sLogCustomDir = "System"
	}
	else
	{
		$sLogCustomDir = "$($env:USERNAME)"
	}
	
	if (($sLogCustomDir -ne "") -and (-not (Test-Path -Path "$sLogDirectory\$sLogCustomDir")))
	{
		New-Item -ItemType directory -Path $sLogDirectory -Name $sLogCustomDir -ErrorAction SilentlyContinue | Out-Null
	}
	
	$script:sLogFile = [string]::Format("{0}{1}\{2}-{3}.log", $sLogDirectory, $sLogCustomDir, $sScriptName, $env:COMPUTERNAME)
	
	Write-Log -Message "***************************************************************************************************"
	Write-Log -Message "Running script $sScriptName version [$sScriptVersion]."
	Write-Log -Message "pptFile : $pptFile"
	Write-Log -Message "CompressionLevel : $compressionLevel"
	Write-Log -Message "KeepFile : $keepFile"
	Write-Log -Message "---------------------------------------------------------------------------------------------------"
}

#Adds an entry to the file log (with log rotation mechanism)
function Write-Log
{
	param ([string]$Message,
		[string]$Severity = "Info")
	
	if ((Get-Item $script:sLogFile).length/1MB -ge $iMaxLogSize)
	{
		$sLogFileOld = "$($script:sLogFile).old"
		if ((Test-Path $sLogFileOld))
		{
			Remove-Item $sLogFileOld
		}
		Move-Item -Path $script:sLogFile -Destination $sLogFileOld
	}
	
	if (!(Test-Path $script:sLogFile))
	{
		New-Item -ItemType file -Path $script:sLogFile -ErrorAction SilentlyContinue | Out-Null
		if (!$?)
		{
			Write-Host -ForegroundColor Red "`t => ERROR : Log File cannot be created"
			Write-Host -ForegroundColor Red "=> End of script"
			Exit 100
		}
	}
	
	switch ($Severity)
	{
		"Info" { Write-Host $Message; break }
		"Success" { Write-Host $Message -ForegroundColor Green; break }
		"Warning" { Write-Host $Message -ForegroundColor Yellow; break }
		"Error" { Write-Host $Message -ForegroundColor Red; break }
		"Title" { Write-Host $Message -ForegroundColor Cyan; break }
	}
	
	[string]$sMessageLog = [string]::Format("{0} - {1}", (Get-Date -UFormat "%d/%m/%Y %T").ToString(), $Message.Replace("`t", "     "))
	Add-Content -Path $script:sLogFile $sMessageLog -ErrorAction SilentlyContinue | Out-Null
	if (!$?)
	{
		Write-Host -ForegroundColor Red "`t => ERROR : Log File cannot be modified"
		Write-Host -ForegroundColor Red "=> End of script"
		Exit 101
	}
}

#Logs end of script execution with exit code
Function Exit-Script
{
	param ([int]$ExitCode)
	$oMtx.ReleaseMutex()
	Write-Log -Message "Exiting script with exit code : $ExitCode"
	Write-Log -Message "***************************************************************************************************"
	Exit $ExitCode
}
#endregion

#region Dll imports to force interface to foreground
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class WinAp {
      [DllImport("user32.dll")]
      [return: MarshalAs(UnmanagedType.Bool)]
      public static extern bool SetForegroundWindow(IntPtr hWnd);

      [DllImport("user32.dll")]
      [return: MarshalAs(UnmanagedType.Bool)]
      public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
"@

Function Set-ForegroundProcess($process, $application)
{
	$h = $process.MainWindowHandle
	$application.Activate()
	[void][WinAp]::SetForegroundWindow($h)
	[void][WinAp]::ShowWindow($h, 3)
	[Microsoft.VisualBasic.Interaction]::AppActivate($process.ID)
}
#endregion

#region Helper oPPTimiz
#Returns displayed string for saved size in octets
Function Get-SizeSaved([int]$octetsNb) {
    [int]$palier = 1024
    [int]$intermediate = 0
    $result = ""

    if ($octetsNb / $palier -ge 1) {
        [int]$intermediate = $octetsNb / $palier
        if ($intermediate / $palier -ge 1) {
            [int]$intermediate = $intermediate / $palier;
            if ($intermediate / $palier -ge 1){
                [int]$intermediate = $intermediate / $palier;
                $result = "$intermediate Go";
            } else {
                $result = "$intermediate Mo";
            }
        } else {
            $result = "$intermediate Ko";
        }
    } else {
        $result = "$octetsNb octets";
    }
    return $result
}

#Retrieves the configured optimization percentage threshold
Function Get-OptimizationThreshold() {
    $threshold = $DefaultOptimizationThreshold

    try {
        $threshold = (Get-ItemProperty -Path $RegKeyOpptimiz -Name $RegValueThreshold -ErrorAction Stop).$RegValueThreshold
    } catch {
        $threshold = $DefaultOptimizationThreshold
    }
    return $threshold
}

#Compresses all images in the presentation
Function Compress-Picture($Application)
{
	$cmdBars = $Application.CommandBars
	$cmdBars.ExecuteMso("PicturesCompress")
	Start-Sleep -s 2
	
	$OSlanguage = (Get-WinSystemLocale).LCID

	Switch ($compressionLevel) 
	{
		"Maximal" {
			If ($Application.Version -eq "16.0") {
				if ($OSlanguage -eq 1036)
				{
					[System.Windows.Forms.SendKeys]::SendWait("%{a}%{c}$keysMaxCompressionM365FR%{m}{ENTER}")
				}
				else
				{
					[System.Windows.Forms.SendKeys]::SendWait("%{a}$keysMaxCompressionM365EN%{m}{ENTER}")
				}
			}
			else
			{
				[System.Windows.Forms.SendKeys]::SendWait("%{a}%{m}{ENTER}")
			}
		}

		"Intermediate" {
			[System.Windows.Forms.SendKeys]::SendWait("%{a}%{w}{ENTER}")
		}
	}
	[System.Windows.Forms.SendKeys]::Flush()
}

#Deletes all unused designs in the presentation
Function Remove-UnusedDesigns($Presentation)
{
	$isUsed = $false;
	for (($i = $Presentation.Designs.Count); ($i -gt 0); ($i--))
	{
		for (($j = $Presentation.Designs[$i].SlideMaster.CustomLayouts.Count); ($j -gt 0); ($j--))
		{
			try
			{
				$Presentation.Designs[$i].SlideMaster.CustomLayouts[$j].Delete();
			}
			catch
			{
				$isUsed = $true;
			}
		}

		if (-not ($isUsed))
		{
			try
			{
				$Presentation.Designs[$i].SlideMaster.Delete();
			}
			catch
			{

			}
		}
		else
		{
			$isUsed = $false;
		}
	}
}

#Adds oPPTimiz PowerPoint properties to the presentation file
Function Add-OpptimizFileProperty($Presentation, $date, $gain, $ratio)
{
	$customProperties = $Presentation.CustomDocumentProperties
    
    $CustomPropertiesToAdd = @(
        [pscustomobject]@{Name=$pptPropertyDate;Value=$date;Type=[Microsoft.Office.Core.MsoDocProperties]::msoPropertyTypeDate}
        [pscustomobject]@{Name=$pptPropertyGain;Value=$gain;Type=[Microsoft.Office.Core.MsoDocProperties]::msoPropertyTypeNumber}
        [pscustomobject]@{Name=$pptPropertyRatio;Value=$ratio;Type=[Microsoft.Office.Core.MsoDocProperties]::msoPropertyTypeNumber}
    )
            
    $CustomPropertiesToAdd | ForEach-Object{
		$isAdded = $false
		
        try
	    {
		    [System.__ComObject].InvokeMember($pptPropertyMethodAdd, ([System.Reflection.BindingFlags]::Default -bor [System.Reflection.BindingFlags]::InvokeMethod), $null, $customProperties, @($_.Name, $false, $_.Type, $_.Value)) | out-null
		    $isAdded = $true
	    } 
	    catch
	    {

	    }
	
	    if($isAdded -eq $false)
	    {
		    try
		    {
			    [System.__ComObject].InvokeMember($pptPropertyMethodUpdate, ([System.Reflection.BindingFlags]::Default -bor [System.Reflection.BindingFlags]::SetProperty), $null, $customProperties, @($_.Name, $_.Value))
		    }
		    catch
		    {

		    }
	    }
    }
}

#endregion

#region Main
[void][System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")

#Instanciate a MUTEX to avoid multiple process
$oMtx = New-Object System.Threading.Mutex($false, "$($sScriptName)-$($env:COMPUTERNAME)")

try
{
	#Wait for all previous executions to stop
	if ($oMtx.WaitOne())
	{
		Start-Log

        #region Handle previous running PowerPoint instances
        $currentProcesses = Get-Process | Where-Object { $_.Name -eq $pptProcessName }

        if ($currentProcesses -ne $Null)
        {
            Add-Type -AssemblyName PresentationCore,PresentationFramework
                
            $OSlanguage = (Get-WinSystemLocale).LCID
            
			$isApplicationActive = $false
            foreach ($currentProcess in $currentProcesses)
            { 
                if ( ($currentProcess.Path -eq $Null) -and ($currentProcess.Handle -eq $Null) )
                {
                    if ($OSlanguage -eq 1036)
                    {
                        $result = [System.Windows.MessageBox]::Show($messageUnableToOptimizeFR, $sScriptName,0,48)
                    }
                    else
                    {
                        $result = [System.Windows.MessageBox]::Show($messageUnableToOptimizeEN, $sScriptName,0,48)
                    }

                    Write-Log -Message "A PowerPoint instance is currently running as administrator. Exiting script..." -Severity Warning

                    Exit-Script -ExitCode 0
                }
                
                if ( ($currentProcess.Path -ne $Null) -and ($currentProcess.Handle -ne $Null) )
                {                    
                    Write-Log -Message "PowerPoint is currently running."

                    if ($OSlanguage -eq 1036)
                    {
                        $result = [System.Windows.MessageBox]::Show($messagePowerPointRunningEN, $sScriptName,1,48)
                    }
                    else
                    {
                        $result = [System.Windows.MessageBox]::Show($messagePowerPointRunningEN, $sScriptName,1,48)
                    }

                    switch($result)
                    {
                        "OK" {
                            
                            $currentApplication = New-Object -ComObject powerpoint.application
                            $currentPresentations = $currentApplication.Presentations
                            $isApplicationActive = $true
                            
                            foreach ($currentPresentation in $currentPresentations)
                            {
                                $currentPresentation.Save()
                                $currentPresentation.Close()
                            }

                            Write-Log -Message "`t => Current presentations saved and closed" -Severity Success

                        }
                            
                        "Cancel" {

                            Write-Log -Message "`t => User input : optimization cancelled"

                            Exit-Script -ExitCode 0
                        }
                    }
                }
            }
        }
		#endregion
		
        #region Retrieve file size before optimization
		$fileInfoOld = [System.IO.FileInfo]::new($pptFile)
		$oldSize = $fileInfoOld.Length
		#endregion

		#region Starting powerpoint
		Write-Log -Message "Launching PowerPoint..."
        if ($isApplicationActive -eq $false)
        {
            $application = New-Object -ComObject powerpoint.application
        }
        else
        {
            $application = $currentApplication
        }
		$presentation = $application.Presentations.open($pptFile)
		
		$process = Get-Process | Where-Object { $_.Name -eq $processName }
        
        #Check if presentation is set to final and temporary disable it if needed
        $Isfinale = $false
        if ($presentation.Final -eq $true)
        {
           $presentation.Final = $false
           $IsFinale = $true 
        }
		
		#Force PowerPoint to foreground
        Set-ForegroundProcess -process $process -application $application
		Write-Log -Message "`t => PowerPoint started" -Severity Success
		#endregion
		
		#region Optimization of file
		Write-Log -Message "Removing unused designs..."
		Remove-UnusedDesigns -Presentation $presentation
		Write-Log -Message "`t => Unused designs removed" -Severity Success
		
		Write-Log -Message "Compressing pictures..."
		$shapeSource = (Get-Item -Path "$(Split-Path $hostinvocation.MyCommand.path)\$PathToImageForCompression").FullName
		$shape = $presentation.Slides[1].Shapes.AddPicture2($shapeSource, $false, $true, 0, 0, -1, -1)
		$shape.Select()
		Compress-Picture -Application $application
		$shape.Delete()
		Write-Log -Message "`t => Pictures compressed successfully" -Severity Success
		#endregion
		
		#region Calculate compression ratio and add properties to file
		$fullPath = $pptFile
		$fileDir = [System.IO.Path]::GetDirectoryName($fullPath)
		$fileName = [System.IO.Path]::GetFileNameWithoutExtension($fullPath)
		$fileExt = [System.IO.Path]::GetExtension($fullPath)
		$fileNameSuffix = "_oPPTimiz"
		if ($keepFile -eq 1)
		{
			$newPath = [System.IO.Path]::Combine($fileDir, "$fileName$fileNameSuffix$fileExt")
		}
		else
		{
			$newPath = [System.IO.Path]::Combine($fileDir, "$fileName$fileExt")
		}
		
		#Save file to new location
		Write-Log -Message "Saving file..."
		$presentation.SaveAs($newPath)
		Write-Log -Message "`t => File saved to $newPath" -Severity Success
		
		#Calculate realized gain
		Write-Log -Message "Computing optimization ratio..."
		$fileInfoNew = [System.IO.FileInfo]::new($newPath)
		
		$newSize = $fileInfoNew.Length
		Write-Log -Message "`t $newPath : $newSize"
		Write-Log -Message "`t $fullPath : $oldSize"
		$gain = $oldsize - $newSize
		$percentage = [System.Math]::Ceiling(($gain / $oldSize) * 100)
		$thresholdOptimization = Get-OptimizationThreshold
		Write-Log -Message "`t => Optimization threshold : $thresholdOptimization" -Severity Success
		Write-Log -Message "`t => Optimization ratio : $percentage" -Severity Success
		
        Write-Log -Message "Adding properties to file..."
		#Adding optimisation informations in file properties
		$date = (Get-Date)
        Add-OpptimizFileProperty -Presentation $presentation -date $date -gain $gain -ratio $percentage
        $presentation.SaveAs($newPath)
        Write-Log -Message "`t => File properties added" -Severity Success
        #endregion
        
        #region Set final state if needed
        If ($IsFinale -eq $true)
        {
            $presentation.final = $true
        }
		#endregion

		#region Closing PowerPoint
		Write-Log -Message "Closing PowerPoint..."
		if ($application.Presentations.Count -eq 1)
		{
			$presentation.Close()
			$process.Kill()
		}
		Write-Log -Message "`t => PowerPoint closed" -Severity Success
		#endregion
		
		#region Delete the file if it is less than optimization threshold
		if (($percentage -lt $thresholdOptimization) -and ($keepFile -eq 1))
		{
			Remove-Item $newPath
			Write-Log -Message "`t => File already opptimized" -Severity Warning
		}
		#endregion
		
		Exit-Script -ExitCode 0
	}
}
catch
{
	Write-Log -Message "`t => Error : $_" -Severity Error
	Exit-Script -ExitCode 1
}
#endregion