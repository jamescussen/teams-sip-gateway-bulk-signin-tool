########################################################################
# Name: Teams SIP Gateway Bulk Sign in Tool
# Version: v1.0.0 (15/6/2024)
# Original Release Date: 15/6/2024
# Created By: James Cussen
# Web Site: http://www.myteamslab.com
# Notes: This is a PowerShell tool. To run the tool, open it from the PowerShell command line or right click on it and select "Run with PowerShell".
#		 For more information on the requirements for setting up and using this tool please visit http://www.myteamslab.com.
#
# Copyright: Copyright (c) 2024, James Cussen (www.myteamslab.com) All rights reserved.
# Licence: 	Redistribution and use of script, source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#				1) Redistributions of script code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				2) Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				3) Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
#				4) This license does not include any resale or commercial use of this software.
#				5) Any portion of this software may not be reproduced, duplicated, copied, sold, resold, or otherwise exploited for any commercial purpose without express written consent of James Cussen.
#			THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; LOSS OF GOODWILL OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
# Prerequisites: There are a number of prerequisistes required to get bulk sign-in to work. Please check out the blog post and www.myteamslab.com for more details.
# Known Issues: None.
# Release Notes:
# 1.00 - Initial Release
#
#########################################################################


param ()

$theVersion = $PSVersionTable.PSVersion
$MajorVersion = $theVersion.Major
$MinorVersion = $theVersion.Minor

$OS = [environment]::OSVersion
if($OS -match "Windows")
{
	Write-Host "This is a Windows Machine. CHECK PASSED!" -foreground "green"
}
else
{
	Write-Host "This is not a Windows machine. You're in untested territory, good luck. If it doesn't work, try Windows." -foreground "Yellow"	
}

$DotNetCoreCommands = $false
Write-Host ""
Write-Host "--------------------------------------------------------------"
Write-Host "Powershell Version Check..." -foreground "yellow"
Write-Host "Powershell Version ${MajorVersion}.${MinorVersion}" -foreground "yellow"
if($MajorVersion -eq  "1")
{
	Write-Host "This machine only has Version 1 Powershell installed.  This version of Powershell is not supported." -foreground "red"
	exit
}
elseif($MajorVersion -eq  "2")
{
	Write-Host "This machine has Version 2 Powershell installed. This version of Powershell is not supported." -foreground "red"
	exit
}
elseif($MajorVersion -eq  "3")
{
	Write-Host "This machine has version 3 Powershell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "4")
{
	Write-Host "This machine has version 4 Powershell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "5")
{
	Write-Host "This machine has version 5 Powershell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "6")
{
	Write-Host "ERROR: This machine has version 6 Powershell installed. It's recommended that you upgrade to a minimum of Version 7" -foreground "red"
	exit
}
elseif($MajorVersion -eq  "7")
{
	Write-Host "This machine has version 7 Powershell installed. CHECK PASSED!" -foreground "green"
	$DotNetCoreCommands = $true
}
else
{
	Write-Host "This machine has version ${MajorVersion}.${MinorVersion} of Powershell installed. This tool in GUI mode is not supported with this version. Try command line mode instead." -foreground "red"
	exit
}
Write-Host "--------------------------------------------------------------"
Write-Host ""

#CONNECT TO MICROSOFT TEAMS
try
{
	Connect-MicrosoftTeams
}
catch
{
	Write-Host "ERROR: Sign in failed." -foreground red
	Write-Host "ERROR: " $_ -foreground red
	exit
}

$Tenant = Get-CsTenant
$theTenantID = $Tenant.TenantId

Write-Host "Your Teams Tenant ID is ${theTenantID}. This means your provisioning URLS will be:" -foreground green
Write-Host
Write-Host "------------------------------- For Most Phones -----------------------------------------------" -foreground green
Write-Host "EMEA: http://emea.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}" -foreground green
Write-Host 
Write-Host "Americas: http://noam.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}" -foreground green
Write-Host 
Write-Host "APAC: http://apac.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}" -foreground green
Write-Host "-----------------------------------------------------------------------------------------------" -foreground green
Write-Host
Write-Host "------------------------------- For Cisco Phones ----------------------------------------------" -foreground yellow
Write-Host "EMEA: http://emea.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}/`$PSN.xml" -foreground yellow
Write-Host 
Write-Host "Americas: http://noam.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}/`$PSN.xml" -foreground yellow
Write-Host 
Write-Host "APAC: http://apac.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}/`$PSN.xml" -foreground yellow
Write-Host "-----------------------------------------------------------------------------------------------" -foreground yellow
Write-Host
Write-Host "------------------------------- For AudioCodes ATAs -------------------------------------------" -foreground green
Write-Host "EMEA: http://emea.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}/mac.ini" -foreground green
Write-Host 
Write-Host "Americas: http://noam.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}/mac.ini" -foreground green
Write-Host 
Write-Host "APAC: http://apac.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}/mac.ini" -foreground green
Write-Host "-----------------------------------------------------------------------------------------------" -foreground green
Write-Host
Write-Host "------------------------------- For Cisco ATAs ------------------------------------------------" -foreground yellow
Write-Host "EMEA: http://emea.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}/`$mac.cfg" -foreground yellow
Write-Host 
Write-Host "Americas: http://noam.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}/`$mac.cfg" -foreground yellow
Write-Host 
Write-Host "APAC: http://apac.ipp.sdg.teams.microsoft.com/tenantid/${theTenantID}/`$mac.cfg" -foreground yellow
Write-Host "-----------------------------------------------------------------------------------------------" -foreground yellow
Write-Host

# Set up the form  ============================================================

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Teams SIP Gateway Bulk Sign in Tool 1.00"
$objForm.Size = New-Object System.Drawing.Size(610,380) 
$objForm.MinimumSize = New-Object System.Drawing.Size(500,320) 
$objForm.MaximizeBox = $false
$objForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$objForm.StartPosition = "CenterScreen"
#$objForm.Resizible
#MyTeamsLab Icon
[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$objForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$objForm.KeyPreview = $True
$objForm.TabStop = $false


$MyLinkLabel = New-Object System.Windows.Forms.LinkLabel
$MyLinkLabel.Location = New-Object System.Drawing.Size(450,10)
$MyLinkLabel.Size = New-Object System.Drawing.Size(135,15)
$MyLinkLabel.DisabledLinkColor = [System.Drawing.Color]::Red
$MyLinkLabel.VisitedLinkColor = [System.Drawing.Color]::Blue
$MyLinkLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$MyLinkLabel.LinkColor = [System.Drawing.Color]::Navy
$MyLinkLabel.TabStop = $False
$MyLinkLabel.Text = "  www.myteamslab.com"
$MyLinkLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$MyLinkLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomRight    #TopRight
$MyLinkLabel.add_click(
{
	 [system.Diagnostics.Process]::start("http://www.myteamslab.com")
})
$objForm.Controls.Add($MyLinkLabel)


$CSVLabel = New-Object System.Windows.Forms.Label
$CSVLabel.Location = New-Object System.Drawing.Size(30,40) 
$CSVLabel.Size = New-Object System.Drawing.Size(60,20) 
$CSVLabel.Text = "CSV File:"
$CSVLabel.TabStop = $false
$CSVLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($CSVLabel)


#IPAddressTextBox ============================================================
$CSVTextBox = new-object System.Windows.Forms.textbox
$CSVTextBox.location = new-object system.drawing.size(100,40)
$CSVTextBox.size = new-object system.drawing.size(330,23)
$CSVTextBox.text = ""
$CSVTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$CSVTextBox.tabIndex = 1
$objform.controls.add($CSVTextBox)

$CSVTextBox.add_KeyUp({
	if ($_.KeyCode -eq "Enter") 
	{
		$PathResult = Test-Path -Path $CSVTextBox.text
		
		if($PathResult -eq $true)
		{
			ProcessCSV -filePath $CSVTextBox.text
		}
		else
		{
			Write-Host "ERROR: The file path is not correct. Check the path." -foreground red
		}
	}
})

# BrowseButton ============================================================
$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Location = New-Object System.Drawing.Size(440,39)
$BrowseButton.Size = New-Object System.Drawing.Size(80,20)
$BrowseButton.Text = "Browse..."
$BrowseButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$BrowseButton.Add_Click(
{

	#File Dialog
	[string] $pathVar = $pathbox.Text
	$Filter="All Files (*.*)|*.*"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objDialog = New-Object System.Windows.Forms.OpenFileDialog
	$objDialog.Filter = $Filter
	$objDialog.Title = "Select File Name"
	$objDialog.CheckFileExists = $false
	$Show = $objDialog.ShowDialog()
	if ($Show -eq "OK")
	{
		$filename = $objDialog.FileName
		$CSVTextBox.text = $filename
		$PathResult = Test-Path -Path $CSVTextBox.text
		if($PathResult -eq $true)
		{
			ProcessCSV -filePath $filename
		}
		else
		{
			Write-Host "ERROR: The file path is not correct. Check the path." -foreground red
		}
	}
	
})
$objForm.Controls.Add($BrowseButton)


$RegionLabel = New-Object System.Windows.Forms.Label
$RegionLabel.Location = New-Object System.Drawing.Size(30,70) 
$RegionLabel.Size = New-Object System.Drawing.Size(60,20) 
$RegionLabel.Text = "Region:"
$RegionLabel.TabStop = $false
$RegionLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($RegionLabel)

$RegionDropDownBox = New-Object System.Windows.Forms.ComboBox 
$RegionDropDownBox.Location = New-Object System.Drawing.Size(100,70) 
$RegionDropDownBox.Size = New-Object System.Drawing.Size(170,15) 
$RegionDropDownBox.DropDownHeight = 100 
$RegionDropDownBox.tabIndex = 2
$RegionDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$RegionDropDownBox.DropDownStyle = "DropDownList"

[void] $RegionDropDownBox.Items.Add("APAC")
[void] $RegionDropDownBox.Items.Add("NOAM")
[void] $RegionDropDownBox.Items.Add("EMEA")

$RegionDropDownBox.SelectedIndex = 0

$objForm.Controls.Add($RegionDropDownBox)





$Script:BulkSigninLoop = $true
$Script:BatchID = ""
# BulkSigninButton ============================================================
$BulkSigninButton = New-Object System.Windows.Forms.Button
$BulkSigninButton.Location = New-Object System.Drawing.Size(220,105)
$BulkSigninButton.Size = New-Object System.Drawing.Size(130,23)
$BulkSigninButton.Text = "Bulk Sign In"
$BulkSigninButton.Enabled = $false
$BulkSigninButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$BulkSigninButton.Add_Click(
{
	$PathResult = Test-Path -Path $CSVTextBox.text
	if($PathResult -eq $true)
	{
		$BulkSigninButton.Enabled = $false
		$CancelButton.Enabled = $true
		$SigningInLabel.Visible = $true
		[string] $selectedRegion = $RegionDropDownBox.SelectedItem
		[string] $selectedFilename = $CSVTextBox.text
		Write-Host "`$newBatchResponse = New-CsSdgBulkSignInRequest -DeviceDetailsFilePath $selectedFilename -Region $selectedRegion" -foreground yellow
		$newBatchResponse = New-CsSdgBulkSignInRequest -DeviceDetailsFilePath $selectedFilename -Region $selectedRegion
		$Script:BatchID = $newBatchResponse.BatchId
		Write-Host "BatchID: $batchID" -foreground yellow
		Write-Host	
			
		while($Script:BulkSigninLoop)
		{
			[System.Windows.Forms.Application]::DoEvents()
			Try
			{ 
				$getBatchStatusResponse = Get-CsSdgBulkSignInRequestStatus -Batchid $Script:BatchID -ErrorVariable errvar
				$getBatchStatusResponse | ft
				$getBatchStatusResponse.BatchItem
						
				$resultItems = $getBatchStatusResponse.BatchItem
									
				$loopCount = 0
				foreach($item in $lv.items)
				{
					$user = $item.Text
					
					foreach($resultItem in $resultItems)
					{
						$resultUsername = $resultItem.UserName
						if($user -eq $resultUsername)
						{
							$theStatus = $resultItem.Status
							$theError = $resultItem.Error
							
							#Write-Host "User: $user , Status: $theStatus , Error: $theError" -foreground yellow
							if($theStatus -eq "failed")
							{
								$lv.items[$loopCount].ForeColor = "Red"
							}
							elseif($theStatus -eq "success")
							{
								$lv.items[$loopCount].ForeColor = "Green"
							}
							$lv.items[$loopCount].SubItems[2].Text = $theStatus
							$lv.items[$loopCount].SubItems[3].Text = $theError

						}
					}
					$loopCount++
					
				}
				Write-Host "Batch Status: " $getBatchStatusResponse.BatchStatus -foreground green
				if($getBatchStatusResponse.BatchStatus -eq "Completed")
				{
					$Script:BulkSigninLoop = $false
					Write-Host "INFO: Finished Running Batch!" -foreground yellow
				}
				
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Seconds 1
				write-host
				[System.Windows.Forms.Application]::DoEvents()
			}
			Catch
			{
				Write-Host "ERROR: " $_.Exception.message -foreground red
				
				if($Error[1] -match "Access Denied")
				{
					Write-Host "INFO: This error is usually caused because your account doesn't have suitable administrator access. You need at least Global Administrator, Privileged Authentication Administrator or the Authentication Administrator role access to do Bulk Sign-in." -foreground yellow
				}
				
				$Script:BulkSigninLoop = $false
			}
		}
		
		$Script:BulkSigninLoop = $true
		$CancelButton.Enabled = $false
		$BulkSigninButton.Enabled = $true
		$SigningInLabel.Visible = $false
	}
	else
	{
		Write-Host "ERROR: The file path is not correct. Check the path." -foreground red
	}
})
$objForm.Controls.Add($BulkSigninButton)


$SigningInLabel = New-Object System.Windows.Forms.Label
$SigningInLabel.Location = New-Object System.Drawing.Size(360,109) 
$SigningInLabel.Size = New-Object System.Drawing.Size(120,20) 
$SigningInLabel.Text = "Signing in..."
$SigningInLabel.ForeColor = "Green"
$SigningInLabel.TabStop = $false
$SigningInLabel.Visible = $false
$SigningInLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($SigningInLabel)


$lv = New-Object windows.forms.ListView
$lv.View = [System.Windows.Forms.View]"Details"
$lv.Size = New-Object System.Drawing.Size(560,160)
$lv.Location = New-Object System.Drawing.Size(20,140)
$lv.FullRowSelect = $true
$lv.GridLines = $true
$lv.HideSelection = $false
#$lv.Sorting = [System.Windows.Forms.SortOrder]"Ascending"
$lv.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
[void]$lv.Columns.Add("User", 155)
[void]$lv.Columns.Add("HardwareID", 100)
[void]$lv.Columns.Add("Status", 80)
[void]$lv.Columns.Add("Result", 220)
$objForm.Controls.Add($lv)

$lv.add_SelectedIndexChanged({
	$lv.SelectedItems.Clear()
})


# CancelButton ============================================================
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(50,310)
$CancelButton.Size = New-Object System.Drawing.Size(120,20)
$CancelButton.Text = "Cancel"
$CancelButton.Enabled = $false
$CancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$CancelButton.Add_Click(
{
	Write-Host "INFO: Cancelling checking of batch. Note: This hasn't stopped the batch processing in the background" -foreground yellow
	Write-Host
	Write-Host "You can continue to check this manually with the following commands:" -foreground yellow
	Write-Host "`$getBatchStatusResponse = Get-CsSdgBulkSignInRequestStatus -Batchid $Script:BatchID" -foreground green
	Write-Host "`$getBatchStatusResponse | ft" -foreground green
	Write-Host "`$getBatchStatusResponse.BatchItem" -foreground green
	Write-Host
	$Script:BulkSigninLoop = $false
	
})
$objForm.Controls.Add($CancelButton)


# ExportButton ============================================================
$ExportButton = New-Object System.Windows.Forms.Button
$ExportButton.Location = New-Object System.Drawing.Size(430,310)
$ExportButton.Size = New-Object System.Drawing.Size(120,20)
$ExportButton.Text = "Export Results"
$ExportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$ExportButton.Add_Click(
{
	ExportDataToCSV	
})
$objForm.Controls.Add($ExportButton)


function ProcessCSV([String] $filepath)
{
	$CSVData = Import-Csv -Path $filepath
	if($CSVData | Get-member -MemberType 'NoteProperty'| where {$_.Name -match "Username"})
	{
		Write-Host "INFO: Username Header Exists"  -foreground yellow
		if($CSVData | Get-member -MemberType 'NoteProperty'| where {$_.Name -match "HardwareId"})
		{
			Write-Host "INFO: HardwareId Header Exists" -foreground yellow
			
			$lv.Items.Clear()
			$importCount = 0
			$numberOfItemsInCSV = $CSVData.count
			Write-Host "INFO: Number of items: $numberOfItemsInCSV" -foreground yellow
			if($numberOfItemsInCSV -gt 100)
			{
				Write-Host "ERROR: The CSV contains more that 100 accounts. The Bulk Sign-in process only support 100 accounts per CSV file. Please edit the CSV file and try again." -foreground red
			}
			else
			{
				foreach($item in $CSVData)
				{
					Write-Host "INFO: Importing $importCount" -foreground yellow
					if($importCount -lt 100)
					{
						if($item.Username -match '^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$')
						{
							if($item.HardwareId -match '^([0-9A-Fa-f]{2}[-]){5}([0-9A-Fa-f]{2})$')
							{
								$lvItem = new-object System.Windows.Forms.ListViewItem($item.Username)
								[void]$lvItem.SubItems.Add($item.HardwareId)
								[void]$lvItem.SubItems.Add("")
								[void]$lvItem.SubItems.Add("")
								$lvItem.ForeColor = "Black"
								[void]$lv.Items.Add($lvItem)
							}
							else
							{
								$lvItem = new-object System.Windows.Forms.ListViewItem($item.Username)
								[void]$lvItem.SubItems.Add($item.HardwareId)
								[void]$lvItem.SubItems.Add("")
								[void]$lvItem.SubItems.Add("MAC Address format is incorrect")
								$lvItem.ForeColor = "Red"
								[void]$lv.Items.Add($lvItem)
								$errorname = $item.Username
								Write-Host "ERROR: $errorname MAC Address format is incorrect, it should be in 11-22-33-44-55-66 format" -foreground red
							}
						}
						else
						{
							$errorname = $item.Username
							Write-Host "ERROR: $errorname email address format is wrong" -foreground red
						}
						$BulkSigninButton.Enabled = $true
					}
					else
					{
						Write-Host "ERROR: Ignoring entries greater than 100" -foreground red
					}
					$importCount++
				}
			}
			$lv.Refresh()
		}
		else
		{
			Write-Host "ERROR: The HardwareId header does not exist in the imported CSV. Please fix the format of the CSV and try again." -foreground red
		}
	}
	else
	{
		Write-Host "ERROR: The Username header does not exist in the imported CSV. Please fix the format of the CSV and try again." -foreground red
	}
	
}

function ExportDataToCSV
{
	$filename = ""
	
	Write-Host "Exporting..." -foreground "yellow"
	[string] $pathVar = "C:\"
	$Filter="All Files (*.*)|*.*"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objDialog = New-Object System.Windows.Forms.SaveFileDialog
	$objDialog.FileName = "TeamsSIPGatewayBulkSigninResults-" + (Get-Date -Format 'yyyy-mm-dd-HH-mm-ss') + ".csv"
	$objDialog.Filter = $Filter
	$objDialog.Title = "Export File Name"
	$objDialog.CheckFileExists = $false
	$Show = $objDialog.ShowDialog()
	if ($Show -eq "OK")
	{
		[string]$content = ""
		[string] $filename = $objDialog.FileName
	}
	
	$csv = ""
	if($filename -ne "")
	{	
		$csv = "`"Username`",`"HardwareID`",`"Status`",`"Result`"`r`n"  
				
		foreach($item in $lv.items)
		{
			$UserName = $item.Text
			$HardwareID = $item.SubItems[1].Text
			$Status = $item.SubItems[2].Text
			$Error = $item.SubItems[3].Text
			Write-Host "`"${UserName}`", `"${HardwareID}`", `"${Status}`", `"${Error}`"`r`n"
			$csv += "`"${UserName}`", `"${HardwareID}`", `"${Status}`", `"${Error}`"`r`n"
		}
		
		#Excel seems to only like UTF-8 for CSV files...
		$csv | out-file -Encoding UTF8 -FilePath $filename -Force
		Write-Host "Completed Export." -foreground "yellow"
	}
}

# Activate the form ============================================================
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()



# SIG # Begin signature block
# MIIm9AYJKoZIhvcNAQcCoIIm5TCCJuECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU04farJU/RWsz/gT//eFWTLrW
# ixqggiCcMIIFjTCCBHWgAwIBAgIQDpsYjvnQLefv21DiCEAYWjANBgkqhkiG9w0B
# AQwFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMjIwODAxMDAwMDAwWhcNMzExMTA5MjM1OTU5WjBiMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQw
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC/5pBzaN675F1KPDAiMGkz
# 7MKnJS7JIT3yithZwuEppz1Yq3aaza57G4QNxDAf8xukOBbrVsaXbR2rsnnyyhHS
# 5F/WBTxSD1Ifxp4VpX6+n6lXFllVcq9ok3DCsrp1mWpzMpTREEQQLt+C8weE5nQ7
# bXHiLQwb7iDVySAdYyktzuxeTsiT+CFhmzTrBcZe7FsavOvJz82sNEBfsXpm7nfI
# SKhmV1efVFiODCu3T6cw2Vbuyntd463JT17lNecxy9qTXtyOj4DatpGYQJB5w3jH
# trHEtWoYOAMQjdjUN6QuBX2I9YI+EJFwq1WCQTLX2wRzKm6RAXwhTNS8rhsDdV14
# Ztk6MUSaM0C/CNdaSaTC5qmgZ92kJ7yhTzm1EVgX9yRcRo9k98FpiHaYdj1ZXUJ2
# h4mXaXpI8OCiEhtmmnTK3kse5w5jrubU75KSOp493ADkRSWJtppEGSt+wJS00mFt
# 6zPZxd9LBADMfRyVw4/3IbKyEbe7f/LVjHAsQWCqsWMYRJUadmJ+9oCw++hkpjPR
# iQfhvbfmQ6QYuKZ3AeEPlAwhHbJUKSWJbOUOUlFHdL4mrLZBdd56rF+NP8m800ER
# ElvlEFDrMcXKchYiCd98THU/Y+whX8QgUWtvsauGi0/C1kVfnSD8oR7FwI+isX4K
# Jpn15GkvmB0t9dmpsh3lGwIDAQABo4IBOjCCATYwDwYDVR0TAQH/BAUwAwEB/zAd
# BgNVHQ4EFgQU7NfjgtJxXWRM3y5nP+e6mK4cD08wHwYDVR0jBBgwFoAUReuir/SS
# y4IxLVGLp6chnfNtyA8wDgYDVR0PAQH/BAQDAgGGMHkGCCsGAQUFBwEBBG0wazAk
# BggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAC
# hjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURS
# b290Q0EuY3J0MEUGA1UdHwQ+MDwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwEQYDVR0gBAowCDAGBgRV
# HSAAMA0GCSqGSIb3DQEBDAUAA4IBAQBwoL9DXFXnOF+go3QbPbYW1/e/Vwe9mqyh
# hyzshV6pGrsi+IcaaVQi7aSId229GhT0E0p6Ly23OO/0/4C5+KH38nLeJLxSA8hO
# 0Cre+i1Wz/n096wwepqLsl7Uz9FDRJtDIeuWcqFItJnLnU+nBgMTdydE1Od/6Fmo
# 8L8vC6bp8jQ87PcDx4eo0kxAGTVGamlUsLihVo7spNU96LHc/RzY9HdaXFSMb++h
# UD38dglohJ9vytsgjTVgHAIDyyCwrFigDkBjxZgiwbJZ9VVrzyerbHbObyMt9H5x
# aiNrIv8SuFQtJ37YOtnwtoeW/VvRXKwYw02fc7cBqZ9Xql4o4rmUMIIGrjCCBJag
# AwIBAgIQBzY3tyRUfNhHrP0oZipeWzANBgkqhkiG9w0BAQsFADBiMQswCQYDVQQG
# EwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNl
# cnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwHhcNMjIw
# MzIzMDAwMDAwWhcNMzcwMzIyMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UE
# ChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQg
# UlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBMIICIjANBgkqhkiG9w0BAQEF
# AAOCAg8AMIICCgKCAgEAxoY1BkmzwT1ySVFVxyUDxPKRN6mXUaHW0oPRnkyibaCw
# zIP5WvYRoUQVQl+kiPNo+n3znIkLf50fng8zH1ATCyZzlm34V6gCff1DtITaEfFz
# sbPuK4CEiiIY3+vaPcQXf6sZKz5C3GeO6lE98NZW1OcoLevTsbV15x8GZY2UKdPZ
# 7Gnf2ZCHRgB720RBidx8ald68Dd5n12sy+iEZLRS8nZH92GDGd1ftFQLIWhuNyG7
# QKxfst5Kfc71ORJn7w6lY2zkpsUdzTYNXNXmG6jBZHRAp8ByxbpOH7G1WE15/teP
# c5OsLDnipUjW8LAxE6lXKZYnLvWHpo9OdhVVJnCYJn+gGkcgQ+NDY4B7dW4nJZCY
# OjgRs/b2nuY7W+yB3iIU2YIqx5K/oN7jPqJz+ucfWmyU8lKVEStYdEAoq3NDzt9K
# oRxrOMUp88qqlnNCaJ+2RrOdOqPVA+C/8KI8ykLcGEh/FDTP0kyr75s9/g64ZCr6
# dSgkQe1CvwWcZklSUPRR8zZJTYsg0ixXNXkrqPNFYLwjjVj33GHek/45wPmyMKVM
# 1+mYSlg+0wOI/rOP015LdhJRk8mMDDtbiiKowSYI+RQQEgN9XyO7ZONj4KbhPvbC
# dLI/Hgl27KtdRnXiYKNYCQEoAA6EVO7O6V3IXjASvUaetdN2udIOa5kM0jO0zbEC
# AwEAAaOCAV0wggFZMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFLoW2W1N
# hS9zKXaaL3WMaiCPnshvMB8GA1UdIwQYMBaAFOzX44LScV1kTN8uZz/nupiuHA9P
# MA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB3BggrBgEFBQcB
# AQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggr
# BgEFBQcwAoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1
# c3RlZFJvb3RHNC5jcnQwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcmwwIAYDVR0gBBkwFzAI
# BgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQB9WY7Ak7Zv
# mKlEIgF+ZtbYIULhsBguEE0TzzBTzr8Y+8dQXeJLKftwig2qKWn8acHPHQfpPmDI
# 2AvlXFvXbYf6hCAlNDFnzbYSlm/EUExiHQwIgqgWvalWzxVzjQEiJc6VaT9Hd/ty
# dBTX/6tPiix6q4XNQ1/tYLaqT5Fmniye4Iqs5f2MvGQmh2ySvZ180HAKfO+ovHVP
# ulr3qRCyXen/KFSJ8NWKcXZl2szwcqMj+sAngkSumScbqyQeJsG33irr9p6xeZmB
# o1aGqwpFyd/EjaDnmPv7pp1yr8THwcFqcdnGE4AJxLafzYeHJLtPo0m5d2aR8XKc
# 6UsCUqc3fpNTrDsdCEkPlM05et3/JWOZJyw9P2un8WbDQc1PtkCbISFA0LcTJM3c
# HXg65J6t5TRxktcma+Q4c6umAU+9Pzt4rUyt+8SVe+0KXzM5h0F4ejjpnOHdI/0d
# KNPH+ejxmF/7K9h+8kaddSweJywm228Vex4Ziza4k9Tm8heZWcpw8De/mADfIBZP
# J/tgZxahZrrdVcA6KYawmKAr7ZVBtzrVFZgxtGIJDwq9gdkT/r+k0fNX2bwE+oLe
# Mt8EifAAzV3C+dAjfwAL5HYCJtnwZXZCpimHCUcr5n8apIUP/JiW9lVUKx+A+sDy
# Divl1vupL0QVSucTDh3bNzgaoSv27dZ8/DCCBrAwggSYoAMCAQICEAitQLJg0pxM
# n17Nqb2TrtkwDQYJKoZIhvcNAQEMBQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UE
# AxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIxMDQyOTAwMDAwMFoXDTM2
# MDQyODIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IENvZGUgU2lnbmluZyBS
# U0E0MDk2IFNIQTM4NCAyMDIxIENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBANW0L0LQKK14t13VOVkbsYhC9TOM6z2Bl3DFu8SFJjCfpI5o2Fz16zQk
# B+FLT9N4Q/QX1x7a+dLVZxpSTw6hV/yImcGRzIEDPk1wJGSzjeIIfTR9TIBXEmtD
# mpnyxTsf8u/LR1oTpkyzASAl8xDTi7L7CPCK4J0JwGWn+piASTWHPVEZ6JAheEUu
# oZ8s4RjCGszF7pNJcEIyj/vG6hzzZWiRok1MghFIUmjeEL0UV13oGBNlxX+yT4Us
# SKRWhDXW+S6cqgAV0Tf+GgaUwnzI6hsy5srC9KejAw50pa85tqtgEuPo1rn3MeHc
# reQYoNjBI0dHs6EPbqOrbZgGgxu3amct0r1EGpIQgY+wOwnXx5syWsL/amBUi0nB
# k+3htFzgb+sm+YzVsvk4EObqzpH1vtP7b5NhNFy8k0UogzYqZihfsHPOiyYlBrKD
# 1Fz2FRlM7WLgXjPy6OjsCqewAyuRsjZ5vvetCB51pmXMu+NIUPN3kRr+21CiRshh
# WJj1fAIWPIMorTmG7NS3DVPQ+EfmdTCN7DCTdhSmW0tddGFNPxKRdt6/WMtyEClB
# 8NXFbSZ2aBFBE1ia3CYrAfSJTVnbeM+BSj5AR1/JgVBzhRAjIVlgimRUwcwhGug4
# GXxmHM14OEUwmU//Y09Mu6oNCFNBfFg9R7P6tuyMMgkCzGw8DFYRAgMBAAGjggFZ
# MIIBVTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBRoN+Drtjv4XxGG+/5h
# ewiIZfROQjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMBwGA1UdIAQVMBMwBwYFZ4EMAQMw
# CAYGZ4EMAQQBMA0GCSqGSIb3DQEBDAUAA4ICAQA6I0Q9jQh27o+8OpnTVuACGqX4
# SDTzLLbmdGb3lHKxAMqvbDAnExKekESfS/2eo3wm1Te8Ol1IbZXVP0n0J7sWgUVQ
# /Zy9toXgdn43ccsi91qqkM/1k2rj6yDR1VB5iJqKisG2vaFIGH7c2IAaERkYzWGZ
# gVb2yeN258TkG19D+D6U/3Y5PZ7Umc9K3SjrXyahlVhI1Rr+1yc//ZDRdobdHLBg
# XPMNqO7giaG9OeE4Ttpuuzad++UhU1rDyulq8aI+20O4M8hPOBSSmfXdzlRt2V0C
# FB9AM3wD4pWywiF1c1LLRtjENByipUuNzW92NyyFPxrOJukYvpAHsEN/lYgggnDw
# zMrv/Sk1XB+JOFX3N4qLCaHLC+kxGv8uGVw5ceG+nKcKBtYmZ7eS5k5f3nqsSc8u
# pHSSrds8pJyGH+PBVhsrI/+PteqIe3Br5qC6/To/RabE6BaRUotBwEiES5ZNq0RA
# 443wFSjO7fEYVgcqLxDEDAhkPDOPriiMPMuPiAsNvzv0zh57ju+168u38HcT5uco
# P6wSrqUvImxB+YJcFWbMbA7KxYbD9iYzDAdLoNMHAmpqQDBISzSoUSC7rRuFCOJZ
# DW3KBVAr6kocnqX9oKcfBnTn8tZSkP2vhUgh+Vc7tJwD7YZF9LRhbr9o4iZghurI
# r6n+lB3nYxs6hlZ4TjCCBsIwggSqoAMCAQICEAVEr/OUnQg5pr/bP1/lYRYwDQYJ
# KoZIhvcNAQELBQAwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQTAeFw0yMzA3MTQwMDAwMDBaFw0zNDEwMTMyMzU5NTla
# MEgxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjEgMB4GA1UE
# AxMXRGlnaUNlcnQgVGltZXN0YW1wIDIwMjMwggIiMA0GCSqGSIb3DQEBAQUAA4IC
# DwAwggIKAoICAQCjU0WHHYOOW6w+VLMj4M+f1+XS512hDgncL0ijl3o7Kpxn3GIV
# WMGpkxGnzaqyat0QKYoeYmNp01icNXG/OpfrlFCPHCDqx5o7L5Zm42nnaf5bw9Yr
# IBzBl5S0pVCB8s/LB6YwaMqDQtr8fwkklKSCGtpqutg7yl3eGRiF+0XqDWFsnf5x
# XsQGmjzwxS55DxtmUuPI1j5f2kPThPXQx/ZILV5FdZZ1/t0QoRuDwbjmUpW1R9d4
# KTlr4HhZl+NEK0rVlc7vCBfqgmRN/yPjyobutKQhZHDr1eWg2mOzLukF7qr2JPUd
# vJscsrdf3/Dudn0xmWVHVZ1KJC+sK5e+n+T9e3M+Mu5SNPvUu+vUoCw0m+PebmQZ
# BzcBkQ8ctVHNqkxmg4hoYru8QRt4GW3k2Q/gWEH72LEs4VGvtK0VBhTqYggT02ke
# fGRNnQ/fztFejKqrUBXJs8q818Q7aESjpTtC/XN97t0K/3k0EH6mXApYTAA+hWl1
# x4Nk1nXNjxJ2VqUk+tfEayG66B80mC866msBsPf7Kobse1I4qZgJoXGybHGvPrhv
# ltXhEBP+YUcKjP7wtsfVx95sJPC/QoLKoHE9nJKTBLRpcCcNT7e1NtHJXwikcKPs
# CvERLmTgyyIryvEoEyFJUX4GZtM7vvrrkTjYUQfKlLfiUKHzOtOKg8tAewIDAQAB
# o4IBizCCAYcwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/
# BAwwCgYIKwYBBQUHAwgwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcB
# MB8GA1UdIwQYMBaAFLoW2W1NhS9zKXaaL3WMaiCPnshvMB0GA1UdDgQWBBSltu8T
# 5+/N0GSh1VapZTGj3tXjSTBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRSU0E0MDk2U0hBMjU2VGltZVN0
# YW1waW5nQ0EuY3JsMIGQBggrBgEFBQcBAQSBgzCBgDAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuZGlnaWNlcnQuY29tMFgGCCsGAQUFBzAChkxodHRwOi8vY2FjZXJ0
# cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRSU0E0MDk2U0hBMjU2VGlt
# ZVN0YW1waW5nQ0EuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCBGtbeoKm1mBe8cI1P
# ijxonNgl/8ss5M3qXSKS7IwiAqm4z4Co2efjxe0mgopxLxjdTrbebNfhYJwr7e09
# SI64a7p8Xb3CYTdoSXej65CqEtcnhfOOHpLawkA4n13IoC4leCWdKgV6hCmYtld5
# j9smViuw86e9NwzYmHZPVrlSwradOKmB521BXIxp0bkrxMZ7z5z6eOKTGnaiaXXT
# UOREEr4gDZ6pRND45Ul3CFohxbTPmJUaVLq5vMFpGbrPFvKDNzRusEEm3d5al08z
# jdSNd311RaGlWCZqA0Xe2VC1UIyvVr1MxeFGxSjTredDAHDezJieGYkD6tSRN+9N
# UvPJYCHEVkft2hFLjDLDiOZY4rbbPvlfsELWj+MXkdGqwFXjhr+sJyxB0JozSqg2
# 1Llyln6XeThIX8rC3D0y33XWNmdaifj2p8flTzU8AL2+nCpseQHc2kTmOt44Owde
# OVj0fHMxVaCAEcsUDH6uvP6k63llqmjWIso765qCNVcoFstp8jKastLYOrixRoZr
# uhf9xHdsFWyuq69zOuhJRrfVf8y2OMDY7Bz1tqG4QyzfTkx9HmhwwHcK1ALgXGC7
# KP845VJa1qwXIiNO9OzTF/tQa/8Hdx9xl0RBybhG02wyfFgvZ0dl5Rtztpn5aywG
# Ru9BHvDwX+Db2a2QgESvgBBBijCCBtswggTDoAMCAQICEAK4JLn3OCTJN67E9GA/
# 16owDQYJKoZIhvcNAQELBQAwaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IENvZGUgU2ln
# bmluZyBSU0E0MDk2IFNIQTM4NCAyMDIxIENBMTAeFw0yMzAyMDgwMDAwMDBaFw0y
# NjAyMDkyMzU5NTlaMGAxCzAJBgNVBAYTAkFVMREwDwYDVQQIEwhWaWN0b3JpYTEQ
# MA4GA1UEBxMHTWl0Y2hhbTEVMBMGA1UEChMMSmFtZXMgQ3Vzc2VuMRUwEwYDVQQD
# EwxKYW1lcyBDdXNzZW4wggGiMA0GCSqGSIb3DQEBAQUAA4IBjwAwggGKAoIBgQCp
# 1qIzJ5FCEbE4hHac6gDX5WGYGdqOODOzlFGzSW7uWj1RJoQAak0uelj8ktq0msv0
# IK46cTTwa0ygvhoNc1D1OmcFmnPwuNtE3PB8B9sxoC20g5tBSKtjQM6xRmTlnhXN
# D1mY/rG8QV1sdrcAg+pl1F4lauFuOui64+7JisQoaRqpZFB8d1XOimGsKE6+Mhip
# e2d2zLqGHkB2bOwgdmFtfmY/Kf7vCQa0yObHLiBORu1+aXIQV434olLOsOV7hj4R
# 5VVoArIYW3e5vk6GdMXLw34GAF8RbRSJaEJC5RDSXIRfBwLSLfK9ICpThF2CRZY9
# Jm8KIXKuT79aM9IGznVPpAhtaFXVOkomrNzE6Hpugs/soUGw0QcAw7yHbDsWG/I7
# ohCtvzOLVjUfQcbjqV4jSjjd0kJSvneZX0uMUdlPv5ivnbs4pwwrezg6rDH7wL1+
# HJLr4Q7TUdZPk6ZCE2x6TVf9lt9UlagP5rLJZESeFwAQUQf0/q+5s4yH5yZIV60C
# AwEAAaOCAgYwggICMB8GA1UdIwQYMBaAFGg34Ou2O/hfEYb7/mF7CIhl9E5CMB0G
# A1UdDgQWBBRDBV7JdjdywqUzs5uQAXqnm90YPDAOBgNVHQ8BAf8EBAMCB4AwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwgbUGA1UdHwSBrTCBqjBToFGgT4ZNaHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29kZVNpZ25pbmdSU0E0
# MDk2U0hBMzg0MjAyMUNBMS5jcmwwU6BRoE+GTWh0dHA6Ly9jcmw0LmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRHNENvZGVTaWduaW5nUlNBNDA5NlNIQTM4NDIw
# MjFDQTEuY3JsMD4GA1UdIAQ3MDUwMwYGZ4EMAQQBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCBlAYIKwYBBQUHAQEEgYcwgYQwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBcBggrBgEFBQcwAoZQ
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29k
# ZVNpZ25pbmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5jcnQwDAYDVR0TAQH/BAIwADAN
# BgkqhkiG9w0BAQsFAAOCAgEAzhq+608iSAQxCw2PxRuiUNmd6iJDEN7a86g4dqt7
# DK/cG7RMVolgktwLtr84GEH5zKm35Ib+ioZD1ZpvUi/zsBwlBHjI7JqMPqdlUEOe
# LviFaB1MzH/maxqLm6HD6xTpj8LqYUBwc4TNqrECAvvGYz26XDQfyt8uDb/E8sv/
# p+NYxLu9Rtp3Y/SFgvwdCieS7zO78Hz9lspH34TotOHisqHc/3ONMXsDHEHBocK+
# R+7h1/X3rb/DrbjV40Mw13TSGxvmSgKjoozC2IMYwRGHB57I9mN5TA3OopbTBI13
# KtfDvIxxFTnAfsJ/3MXYtl8bSgsNtZNI7DawswvIV9HBxknXD1fmuaSNRX2ubNyB
# sj5mnSgKxLLKOg5cWeLzkU99IF47XcRYNPAD+PY3SNahB98Zusb9RKonyOkowtcR
# fsbOtEcyqvOSk/5Drqpqu8imCc1NtnOlx1frWSRmIzmfyYRoNXp7uuzU1cGsX77a
# elvtBRd0Chcg/EotCJNuDl/nTAgORYRRXcay/LovPXRc9LzzsNqDkFnjJZ/eSlXh
# tsqdEY9lbRC7dA4dslM8V7dpR+khDQyw5GfDa3outtY4MyKVt+mh7vGt0vFtHHva
# NuPAgKubT/nk+Xprz5OAP87UBeWfdnp4Dme13XNC8VYnNTJhbZOqsQ5pATLkzKuO
# UrMxggXCMIIFvgIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25p
# bmcgUlNBNDA5NiBTSEEzODQgMjAyMSBDQTECEAK4JLn3OCTJN67E9GA/16owCQYF
# Kw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkD
# MQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJ
# KoZIhvcNAQkEMRYEFB1xcWMxUqtYZFhNw5eE7FMdFEhJMA0GCSqGSIb3DQEBAQUA
# BIIBgDfv/GBJLnqUXMfpxsJ5J1Gyiq3N27pjvCZU+WmvROjnP9S+ZZIrDGdu79Et
# M1SUcafPJXKrvj9jBL7Xh6wAnC9NcL12hgbpWoC8zkQtap7+cKbSSMDTqmbpxGQi
# uCUrBPHPBrniCS/QmAXGvjgjnMtYrmeGKysgnWLOLS2qiXocAm4MHv2YXmbDAIIb
# yO09vP2L3thEvKxne54g4Xx3qlXbzOC4aso7NkYQ30brbS2Mgclbi3R1Yip2hX2j
# bRPhm4hr+5oQoR3/H/h2nf4wCHst5Fq1GmlBP79ja92xbRF6u7/fU0h1FJKtwDTL
# i+L6P9VFLF01+Ftz/WHkCKGWcROvQXiW/VIU4rCyZgThcEYjOnO9D318CQTaWAGd
# 0wNTRex8fFsUSrDs4pOvjiEYyPAvsrYWOy15RJJAsOLs8eLQzvhfu3oW5d2nlhMw
# MmCI4nRVeyQtmd2zCPVNk/TxGct05q9nMpLxhUt8MIN6/4WsDjvrB+2zs9qAq3e1
# MsAA5qGCAyAwggMcBgkqhkiG9w0BCQYxggMNMIIDCQIBATB3MGMxCzAJBgNVBAYT
# AlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQg
# VHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0ECEAVEr/OU
# nQg5pr/bP1/lYRYwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNDA2MTYxMjUzMjNaMC8GCSqGSIb3DQEJ
# BDEiBCCIBOPl2jSg2yuefEQ62K31VEdSYaaQgHFQ+vAfTCOUsDANBgkqhkiG9w0B
# AQEFAASCAgBiMLiqzMrHZtsU9cfEM2HasA5wWUEY0Wk8h6Fadt+g3uOPZkIW9QyS
# Y96a+c21Sg81B/Bb1IY2hjMTOgpeIntPaNK0tPpfOcZHvIBwGXGhkzwigF+Y5Mks
# 8YUbW1dpZC4cXiJKP2jWnucorU6W+kNYI2LClkCISf1CGy5XDuVJ7bJBSNakHs7x
# GBF9zGInXIxd1r5uRmxgvG0U4EEj3HjdKhfGsdmw7jDg8G2K6R3EzyQuhZH5Xp48
# /u8lxtMZ8l1HTF+bOLaLZ919vO+0RVO+hMuNvYduWYjCZW/OZmFqJRZvgNNiGaWX
# vnG7kDQ6Csz9Ddjz/7/HjUJxowHh0/2vO+LpOazIhw6pGUeHSuiZ4z6di8ke+/ZK
# yXAaU5XAQYoV6dJZDQo3KCIdXyV0DRvBwnKyfdzEXNt/wsUDCaJcLNwlXXbzD9I+
# j/vYQAjFdPJq+h6XdCZcE2u1XGGMflEpiIEzMAfpG9AkeUnw4/ITl83qjGD0l2Eo
# IyX4kBH1W69G6HmcuDAiThWVN/5cvPgOQcn4d8JbuEjIevFughJae/qqpbaI+fnv
# RkacQE6QIEIOXlIcufoXG3W1xAYyWGQPibxIvwqmBE+Vc80gUmu0qw4HjLJRouvn
# MBT0BZx9uu8pKgyVKb7zq5fJr8tPARYtxPi0GnrWXQMszME57gDmdQ==
# SIG # End signature block
