# Written by Kris Stevenson
# 02/08/18
# IIT Help Installer V1.1
#
# A Powershell Script to automate the installing of the XPanel/Adobe AIR 
# software, copy the IIT Help Page files into their correct directory, and 
# create a public desktop shortcut to be accessed by anyone logging on to the 
# instructors PC. 
# The script needs to be accompanied by the Icon file/Installers in the XPanel
# folder, along with each room's XPanel folder (which should include its own 
# XPanel file) 
# Eg. script on root E:\ and the files in E:\XPanels\*room names*

# V1.0 notes (31st Jan 2018)
# - IMPORTANT! Currently, the switch in the main function is set so only 
# specific host PC's (Instructors stations) will work. Adding extra rooms/PC's
# would require changing this switch and adding the relevant files/folders.
#
# - There is basic error checking in this installer which should cover the
# majority of, if not all possible errors. However, the install can easily be
# checked by running the program from the desktop shortcut to see if it is 
# functioning correctly.

# V1.1 changes (8th Feb 2018)
# - Script now checks to see if the required XPanel/Adobe Air are installed on
# the PC, and installs them if they are not.
#
# - Only the XPanel file needs to be in each room's folder now. 
# Previously the XPanel/Adobe AIR installers, along with the Icon file, needed
# to be in every room folder. Now, these extra files need only be in the main
# XPanel folder.
# This change was made so if files are updated (Installer/icon) then they only
# need to be copied to one location, and not every room folder as was required
# in V1.0
#
# - Moved the script variables for file/directory names up to the top of the 
# program so they can be easily changed if required.
# EG. If the icon file needs to be updated/renamed, etc.
# 
# - General code clean up, encapsulating, and commenting.



# ********** Script variables for file names **********
# If newer versions/updated files with different names are needed, change
# those file names here.
#
$Script:xPanelFile = "DTCC Fusion Rooms v2.c3p"
$Script:xPanelInstaller	= "CrestronXPanel Installer.exe"
$Script:airFile	= "CrestronXPanel Installer.air"
$Script:iconFile = "IIT HELP v3.ico"
#
# *****************************************************


# Find Computer host name
Function getPCHostName($script:pcHostName)
{
	# Attempt to retrieve the PC's host name
	if (!([System.Net.Dns]::GetHostName()))
	{
		write-host "`nError: Unable to retrieve PC Hostname"
		$script:errorThrown = $true
	}
	# If found, copy the value in to $script:pcHostName
	else
	{
		$script:pcHostName = [System.Net.Dns]::GetHostName() 
		write-host "PC HostName: $script:pcHostName"
	}
}

# Check if XPanel/Adobe AIR programs are installed, and install them if they are not
Function checkXPanelInstall()
{
	write-host "`nChecking XPanel/Adobe AIR installs"
	$Script:adobeAirExists = Test-Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Adobe AIR'
	$Script:xPanelExists = Test-Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\CrestronXPanel'
	write-host "Adobe AIR installed: $Script:adobeAirExists"
	write-host "XPanel installed: $Script:xPanelExists"

	# Check if either XPanel or Adobe AIR programs are missing and install them if they are
	if ((-not $Script:adobeAirExists) -or (-not $Script:xPanelExists))
	{
		Write-Host "Installing XPanel/Adobe AIR software"
		
		# Silently install the XPanel software (which includes Adobe AIR)
		# passthru/wait will collect any exit codes/wait for the installer to 
		# finish before moving on
		$Script:installProcess = (Start-Process -FilePath "$Script:scriptPath\XPanels\$Script:xPanelInstaller" -ArgumentList "-silent -eulaAccepted -allowDownload" -PassThru -Wait)

		# If exit code 0 = Install complete
		if ($Script:installProcess.exitcode -eq "0")
		{
			write-host "Software installed"
		}
		# Throw an error for any other exit code and write the code to the window
		else
		{
			$script:errorThrown = $true
			write-host "Installer exit code: " $Script:installProcess.exitcode
			write-host "`nERROR: Could not install XPanel/Adobe AIR Software."
		}
	}
	#If both programs are installed, skip the above code and inform the user
	else
	{
		write-host "XPanel/Adobe AIR software is already installed, moving to next stage"
	}
}
# mirror files into their directory (From V1.0 - currently an unused function )
Function mirrorSourceFiles($script:roomName)
{

	if ($script:errorThrown -eq $false)
	{
		write-host "Mirroring source files"
		$source = "$Script:scriptPath\XPanels\$script:roomName"
		$destination = "C:\Program Files (x86)\Crestron\$script:roomName"
		# $Files = "*.*"  
		# $Switches = "/R:0 /W:0 /E /MIR /TEE /Z /TIMFIX"

		robocopy $source $destination /mir /NJH /NFL /NDL #/NJS
		
		if ($lastexitcode -lt 8)
		{
			write-host "`nCopied files for $script:roomName"
		}
		else
		{
			write-host "`nError: Copy failed with exit code: $lastexitcode `nCould not copy files due to errors"
			$script:errorThrown = $true
		}
	}
}

# Copy files into their directory
Function copySourceFiles($script:roomName)
{
	# Robocopy syntax (For reference)
	# Robocopy [source] [destination] [file to be copied]
	

	# ********** Copy XPanel File **********	
	# If no errors have been thrown previously, attempt to copy files
	if ($script:errorThrown -eq $false)
	{
		write-host "`nCopying source files"
		
		$source = "$Script:scriptPath\XPanels\$script:roomName"
		$destination = "C:\Program Files (x86)\Crestron\$script:roomName"
		
		robocopy $source $destination $Script:xPanelFile /NJH /NFL /NDL /NJS
		
		if ($lastexitcode -lt 8)
		{
			write-host "Copied XPanel file for $script:roomName"
		}
		else
		{
			write-host "`nError: Copy failed with exit code: $lastexitcode `nCould not copy XPanel file due to errors"
			$script:errorThrown = $true
		}
	}
	
	
	
	# ********** Copy XPanel installer **********
	if ($script:errorThrown -eq $false)
	{
		$source = "$Script:scriptPath\XPanels"
		$destination = "C:\Program Files (x86)\Crestron"

		robocopy $source $destination $Script:xPanelInstaller /NJH /NFL /NDL /NJS
		
		if ($lastexitcode -lt 8)
		{
			write-host "Copied XPanel installer"
		}
		else
		{
			write-host "`nError: Copy failed with exit code: $lastexitcode `nCould not copy XPanel installer file due to errors"
			$script:errorThrown = $true
		}
	}
	
	
	
	# ********** Copy Adobe Air Package **********
	if ($script:errorThrown -eq $false)
	{
		$source = "$Script:scriptPath\XPanels"
		$destination = "C:\Program Files (x86)\Crestron"

		robocopy $source $destination $Script:airFile /NJH /NFL /NDL /NJS
		
		if ($lastexitcode -lt 8)
		{
			write-host "Copied Adobe Air package"
		}
		else
		{
			write-host "`nError: Copy failed with exit code: $lastexitcode `nCould not copy Adobe Air package file due to errors"
			$script:errorThrown = $true
		}
	}
	
	
	
	# ********** Copy IIT Help Icon File **********
	if ($script:errorThrown -eq $false)
	{
		$source = "$Script:scriptPath\XPanels"
		$destination = "C:\Program Files (x86)\Crestron"

		robocopy $source $destination $Script:iconFile /NJH /NFL /NDL /NJS
		
		if ($lastexitcode -lt 8)
		{
			write-host "Copied Help Icon file"
		}
		else
		{
			write-host "`nError: Copy failed with exit code: $lastexitcode `nCould not copy Help Icon file due to errors"
			$script:errorThrown = $true
		}
	}
}

# Create shortcut file and assign it the IIT Help icon
Function createShortcutFile($script:roomName)
{
	# If no errors have been thrown previously, attempt to create the shortcut file
	if ($script:errorThrown -eq $false)
	{
		write-host "`nCreating shortcut file."
		$script:targetFile = "Null"
		$script:shortcutFile = "Null"
		$script:iconSource = "Null"
		
		# Check the shortcut's target path 
		# -Throw an error if it is not valid
		# -Set the target path if it is valid
		if (!(Test-Path "C:\Program Files (x86)\Crestron\$script:roomName\$Script:xPanelFile"))
		{
			write-host "`nError: Unable to see/access local PC's $script:roomName target XPanel File"
			$script:errorThrown = $true
		}
		else { $script:targetFile = "C:\Program Files (x86)\Crestron\$script:roomName\$Script:xPanelFile"}
		
		# Check the shortcut's file path 
		# -Throw an error if it is not valid
		# -Set the file path if it is valid
		if (!(Test-Path -IsValid "C:\Users\Public\Desktop\IIT Help.lnk"))
		{
			write-host "`nError: Unable to see/access local PC's public desktop"
			$script:errorThrown = $true
		}
		else { $script:shortcutFile = "C:\Users\Public\Desktop\IIT Help.lnk" }
		

		# Check the shortcut's icon file path 
		# -Throw an error if it is not valid
		# -Set the icon file path if it is valid
		if (!(Test-Path "C:\Program Files (x86)\Crestron\$Script:iconFile"))
		{
			write-host "`nError: Unable to see/access the local PC's icon file"
			$script:errorThrown = $true
		}
		else { $script:iconSource = "C:\Program Files (x86)\Crestron\$Script:iconFile" }

		# Use the above parameters to create the shortcut
		$script:wScriptShell = New-Object -ComObject WScript.Shell
		$script:shortcut = $script:wScriptShell.CreateShortcut($script:shortcutFile)
		$script:shortcut.TargetPath = $script:targetFile
		$script:shortcut.IconLocation = "$script:iconSource, 0"
		if ($script:errorThrown -eq $false)
		{
			$script:Shortcut.Save()
			write-host "Desktop Shortcut (Public) for $script:roomName has been created"
			write-host "`n`nInstall complete."
			write-host "Please close the browser window or wait 20 seconds..."
		}			
	}
}

# Main function
function Main()
{
	# Script Variables 
	$Script:scriptPath = (Get-Location).Path
	$script:pcHostName = "NULL"
	$script:roomName = "NULL"
	$script:roomValid = $false
	$script:errorThrown = $false

	
	write-host "IIT Help Page Installer V1.1"
	Start-Sleep -s 1
	write-host "`nScript is in path... $Script:scriptPath"
	
	
	getPCHostName($script:pcHostName)
	
	# Switch to parse $script:pcHostName, setting $script:roomName and setting
	# the bool value $script:roomValid to true if that room is in the list
	#
	# EG. host name 'S-A200-01' becomes room 'A200' and $script:roomValid is 
	# set to true
	#
	# IMPORTANT! Currently this switch is set so only specific host PC's 
	# (instructors stations) will work. Adding extra rooms/PC's would require 
	# adding the host name to this switch function and adding the corresponding 
	# directory files to the relevant folder.
	Switch -wildcard ($script:pcHostName)
	{	 
		"*Orion*"		{ $script:roomName = "A110"; $script:roomValid = $true }
		# ********** A-Wing **********
		"*s-a110-01*"	{ $script:roomName = "A110"; $script:roomValid = $true }
		"*s-a153-01*"	{ $script:roomName = "A153"; $script:roomValid = $true }
		"*s-a156-01*"	{ $script:roomName = "A156"; $script:roomValid = $true }
		"*s-a200-01*"	{ $script:roomName = "A200"; $script:roomValid = $true }
		"*s-a201-01*"	{ $script:roomName = "A201"; $script:roomValid = $true }
		"*s-a202-01*"	{ $script:roomName = "A202"; $script:roomValid = $true }
		"*s-a204-01*"	{ $script:roomName = "A204"; $script:roomValid = $true }
		"*s-a206-01*"	{ $script:roomName = "A206"; $script:roomValid = $true }
		"*s-a208-01*"	{ $script:roomName = "A208"; $script:roomValid = $true }
		"*s-a210-01*"	{ $script:roomName = "A210"; $script:roomValid = $true }
		"*s-a211-01*"	{ $script:roomName = "A211"; $script:roomValid = $true }
		"*s-a212-01*"	{ $script:roomName = "A212"; $script:roomValid = $true }
		"*s-a213-01*"	{ $script:roomName = "A213"; $script:roomValid = $true }
		"*s-a220-01*"	{ $script:roomName = "A220"; $script:roomValid = $true }
		"*s-a223-01*"	{ $script:roomName = "A223"; $script:roomValid = $true }
		"*s-a224-01*"	{ $script:roomName = "A224"; $script:roomValid = $true }
		"*s-a225-01*"	{ $script:roomName = "A225"; $script:roomValid = $true }
		"*s-a226-01*"	{ $script:roomName = "A226"; $script:roomValid = $true }
		"*s-a230-01*"	{ $script:roomName = "A230"; $script:roomValid = $true }
		"*s-a231-01*"	{ $script:roomName = "A231"; $script:roomValid = $true }
		# ********** B-Wing **********
		"*s-b103-01*"	{ $script:roomName = "B103"; $script:roomValid = $true }
		"*s-b111-01*"	{ $script:roomName = "B111"; $script:roomValid = $true }
		"*s-b120-01*"	{ $script:roomName = "B120"; $script:roomValid = $true }
		"*s-b128-01*"	{ $script:roomName = "B128"; $script:roomValid = $true }
		"*s-b129-01*"	{ $script:roomName = "B129"; $script:roomValid = $true }
		"*s-b130-62*"	{ $script:roomName = "B130"; $script:roomValid = $true }
		"*s-b211-01*"	{ $script:roomName = "B211"; $script:roomValid = $true }
		"*s-b212-01*"	{ $script:roomName = "B212"; $script:roomValid = $true }
		"*s-b228-01*"	{ $script:roomName = "B228"; $script:roomValid = $true }
		"*s-b231-01*"	{ $script:roomName = "B231"; $script:roomValid = $true }
		"*s-b234-01*"	{ $script:roomName = "B234"; $script:roomValid = $true }
		"*s-b237-01*"	{ $script:roomName = "B237"; $script:roomValid = $true }
		# ********** C-Wing **********
		"*s-c128-01*"	{ $script:roomName = "C128"; $script:roomValid = $true }
		"*s-c130-01*"	{ $script:roomName = "C130"; $script:roomValid = $true }
		"*s-c138-01*"	{ $script:roomName = "C138"; $script:roomValid = $true }
		"*s-c213-01*"	{ $script:roomName = "C213"; $script:roomValid = $true }
		"*s-c215-01*"	{ $script:roomName = "C215"; $script:roomValid = $true }
		"*s-c224-01*"	{ $script:roomName = "C224"; $script:roomValid = $true }
		"*s-c225-01*"	{ $script:roomName = "C225"; $script:roomValid = $true }
		"*s-c232-01*"	{ $script:roomName = "C232"; $script:roomValid = $true }
		# ********** D-Wing **********
		#"*placeholder*"        { $script:roomName = "*placeholder*"; $script:roomValid = $true }
		# ********** E-Wing **********
		"*s-e101-01*"	{ $script:roomName = "E101"; $script:roomValid = $true }
		"*s-e103-01*"	{ $script:roomName = "E103"; $script:roomValid = $true }
		"*s-e104-02*"	{ $script:roomName = "E104"; $script:roomValid = $true }
		"*s-e107-01*"	{ $script:roomName = "E107"; $script:roomValid = $true }
		"*s-e109-01*"	{ $script:roomName = "E109"; $script:roomValid = $true }
		"*s-e114-01*"	{ $script:roomName = "E114"; $script:roomValid = $true }
		"*s-e119-01*"	{ $script:roomName = "E119"; $script:roomValid = $true }
		"*s-e123-01*"	{ $script:roomName = "E123"; $script:roomValid = $true }
		"*s-e203-01*"	{ $script:roomName = "E203"; $script:roomValid = $true }
		"*s-e204-01*"	{ $script:roomName = "E204"; $script:roomValid = $true }
		"*s-e205-01*"	{ $script:roomName = "E205"; $script:roomValid = $true }
		"*s-e206-02*"	{ $script:roomName = "E206"; $script:roomValid = $true }
		"*s-e207-01*"	{ $script:roomName = "E207"; $script:roomValid = $true }
		"*s-e211-01*"	{ $script:roomName = "E211"; $script:roomValid = $true }
		"*s-e213-01*"	{ $script:roomName = "E213"; $script:roomValid = $true }
		"*s-e214-01*"	{ $script:roomName = "E214"; $script:roomValid = $true }
		"*s-e216-01*"	{ $script:roomName = "E216"; $script:roomValid = $true }
		"*s-e222-01*"	{ $script:roomName = "E222"; $script:roomValid = $true }
		"*s-e223-01*"	{ $script:roomName = "E223"; $script:roomValid = $true }
		"*s-e224-01*"	{ $script:roomName = "E224"; $script:roomValid = $true }
		# ********** F-Wing **********
		"*s-f101-01*"	{ $script:roomName = "F101"; $script:roomValid = $true }
		"*s-f103-01*"	{ $script:roomName = "F103"; $script:roomValid = $true }
		"*s-f105-01*"	{ $script:roomName = "F105"; $script:roomValid = $true }
		"*s-f108-01*"	{ $script:roomName = "F108"; $script:roomValid = $true }
		"*s-f112-01*"	{ $script:roomName = "F112"; $script:roomValid = $true }
		"*s-f118-01*"	{ $script:roomName = "F118"; $script:roomValid = $true }
		"*s-f121-01*"	{ $script:roomName = "F121"; $script:roomValid = $true }
		"*s-f125-01*"	{ $script:roomName = "F125"; $script:roomValid = $true }
		"*s-f128-01*"	{ $script:roomName = "F128"; $script:roomValid = $true }
		"*s-f131-01*"	{ $script:roomName = "F131"; $script:roomValid = $true }
		"*s-f132-01*"	{ $script:roomName = "F132"; $script:roomValid = $true }
		# ********** G-Wing **********
		"*s-g107a-01*"	{ $script:roomName = "G107A"; $script:roomValid = $true }
		"*s-g107b-01*"	{ $script:roomName = "G107B"; $script:roomValid = $true }
		"*s-g109-01*"	{ $script:roomName = "G109"; $script:roomValid = $true }
		"*s-g110-01*"	{ $script:roomName = "G110"; $script:roomValid = $true }
		# ********** Modular **********
		"*s-class-01*"	{ $script:roomName = "Modular 1"; $script:roomValid = $true }
		"*s-class-02*"	{ $script:roomName = "Modular 2"; $script:roomValid = $true }
		"*s-class-03*"	{ $script:roomName = "Modular 3"; $script:roomValid = $true }
		"*s-class-04*"	{ $script:roomName = "Modular 4"; $script:roomValid = $true }

		default 
		{ 
			$script:roomValid = $false
			write-host "Host name $script:pcHostName is not in the install list for this Powershell file."
		}
	}
	
	# Check if the host name is valid (in the switch statement) and that no 
	# errors have been thrown. If so, attempt to install/copy the relevant files
	if (($script:roomValid -eq $true) -and ($script:errorThrown -eq $false))
	{
		write-host "Room Name $script:roomName is valid, starting install..."
		checkXPanelInstall
		copySourceFiles ($script:roomName)
		createShortcutFile ($script:roomName)
	}
	# If room is not valid, inform the user
	else
	{
		write-host "Room Name $script:roomName is not valid..."
		write-host "IIT Help Page not installed."	
		write-host "Please contact Jason Brown or Kris Stevenson (Stanton Classroom Technology)"
		"for further assistance."
		write-host "The browser window will close in 20 seconds..."
	}
	# Check if any errors have been thrown earlier in the code and inform the user
	If ($script:errorThrown -eq $true)
	{
		write-host "`nCould not complete the install due to errors."
		write-host "Please note the error and contact Jason Brown or Kris Stevenson"
		write-host "(Classroom Technology - Stanton) for further assistance."
		write-host "The browser window will close in 60 seconds..."
		Start-Sleep -s 40
	}
	
	Start-Sleep -s 20
}

# Run the Main function
Main
