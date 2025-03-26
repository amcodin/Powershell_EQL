#requires -Version 3.0
##########################################
#        Client Data Backup Tool         #
#       Written By Stephen Onions        #
#    Edited By Kevin King 04.12.2023     #    I Will be editing and working on this as we go when i have more time outside of DSO hours
#   Edited by Jared Vosters 6/12/23 	#  Prompt user to run GP Update & rearranged persist & global scope for New-PSDrive to fix energyq drive issues
##########################################
$VerbosePreference = 'Continue'
<# 
		This is Designed to run on the client's machine under the client's security token. (i.e. as the client)
        Kevin has updated this to comply with Surface Refresh Tech Install Scope

        It will first prompt user to confirm if gpupdate has been run then,
		By Default this will Backup/Restore the following Directories.
		- C:\User
		- C:\Temp
		- Desktop
        - Documents
		- KML Files
        - Outlook Signatures
        - Chrome Bookmarks
        - Quick Access Shortcuts

		It will also Backup/Restore
        - Network Drives
        - Printer Mappings

		All of this will be saved to a location selected by the client through a save dialog.
#>

Add-Type -AssemblyName presentationframework #add the WPF Assembly. Needed for all the GUI elements.
Add-Type -Language CSharp -TypeDefinition @' 
public class RemoteInfo{
	public string Computername;
	public string ProfileFolder;
	public string ProfileName;
	public bool IsInUse;
	public string SSID;
	public Microsoft.Win32.RegistryKey UserHive;
};
'@ #Custom Class to hold some data across the script.

#region Functions
#all of the functional code (i.e. code that does stuff) is in this region. Only exception is the code that is called on the ContentRendered event of the form.
Function Show-ProgressBar
{
	param([Parameter(Mandatory)][String]$Status,[Parameter(Mandatory)][int]$ProgressValue,[Parameter(Mandatory)][string]$SubStatus,[Parameter(Mandatory)][int]$SubProgressValue)
	$script:Hash_ProgressBar = [Hashtable]::Synchronized(@{}) #create a thread safe hash table to store form items for a progress bar.
	$script:Hash_ProgressBarUpdate = [Hashtable]::Synchronized(@{})
	
	$SBProgressbar = 
	{
		#script block for the creation and running of the progress bar form. this has to run in a separate thread/runspace otherwise it cannot update realtime.
		Add-Type -AssemblyName PresentationFramework #add the presentation framework (new runspace so it won't carry over).
		#the XAML code for the form.
		[XML]$XAML = 
		@'

<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    ResizeMode="NoResize"
    Width="600"
    Height="200"
    Title="Backup Completion"
    WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
      <Label Name="Status" Width="Auto" Height="30" HorizontalAlignment="Stretch" VerticalAlignment="Top" Margin="10,10,10,0"/>
      <ProgressBar Name="pbComplete" Width="Auto" Height="20" HorizontalAlignment="Stretch" VerticalAlignment="Top" Margin="10,45,10,0" Value="0"/>
      <Label Name="SubStatus" Width="Auto" Height="30" HorizontalAlignment="Stretch" VerticalAlignment="Top" Margin="20,65,10,0"/>
      <ProgressBar Name="pbSub" Width="Auto" Height="20" HorizontalAlignment="Stretch" VerticalAlignment="Top" Margin="20,95,10,0" Value="0"/>
    </Grid>
</Window>

'@

		$script:Hash_ProgressBarUpdate.StatusTxt = '' #text for the parent progress bar
		$script:Hash_ProgressBarUpdate.pbCompleteValue = 0 #completion value for the parent progress bar
		$script:Hash_ProgressBarUpdate.SubStatusTxt = '' #text for the child progress bar
		$script:Hash_ProgressBarUpdate.pbSubValue = 0 #competeion value for the child progress bar

		$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $XAML) #convert the xaml
		$script:Hash_ProgressBar.Form = [windows.markup.xamlreader]::Load($reader) #load the form and define the form items to the hash table
		$script:Hash_ProgressBar.Status = $script:Hash_ProgressBar.Form.findname('Status') 
		$script:Hash_ProgressBar.pbComplete = $script:Hash_ProgressBar.Form.findname('pbComplete')
		$script:Hash_ProgressBar.SubStatus = $script:Hash_ProgressBar.Form.findname('SubStatus')
		$script:Hash_ProgressBar.pbSub = $script:Hash_ProgressBar.Form.findname('pbSub')
		
		$script:Hash_ProgressBar.Status.Content = $Status
		$script:Hash_ProgressBar.pbComplete.Value = $ProgressValue
		$script:Hash_ProgressBar.SubStatus.Content = $SubStatus
		$script:Hash_ProgressBar.pbSub.Value = $SubProgressValue
		
		$null = $script:Hash_ProgressBar.form.showdialog() # show the dialog
	}

	$runspace = [runspacefactory]::CreateRunspace() #create a new runspace
	$runspace.ApartmentState = 'STA' #single apartement thread - needed for WPF forms
	$runspace.ThreadOptions = 'ReuseThread' #reuse the thread
	$runspace.Open() #open the runspace
	$runspace.SessionStateProxy.SetVariable('Hash_ProgressBar',$script:Hash_ProgressBar) #add the hashtable to the runspace
	$runspace.SessionStateProxy.SetVariable('Hash_ProgressBarUpdate',$script:Hash_ProgressBarUpdate) #add the hashtable to the runspace
	$PSCmd = [powershell]::Create() #create a new powershell pipeline
	$PSCmd.AddScript($SBProgressbar) #add the script block for it to run
	$PSCmd.Runspace = $runspace #tell it to run in the created runspace
	$PSCmd.BeginInvoke() #begin the invocation
	Start-Sleep -Seconds 1
}
Function Update-ProgressBar
	{
		#Private function to update the progress bar
		param([Parameter(Mandatory)][String]$Status,[Parameter(Mandatory)][int]$ProgressValue,[Parameter(Mandatory)][string]$SubStatus,[Parameter(Mandatory)][int]$SubProgressValue)
		$script:Hash_ProgressBarUpdate.StatusTxt = $Status #assign the supplied status info to the hash table and repeat for all sections
		$script:Hash_ProgressBarUpdate.pbCompleteValue = $ProgressValue
		$script:Hash_ProgressBarUpdate.SubStatusTxt = $SubStatus
		$script:Hash_ProgressBarUpdate.pbSubValue = $SubProgressValue
		#call the form.dispatcher.invoke method to update the form across runspaces. - alternate methods result in a unowned object error.
		$script:Hash_ProgressBar.Form.Dispatcher.Invoke([Action][ScriptBlock]::Create({
					$script:Hash_ProgressBar.Status.Content = $script:Hash_ProgressBarUpdate.Statustxt
					$script:Hash_ProgressBar.pbComplete.Value = $script:Hash_ProgressBarUpdate.pbCompleteValue
					$script:Hash_ProgressBar.SubStatus.Content = $script:Hash_ProgressBarUpdate.SubStatusTxt
					$script:Hash_ProgressBar.pbSub.Value = $script:Hash_ProgressBarUpdate.pbSubValue
		}),'Normal')
	}
Function Get-SSID 
{#prompt support staff to select the profile to run against.
	param([parameter(Mandatory)][Microsoft.Management.Infrastructure.CimSession]$Session)
	$profiles = Get-CimInstance -Query 'SELECT * FROM Win32_UserProfile WHERE Special=False' -CimSession $session  #get the profiles from the remote machine 
	$ProfileDataArray = New-Object -TypeName System.Collections.ArrayList #create an empty arraylist.
	$i=0
	Foreach($account in $profiles)
	{ #for each of the profiles.
		$i++
		Try
		{
			$accountSSID = New-Object -TypeName System.Security.Principal.SecurityIdentifier -ArgumentList ($account.SID) #get the SSID and convert it into a domain\username
			$accountName = $accountSSID.Translate([Security.Principal.NTAccount]).Value
			$ProfileName = (($AccountName -split '\\')[1]) #username
			$Domain = (($accountName -split '\\')[0]) #domain
		}
		Catch
		{
			$accountName = $account.SID #if unable to convert - use the SSID.
			$ProfileName = (($AccountName -split '\\')[1]) #username
			$Domain = 'Unknown'
		}
		Switch($account.Status)
		{#check the account status
			1 {$AccountType = 'Temporary'} 
			2 {$AccountType = 'Roaming'}
			4 {$AccountType = 'Mandatory'}
			8 {$AccountType = 'Corrupted'}
			Default {$AccountType = 'LOCAL'}
		}
		$obj_AccountDetails = [pscustomobject]@{
			ProfileName = $ProfileName
			Domain = $Domain
			ProfileFolder = (($account.LocalPath -split '\\')[2]) #just the userfolder not the full path.
			Type = $AccountType #the type of account (local,temp,roaming etc)
			IsInUse = $Account.Loaded #is the account in use.
			SSID = $account.SID #the SSID of the account
		}
		$null = $ProfileDataArray.Add($obj_AccountDetails) #add account to arraylist.
	}
	Write-Output -InputObject $ProfileDataArray
}
Function Get-ComputerNameGUI
{
	Add-Type -AssemblyName PresentationFramework
	[xml]$PCPrompt = 
	@'
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    ResizeMode="CanResize"
    SizeToContent="WidthAndHeight" 
    MaxWidth="500"
    Title="An Error Has Occurred"
    WindowStartupLocation="CenterScreen">
    <Grid Margin="0,0,0,0">
      <Label VerticalAlignment="Top" HorizontalAlignment="Left" Margin="10,10,10,100" Content="Enter Computer Name of source machine."/>
      <TextBox Name="txtPCNum" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="20,40,10,50" Width="200" Height="30"/>
		  <Button Name="BtnOK" Width="Auto" Height="Auto" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="5,5,65,5" Padding="10" Content="Okay"/>
		  <Button Name="BtnCancel" Width="Auto" Height="Auto" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="5,5,5,5" Padding="10" Content="Cancel"/>
    </Grid>
</Window>
'@
	$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $PCPrompt) #interpret the XAML
	$PCWindow = [windows.markup.xamlreader]::Load($reader) #convert the XAML into a window object
	$txtPCNum = $PCWindow.FindName('txtPCNum') #define a variable holding a control on the form
	$btnOK = $PCWindow.FindName('BtnOK') #define a variable holding a control on the form
	$btnCancel = $PCWindow.FindName('BtnCancel') #define a variable holding a control on the form

	$btnOK.Add_click({ #on clicking ok
			IF(-not [string]::IsNullOrEmpty($txtPCNum.Text) -And $txtPCNum.Text -match '(?:^SOE|^PC|^EQ|^Q\D*[012])\D*\d{5,5}\s*$|^localhost\s*$') #check that entered PC number is not null and that it matches a corporate SOE number
			{
				Set-Variable -Name ComputerName -Value $txtPCNum.Text -Scope 1 #define the computername variable for return
				$PCWindow.Close() #close the form
			}
			Else
			{
				[Windows.MessageBox]::Show('Invalid Computername entered. Please enter a computername (localhost is acceptable)') #throw an error and return to form if invalid entry
			}
	})
	$btnCancel.Add_click({#close the form and return null if cancel is selected
			$PCWindow.Close() 
			Return
	})
	$null = $PCWindow.ShowDialog() #show the dialog
	Write-Output -InputObject $Computername.Trim() #return the computername without trailing spaces.
}
Function New-FilePathObject
{
	Param
	(
		[parameter(Mandatory)][String]$Path
	)
	$fso = New-Object -ComObject Scripting.FileSystemObject #create a filessystem com object. - used to grab the size of files and folders.
	
	$FileSystemObject = Get-Item -Path $Path #get the item at the supplied path (i.e the actual folder or file)
	
	IF($FileSystemObject -is [IO.DirectoryInfo]) #if a folder.
	{
		$size = [math]::round($FSO.GetFolder($path).Size/1mb,3) #determin the size
		$type = 'Folder' #set the type.
	}
	
	$object = [pscustomobject]@{ #create a pscustomobject containing the file/folder info.
		Name = $FileSystemObject.Name
		FullPath = $FileSystemObject.FullName
		Type = $type
		Selected = $FileSystemObject.Name -notmatch 'Redirected'
		Size = $size
	}
	Write-Output -InputObject $object #return the object
}
Function Copy-File
{
	Param
	(
		[Parameter(Mandatory)][String]$Source,
		[Parameter(Mandatory)][String]$Destination
	)
	Start-Sleep -Milliseconds 50
	Add-Type -AssemblyName PresentationFramework
	$SaveLoc = $txtSaveLoc.Text #define the save location
	$k=0 #counter
	write-host $Destination -ForegroundColor Yellow
	While((Test-Path -Path $Destination) -and (Get-Item -Path $Destination) -isnot [IO.DirectoryInfo]) #if there is already a matching file at the destination and it's not a folder.
	{
		$k++ #increment counter
		$Destination = $Destination -replace '(?((?:\s\(\d\))*(\.)(?=[^\\/:*?"<>|\s.]{1,255}$))(?:\s\(\d\))*(\.)(?=[^\\/:*?"<>|\s.]{1,255}$)|(?<!\.[^\\/:*?"<>|\s.]{1,255})$)', " ($k)`$1" #add (1) to the end of the file name - regex above will match'file (1).txt' or file.txt or file and replace either (1)., . or the end of line character with ($k). or just ($k)
	}
	write-host $Source -ForegroundColor Cyan
	write-host $Destination -ForegroundColor Magenta
	IF((Test-path $Destination) -and (Get-Item -Path $Destination) -is [IO.DirectoryInfo])
	{ #if the item being copied exists and is a folder skip.
		Return
	}
	Try
	{
		Write-Verbose 'Initial Try'
		Copy-Item -Path $source -Destination $Destination -ErrorAction Stop #try the copy
		$logobject = [PSCustomObject]@{ #create a entry in the log file for the copied file.
			OriginalLoc = $Source
			Destination = $Destination
		}
		Export-Csv -InputObject $logobject -Path ('{0}\FileList_{1}.csv' -f $SaveLoc, $script:action) -Append -NoTypeInformation #export to log
	}
	Catch [IO.PathTooLongException]
	{#if file path is too long
		Try
		{
			Write-Verbose 'Path to long Try'
			RoboCopy.exe $Source $Destination #try the copy using robocopy - robocopy isn't used by default as it doesn't throw terminating errors to powershell (makes error handling extrememly hard)
			$logobject = [PSCustomObject]@{#create the log entry.
				OriginalLoc = $Source
				Destination = $Destination
			}
			Export-Csv -InputObject $logobject -Path ('{0}\FileList_{1}.csv' -f $SaveLoc, $script:action) -Append -NoTypeInformation #export to log
		}
		Catch
		{#if by some miracle robocoy throws a terminating error - show an error to the client
			Write-Verbose 'Path to long catch'
			$Path = Split-Path -Path $Source -Parent 
			$Name = Split-Path -Path $Source -Leaf
			$Message = "A Path to long Error Occurred while Copying {0} From {1} And the tool was unable to copy the file`r`n`r`nError Message: {2}`r`nCategory: {3}" -f $Name, $Path, $_.Exception.Message, $_.CategoryInfo.Reason 
			$userchoice = Show-UserPrompt -Message $Message -Button3 'Ok'
			Continue
		}
	}
	Catch [IO.DirectoryNotFoundException]
	{ #if the destination directory is not found
		Try
		{
			Write-Verbose 'Directory not found Try'
			$DestinationParent = Split-Path -Path $Destination -Parent #get the parent directory
			New-Item -Path $DestinationParent -ItemType Directory -Force
			Copy-Item -Path $Source -Destination $Destination
			$logobject = [PSCustomObject]@{
				OriginalLoc = $Source
				Destination = $Destination
			}
			Export-Csv -InputObject $logobject -Path ('{0}\FileList_{1}.csv' -f $SaveLoc, $script:action) -Append -NoTypeInformation #export to log
		}
		Catch
		{
			Write-Verbose 'Directory not found Catch'
			$Name = Split-Path -Path $file.Fullname -Leaf
			#copy failed twice-  continue to next file.
			$null = [Windows.MessageBox]::Show("The File {0} has failed to copy.`r`nThe script will now skip over this file and proceed with the next" -f $name,'Retry Failed','OK')
			Continue
		}
	}
	Catch
	{
		Write-Verbose 'Generic Catch'
		#on any other terminating error
		#create message and prompt use to retry/abort/ignore
		$Path = Split-Path -Path $Source -Parent 
		$Name = Split-Path -Path $Source -Leaf
		$Message = "An Error Occurred while Copying {0} From {1}`r`n`r`nError Message: {2}`r`nCategory: {3}" -f $Name, $Path, $_.Exception.Message, $_.CategoryInfo.Reason 
		$userchoice = Show-UserPrompt -Message $Message -Button1 'Retry' -Button2 'Ignore' -Button3 'Abort'
		IF($userchoice -eq 'Abort')
		{
			#on abort terminate backup
			Break
		}
		ElseIF($userchoice -eq 'Ignore')
		{
			#on ignore continue to next file
			Continue
		}
		ElseIF($userchoice -eq 'Retry')
		{
			#on retry try the copy again.
			Try
			{
				Copy-Item -Path $Source -Destination $Destination
				$logobject = [PSCustomObject]@{
					OriginalLoc = $Source
					Destination = $Destination
				}
				Export-Csv -InputObject $logobject -Path ('{0}\FileList_{1}.csv' -f $SaveLoc, $script:action) -Append -NoTypeInformation #export to log
			}
			Catch
			{
				#if copy files a second time let the client know and continue to the next file.
				$null = [Windows.MessageBox]::Show("The File has failed to copy for the second time.`r`nThe script will now skip over this file and proceed with the next",'Retry Failed','OK')
				Continue
			}
		}
	}
}
Function Show-UserPrompt 
{
	#A one to three button form to show a message to the user and get a response back. Returns the content (text) of the selected button which is set when calling the function.
	Param(
		[parameter(Mandatory)][String]$Message, #message to be shown
		[string]$Button1 = $null, #by default show an OK button (Left) 
		[string]$Button2 = $null, #the text for the second (Centre) button (optional)
		[string]$Button3 = 'OK' #the text for the third (Right) button (optional)
	)
	
	#Create the WPF Form using XAML. this seems to be a fairly efficient way of creating GUIs.
	Add-Type -AssemblyName PresentationFramework
	[xml]$ErrorForm = 
	@'
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    ResizeMode="CanResize"
    SizeToContent="WidthAndHeight" 
    MaxWidth="500"
    Title=""
    WindowStartupLocation="CenterScreen">
    <Grid Margin="0,0,0,0">
		<Label Margin="10,10,10,100">
		  <TextBlock TextWrapping="WrapWithOverflow" Text="{Binding Path=Message}"/>
		</Label>
		<StackPanel Orientation="Horizontal" VerticalAlignment="Bottom" HorizontalAlignment="Right" Margin="0,0,5,5" Height="100" Width="Auto" MaxWidth="250">
		  <Button Name="BtnOne" Width="Auto" Height="Auto" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="5,5,5,5" Padding="10" Content="{Binding Path=ButtonOne}"/>
		  <Button Name="BtnTwo" Width="Auto" Height="Auto" HorizontalAlignment="Center" VerticalAlignment="Bottom" Margin="5,5,5,5" Padding="10" Content="{Binding Path=ButtonTwo}"/>
		  <Button Name="BtnThree" Width="Auto" Height="Auto" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="5,5,5,5" Padding="10" Content="{Binding Path=ButtonThree}"/>
    </StackPanel>
    </Grid>
</Window>
'@

	$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ErrorForm) #interpret the XAML
	$ErrorWindow = [windows.markup.xamlreader]::Load($reader) #Load the form.
	
	#create a custom object that is later bound as the data context of the form for databinding purposes.
	$ErrorDetails = [PsCustomObject]@{ 
		Message     = $Message
		ButtonOne   = $Button1
		ButtonTwo   = $Button2
		ButtonThree = $Button3
	}
	
	#declare the items of the form as variables for later use
	$btnOne = $ErrorWindow.Findname('BtnOne')
	$btnTwo = $ErrorWindow.Findname('BtnTwo')
	$btnThree = $ErrorWindow.Findname('BtnThree')
	
	
	if([string]::IsNullOrEmpty($Button1)) #if button 1 parameter is a null or empty string
	{
		$btnOne.Visibility = 'Hidden' #hide the button
	}
	Else
	{
		$btnOne.Visibility = 'Visible' #else Show the button on the form.
	}
	if([string]::IsNullOrEmpty($Button2)) #repeat for button 2
	{
		$btnTwo.Visibility = 'Hidden'
	}
	Else
	{
		$btnTwo.Visibility = 'Visible'
	}
	if([string]::IsNullOrEmpty($Button3)) #repeat for button 3
	{
		$btnThree.Visibility = 'Hidden'
	}
	Else
	{
		$btnThree.Visibility = 'Visible'
	}
	
	$ErrorWindow.DataContext = $ErrorDetails #set the datacontext for the databindings that are specified in the XAML
	
	[string]$Return = '' #create an empty return variable of type string.
	$btnOne.Add_Click({
			#on button click
			$ErrorWindow.Close() #close the form
			Set-Variable -Name 'Return' -Value $btnOne.Content -Scope 1 #set the value of the return variable to the text of the button (the -scope 1 parameter makes it modify the variable in the parent scope.
	})
	$btnTwo.Add_Click({
			$ErrorWindow.Close()
			Set-Variable -Name 'Return' -Value $btnTwo.Content -Scope 1 #set the value of the return variable to the text of the button (the -scope 1 parameter makes it modify the variable in the parent scope.
	})
	$btnThree.Add_Click({
			$ErrorWindow.Close()
			Set-Variable -Name 'Return' -Value $btnThree.Content -Scope 1 #set the value of the return variable to the text of the button (the -scope 1 parameter makes it modify the variable in the parent scope.
	})
	$null = $ErrorWindow.ShowDialog() #show the form now that the databindings and click actions have been set. function execution pauses here until the form is closed.
	Return $Return #return the selected option to the calling variable / function.
}
Function Format-RegexSafe 
{
	<# 
			.SYNOPSIS
			Function/tool to return a Regex safe version of an input string

			.DESCRIPTION
			Some comparison operators in Powershell like -match use regular expressions. 
			The comparison will not produce the expected result if the string being compared to contains 
			any of the 14 PS Regex metacharacters. One way to solve this issue is to escape any of the 
			Regex metacharacters in the string we're comparing to. This is exactly what this script does.

			.PARAMETER String
			The string to be processed by escaping every Regex metacharacter in it.

			.EXAMPLE
			RegexSafe -String 'Hi there'
			Returns the same string since it has no Regex metachacters. 

			.EXAMPLE
			RegexSafe 'vHost01d(Web)-D.vhdx'
			Returns 'vHost01d\(Web\)-D\.vhdx', which is a string that can be used with -match 
			and other Powershell comparison operators to produce accurate results.

			.EXAMPLE
			'\\server\share\folder\file.ext' | RegexSafe
			Returns '\\\\server\\share\\folder\\file\.ext', which is a string that can be used 
			with -match and other Powershell comparison operators to produce accurate results. 

			.EXAMPLE
			'\\Serv1\sh1\file1.ext','\\Serv1\sh2\file2.ext','\\Serv2\sh1\file3.ext' -match 'sh2\file2.ext'
			Returns no matches since 'sh2\file2.ext' includes Regex metacharacters. 

			However:
			'\\Serv1\sh1\file1.ext','\\Serv1\sh2\file2.ext','\\Serv2\sh1\file3.ext' -match (RegexSafe 'sh2\file2.ext')
			Returns '\\Serv1\sh2\file2.ext' which is the desired result.

			.LINK
			https://superwidgets.wordpress.com/category/powershell/

			.INPUTS
			String

			.OUTPUTS
			String consisting of input string with escaped Regex metacharacters if any existed.

			.NOTES
			Function by Sam Boutros
			v1.0 - 12/26/2014

	#>

	[CmdletBinding(ConfirmImpact = 'Low')] 
	Param(
		[Parameter(Mandatory,
				ValueFromPipeLine,
				ValueFromPipeLineByPropertyName,
		Position = 0)]
		[String]$String #string of text to be made regex safe
	)
	$Output = '' #create empty output string
	$RegexMeta = '()\^$.*+|?[]{}' #characters to be made safe in the string. These are all characters that regex interprets as commands/syntax elements unless escaped.
	for ($i = 0; $i -lt $String.Length; $i++ ) #for each character in string.
	{
		for ($j = 0; $j -lt $RegexMeta.Length; $j++ ) #for each character in the Meta
		{
			if ($String[$i] -eq $RegexMeta[$j])  #if the string character and the Meta character match
			{ 
				#Write-Verbose -Message ('found {0}' -f ($String[$i]))  
				$Output += '\'#add an escape character to output
			} 
		}
		$Output += $String[$i] #add the current character to output
	}
	$Output #return the output
}
Function Show-DataGrid
{
	Param
	(
		[Parameter(Mandatory,Position=0,ValueFromPipeline)]$InputObjects,
		[Parameter(Position=1)][String]$Message='',
		[Parameter(Position=2)][String]$WindowTitle='DataGridView',
		[Switch]$MultiSelect
	)
	Begin
	{
		Add-Type -AssemblyName presentationframework
		$DataTable = New-Object -TypeName Data.datatable
		Function Save-File 
		{#prompt user on where to save file - defaults to C:\localdata.
			[CmdletBinding()]
			param([string]$initialDirectory = "$env:SystemDrive\LocalData")
			Add-Type -AssemblyName PresentationFramework
			$SaveFileDialog = New-Object -TypeName Microsoft.Win32.SaveFileDialog #create the savefiledialog
			$SaveFileDialog.InitialDirectory = $initialDirectory #set the default location to C:\localdata. 
			$SaveFileDialog.Filter = 'Csv files (*.csv)| *.csv' #set the filter to only show csv files
			$null = $SaveFileDialog.ShowDialog() #show the dialog
			Write-Output -InputObject $SaveFileDialog.FileName #return the selected file path
		}
		
		Function Get-Type 
		{
			#a support function for Show-Datagrid - used to get the column property type.
			param([Parameter(Mandatory)][String]$type)  
			$types = 
			@( 
				'System.Boolean', 
				'System.Byte[]', 
				'System.Byte', 
				'System.Char', 
				'System.Datetime', 
				'System.Decimal', 
				'System.Double', 
				'System.Guid', 
				'System.Int16', 
				'System.Int32', 
				'System.Int64', 
				'System.Single', 
				'System.UInt16', 
				'System.UInt32', 
				'System.UInt64'
			) 
			If ( $types -contains $type ) 
			{
				Write-Output -InputObject $type
			} 
			Else 
			{
				Write-Output -InputObject 'System.String'
			} 
		}
		$first = $true
	}
	Process
	{
		IF($first){
			$properties = ($InputObjects|Select-Object -First 1).psobject.get_properties()
			Foreach($property in $properties)
			{
				$Column = New-Object -TypeName Data.DataColumn
				$Column.ColumnName = $property.Name
				$Column.DataType = Get-Type -type $property.TypeNameOfValue
				$DataTable.Columns.Add($Column)
			}
			$first = $false
		}
		Foreach($Object in $InputObjects)
		{
			$DataRow = $DataTable.NewRow()
			Foreach($property in $object.PSObject.Get_properties())
			{
				IF($properties.Gettype().isArray)
				{
					$DataRow.Item($property.Name) = $property.Value|ConvertTo-Xml -As String -NoTypeInformation -Depth 1
				}
				Else
				{
					$DataRow.Item($property.Name) = $property.Value
				}
			}
			$DataTable.Rows.Add($DataRow)
		}
	}
	End
	{
		[XML]$XAML = 
		@'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Width="Auto"
    Height="Auto"
    WindowStartupLocation="CenterScreen">
    <Grid>
      <Label Name="lblMessage" VerticalAlignment="Top" HorizontalAlignment="Left" Width="Auto" Height="30" Margin="10"/>
      <DataGrid Name="dgItems" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" Width="Auto" Height="Auto" Margin="10,50,10,50" AutoGenerateColumns="True" AlternatingRowBackground="Gainsboro" BorderThickness="1"/>
      <Button Name="btnSelect" VerticalAlignment="Bottom" HorizontalAlignment="Right" Width="60" Height="20" Content="Select" Margin="0,0,135,10"/>
      <Button Name="btnExport" VerticalAlignment="Bottom" HorizontalAlignment="Right" Width="60" Height="20" Content="Export" Margin="0,0,70,10"/>
      <Button Name="btnCancel" VerticalAlignment="Bottom" HorizontalAlignment="Right" Width="60" Height="20" Content="Cancel" Margin="0,0,5,10"/>
    </Grid>
</Window>
'@

		$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $XAML)  #convert xaml 
		$Window = [Windows.Markup.XamlReader]::Load($reader)
		$btnSelect = $Window.FindName('btnSelect')
		$btnExport = $Window.FindName('btnExport')
		$btnCancel = $Window.FindName('btnCancel')
		$dgitems = $window.Findname('dgItems')
		$lblMessage = $window.FindName('lblMessage')
		
		$dgitems.ItemsSource = $DataTable.DefaultView
		$lblMessage.Content = $Message
		$Window.Title = $WindowTitle
		
		$dgitems.SelectionMode='Single'
		IF($MultiSelect)
		{
			$dgitems.SelectionMode='Extended'
		}
		
		$btnCancel.Add_Click({$window.Close()})
		
		$btnSelect.Add_Click({
				#on click
				$selected = @()
				Foreach($Row in $dgitems.SelectedItems) 
				{
					#foreach of the selected rows
					$out = New-Object -TypeName PsObject #create ps object
					Foreach($Header in $dgitems.Columns.Header)
					{
						$out|Add-Member -MemberType NoteProperty -Name $Header -Value $row.$header
					}
					$Selected += $out #add the Psobject to the return
				}
				$Window.close() #close the form
				Set-Variable -Name Return -Value $selected -Scope 1
		})
		
		$btnExport.Add_Click({
				$DataTable|Export-Csv -LiteralPath (Save-File) -NoTypeInformation #export datatable to selected file.
				$Window.close() #close form
		})
		
		$null = $Window.ShowDialog()
		Write-Output -InputObject $Return
	}
}
Function Set-InitialStateBackup
{
	#sets the initial state of the GUI when backup is the selected action.
	Add-Type -AssemblyName PresentationFramework
	$fso = New-Object -ComObject Scripting.FileSystemObject #create a filessystem com object. - used to grab the size of files and folders.
	$script:FilePaths = @() #script wide variable to hold the files/folders shown in the listview of the GUI.
	$OS         = Get-CimInstance -Query 'Select Caption FROM WIN32_OperatingSystem'|Select-Object -ExpandProperty Caption #get the OS Caption (Windows 7 or Windows 10)
	#add the user shell folders to the filepath lists 
	$script:FilePaths += New-FilePathObject -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop))
    $script:FilePaths += New-FilePathObject -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::Documents))
	$script:FilePaths += New-FilePathObject -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::Favorites))

	IF(Test-path -Path "$env:APPDATA\Microsoft\Signatures") #outlook signatures
	{
		$script:FilePaths += New-FilePathObject -Path "$env:APPDATA\Microsoft\Signatures"
	}
	IF(Test-Path -Path "$env:SystemDrive\User") #same for C:\User
	{
		$script:FilePaths +=  New-FilePathObject -Path "$env:SystemDrive\User"
		
	}
    IF(Test-Path -Path "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms") #same for Quick Access
	{
		$QuickAccess = New-FilePathObject -Path "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms"
		$script:FilePaths += $QuickAccess
	}
	IF(Test-Path -Path "$env:SystemDrive\Temp") #same for C:\temp
	{
		$Temp = New-FilePathObject -Path "$env:SystemDrive\Temp"
		$script:FilePaths += $Temp
	}
	IF(Test-Path -path "$env:APPDATA\google\googleearth\myplaces.kml") #google Earth default KML
	{
		$KML = New-FilePathObject -Path "$env:APPDATA\google\googleearth\myplaces.kml"
		$script:FilePaths += $KML
	}
	IF(Test-Path -Path "$ENV:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks") #see if there is a Chrome Bookmarks file
	{
		$script:FilePaths += New-FilePathObject -Path "$ENV:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks"
	}	
	
	[float]$script:backupSize = 0 #instatiate a backup size variable of type Float
	$script:backupSize = ($script:FilePaths|Where-Object{$_.Selected -eq $true}|Measure-Object -Property Size -Sum).Sum
	
	$lblrequiredSpace.Content = ("Required Space for Backup:`t{0} Gb" -f ([math]::round($script:backupSize/1kb,3))) #update label on the form with the backup size in Gb

	$Script:lvview = [Windows.Data.ListCollectionView]$script:FilePaths #create a listcollectionview from FilePaths array
	$lvwFileList.ItemsSource = $Script:lvview #bind the listcollectionview to the listview.(i.e. add the items to the GUI)
}
Function Set-InitialStateAdminBackup
{
	Add-Type -AssemblyName PresentationFramework
	
	$Script:RemoteInfo = New-Object -TypeName RemoteInfo #create new remoteinfo object (Custom C# class defined at script launch)
	$Script:RemoteInfo.Computername = Get-ComputerNameGUI #prompt for the computer name
	$fso = New-Object -ComObject Scripting.FileSystemObject #create a file system com object.
	
	IF(-not (Test-Connection -ComputerName $Script:RemoteInfo.Computername -Count 1 -Quiet) -or -not (Test-Path -Path ('\\{0}\C$' -f $Script:RemoteInfo.Computername))) 
	{
		#see if remote machine is online and admin share is available. if not show an error and stop processing.
		[Windows.Messagebox]::Show(('Cannot connect to {0}' -f $Script:RemoteInfo.Computername))
		Return
	}
	$Script:Filepaths = @() #create filepaths array
	IF($Script:RemoteInfo.Computername -match 'localhost') #if the computername is localhost the cim session needs to be created slightly differently than if it's a netbios name.
	{
		$CimSession = New-CimSession -ComputerName $Script:RemoteInfo.Computername 
	}
	Else
	{
		$CimSession = New-CimSession -ComputerName ([net.dns]::GetHostByName($Script:RemoteInfo.Computername).Hostname)
	}
	$OS = Get-CimInstance -Query 'SELECT Caption FROM Win32_OperatingSystem' -CimSession $CimSession|Select-Object -ExpandProperty Caption #get the Os type.
	
	$profile = Get-SSID -Session $CimSession | Show-DataGrid -WindowTitle 'Select User' #prompt user to select the profile to run against.
	IF(-not $profile){return} #if no profile is selected stop processing.
	
	$Script:RemoteInfo.ProfileFolder = '\\{0}\C$\Users\{1}' -f $Script:RemoteInfo.Computername, $profile.ProfileFolder #define the profile folder for the select account.
	$Script:RemoteInfo.IsInUse = $profile.IsInUse #is the profile in use (i.e. is the user hive loaded in registry)
	$Script:RemoteInfo.ProfileName = $profile.ProfileName #the name of the profile.
	$Script:RemoteInfo.SSID = $profile.SSID #SSID of the profile. 
	$names = @('Desktop','Documents','Favorites') #define the names of the special folders.
	IF($Script:RemoteInfo.IsInUse)
	{#if profile is loaded.
		Try
		{ 
			$Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Users',$Script:RemoteInfo.Computername) #open HKey_Users on the remote machine.
			$Script:RemoteInfo.UserHive = $Registry.OpenSubKey($Script:RemoteInfo.SSID,$True) #open the user hive/
		}
		Catch
		{
			Write-Error -Message $_
			[Windows.Messagebox]::Show('Failed to open remote Registry hive')
			Return
		}
	}
	Else
	{#if profile is not loaded.
		$keyName = 'HKU\{0}_remote' -f $Script:RemoteInfo.SSID #define the key name
		$keyLocation = '{0}\ntuser.dat' -f $Script:RemoteInfo.ProfileFolder #path to user hive (ntuser.dat file)
		Try
		{
			Reg.exe Load $keyName $keyLocation #load the hive on the local machine.
			New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_Users -ErrorAction Stop #create a PSDrive to Hkey_Users 
			$Script:RemoteInfo.UserHive = get-item -Path ('HKU:\{0}_remote' -f $Script:RemoteInfo.SSID) #open the user hive.
		}
		Catch
		{
			Write-Error -Message $_
			[Windows.Messagebox]::Show('Failed to open remote Registry hive')
			Return
		}
	}
	$regkey = $Script:RemoteInfo.UserHive.OpenSubKey('Software\\microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders',$True) #open the particular reg key we need
	Foreach($name in $names) #foreach of the user shell folders.
	{
		$path = $RegKey.GetValue($name) -replace ('C:\\Users\\{0}' -f $env:USERNAME), $Script:RemoteInfo.ProfileFolder #get the path to the folder and replace C:\users\sys... with the remote path to the profile. this is needed because the registry values use enviroment variables for the paths which when resolved match to the user running the script.
		$Script:Filepaths += new-FilePathObject -Path $path #create the entry for the list view.
	}
	$adminshare = '\\{0}\C$' -f $script:remoteInfo.Computername #define the admin share.
	$AppdataLocal = '{0}\Appdata\Local' -f $script:RemoteInfo.ProfileFolder #define appdata\local
	$AppdataRoaming = '{0}\Appdata\Roaming' -f $script:RemoteInfo.ProfileFolder #define appdata\roaming.
	
	$Paths = #default paths we want to add if they exist.
	@(
		('{0}\User' -f $adminshare),
		('{0}\Temp' -f $adminshare),
        ('{0}\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms' -f $AppDataRoaming),
		('{0}\Google\GoogleEarth\MyPlaces.kml' -f $AppDataRoaming),
		('{0}\Microsoft\Signatures' -f $AppdataRoaming),
		('{0}\Google\Chrome\User Data\Default\Bookmarks' -f $AppdataLocal)
	)
	foreach ($Path in $Paths) #for each of the paths.
	{
		IF(Test-Path -Path $path) #test the location
		{
			$Script:Filepaths += New-FilePathObject -Path $Path #if it exists add to filepaths.
		}
	}
	
	$script:backupSize = 0 #instatiate a backup size variable of type Float
	$script:backupSize = ($script:FilePaths|Where-Object{$_.Selected -eq $true}|Measure-Object -Property Size -Sum).Sum
	
	$lblrequiredSpace.Content = ("Required Space for Backup:`t{0} Gb" -f ([math]::round($script:backupSize/1kb,3))) #update label on the form with the backup size

	$Script:lvview = [Windows.Data.ListCollectionView]$script:FilePaths #create a listcollectionview from FilePaths array
	$lvwFileList.ItemsSource = $Script:lvview #bind the listcollectionview to the listview.(i.e. add the items to the GUI)
}
Function Set-InitialStateRestore 
{
	#set the initial state of the GUI when restore is the selected action.
	Add-Type -AssemblyName PresentationFramework
	Set-SaveLocation #force the user to set the location of the previous backup straight away.
	$script:FilePaths = @() #create the file paths variable. - this doesn't exist until created but one of the initial state functions.
	$fso = New-Object -ComObject Scripting.FileSystemObject
	Foreach($Item in (Get-ChildItem -Path $txtSaveLoc.Text -Exclude 'FileList*.csv', 'Drives.csv', 'Printers.txt')) #foreach of the folders and files in the backup path excluding the files generated by the backup.
	{
		IF($Item -is [IO.DirectoryInfo]) #if the item is a folder add to filepaths array as a folder
		{
			$script:FilePaths += New-FilePathObject -Path $Item.Fullname
		}
		Else
		{
			#else add as a file
			$script:FilePaths += New-FilePathObject -Path $Item.Fullname
		}
	}
	
	$btnAddLoc.IsEnabled = $false #disable the button to add a location
	$Script:lvview = [Windows.Data.ListCollectionView]$script:FilePaths #create the listviewcollection
	$lvwFileList.ItemsSource = $Script:lvview #bind listviewcollection to listview.
}
Function Get-FilePath 
{
	#Launches a folder browser dialog to select a file path - default location is the homedrive.
	Add-Type -AssemblyName System.Windows.Forms
	$SysDrive = "$env:SystemDrive\" #get the system drive
	$FolderSel = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog #create the folder browser dialog
	$FolderSel.SelectedPath = $env:HOMEDRIVE #set the default path to homedrive
	$FolderSel.ShowNewFolderButton = $true #show the button to create a new folder
	IF($Script:Action -eq 'Backup') #if the selected action is backup
	{
		#set the descriptive text on the dialog
		$FolderSel.Description = "Please Select A Folder to save the backup to.`r`n Please Note that a Backup Folder will be created at the selected location as part of the process"
	}
	Elseif($Script:Action -eq 'Restore') #else if action is restore
	{
		#set the descriptive text on the dialog
		$FolderSel.Description = 'Please Select the folder containing the backup.'
	}
	$response = $FolderSel.ShowDialog() #show the dialog - function exection pauses here.
	IF($response -eq 'Cancel'){Return}
	While($FolderSel.SelectedPath -eq $SysDrive -or $FolderSel.SelectedPath -eq $null) #while the select path is the system drive root or a null value
	{
		$prompt = [Windows.Forms.MessageBox]::Show(('{0} is not a valid selection Please select or create a folder.' -f $FolderSel.SelectedPath),'',0x1) #show a prompt to user advising of issue.
		if($prompt -eq 'Cancel') #if they hit cancel on the prompt
		{
			Break #stop execution of function
		}
		$Response = $FolderSel.ShowDialog() #else show the dialog again.
		IF($response -eq 'Cancel'){Return}
	}
	Return $FolderSel.SelectedPath #return the selected folder path.
}
Function Set-GPupdate
	{	
	write-host "Inside Set-GPupdate" -ForegroundColor Cyan
	# Start a new process for gpupdate
	$process = Start-Process -FilePath "gpupdate" -NoNewWindow -PassThru

	# Wait for the process to finish
	$process.WaitForExit()

	# Check the exit code
	if ($process.ExitCode -eq 0) {
		Write-Output "Group Policy update was successful."
	} else {
		Write-Output "Group Policy update failed with exit code $($process.ExitCode)."
	}	
}
Function Set-SaveLocation 
{
	#set the save location for the backup/restore
	IF($Script:Action -eq 'Backup')  #if we're running a backup
	{
		$selectedSaveLoc = Get-FilePath
		IF($selectedSaveLoc -eq $null){return}
		$date = get-date -Format 'dd-MM-yy'
		IF($Script:Admin)
		{
			$txtSaveLoc.text = ('{0}\Backup_{1}_{2}' -f $selectedSaveLoc,$date,$Script:RemoteInfo.ProfileName) -replace '\\\\', '\' #prompt client for a save location - replace any double \ with a single \ - just in case
		}
		Else
		{
			$txtSaveLoc.text = ('{0}\Backup_{1}_{2}' -f $selectedSaveLoc,$date,$env:USERNAME) -replace '\\\\', '\' #prompt client for a save location - replace any double \ with a single \ - just in case
		}
		$driveLetter = Split-Path -Path $txtSaveLoc.Text -Qualifier #get the drive letter of the selected path
		$Script:DriveSpace = [Math]::Round((Get-CimInstance -Query "SELECT * FROM Win32_LogicalDisk Where DeviceID='$driveLetter'"|Select-Object -ExpandProperty FreeSpace)/1Gb,3) #get the space available on the drive and round to three places
		$lblFreeSpace.Content = ("Free Space on selected Drive:`t{0} Gb" -f $Script:DriveSpace) #update label with how much space is available on the drive.
	}
	ElseIF($Script:Action -eq 'Restore') #if we're running a restore action
	{
		$txtSaveLoc.text = (Get-FilePath) -replace '\\\\', '\' #prompt client for the file path of the restore
		$fso = New-Object -ComObject Scripting.FileSystemObject #create and File system object to get the size of the backup folder
		$script:backupSize = $fso.Getfolder($txtSaveLoc.text).Size #get the size of the folder
		$bsize = [math]::Round($($script:backupSize/1mb),3) #round to 3 places
		$lblrequiredSpace.Content = ("Backup Size is:`t{0} mb" -f $bsize) # update label with size of the backup to be restored
		$Script:DriveSpace = [Math]::Round((Get-CimInstance -Query "SELECT * FROM Win32_LogicalDisk Where DeviceID='$env:SystemDrive'"|Select-Object -ExpandProperty FreeSpace)/1Gb,3) #get the space available on the system drive and round to three places.
		$lblFreeSpace.Content = ("Free Space on System Drive:`t{0} Gb" -f $Script:DriveSpace) #update label with available space.
	}
}
Function Add-BackupLocation
{
	Add-Type -AssemblyName PresentationFramework
	$userchoice = Show-UserPrompt -Message 'Are we adding a file or a folder?' -Button2 'File' -Button3 'Folder' #prompt user and ask if we're adding a file or a folder to the backup.

	IF($userchoice -eq 'File')
	{
		#if file
		$FileSel = New-Object -TypeName Microsoft.Win32.OpenFileDialog #create open file dialog
		$FileSel.InitialDirectory = $env:SystemDrive #set default directory to system drive
		$FileSel.Multiselect = $true #multiselect is true so that multiple files can be added.
		$fso = New-Object -ComObject Scripting.FileSystemObject #create file system object
		$dialogresponse = $FileSel.ShowDialog() #show the dialog
		IF($dialogresponse -eq $false){Return}
		Foreach($file in $FileSel.FileNames) #for each selected file create an object and add to the filepaths array.
		{
			$included = $false #used later.
			$parent = split-path $file -Parent #get the parent path
			$leaf = split-path $file -Leaf #get the last element of the path (folder or filename)
			Foreach($Path in ($lvwFileList.Items).FullPath) #for each row in the listview.
			{
				IF($parent -match (Format-RegexSafe $Path)) #if parent matches the path (i.e. if parent is a subfolder)
				{
					Show-UserPrompt -Message ('{0} is already in the list under {1}' -f $leaf, $Path) -Button3 'OK' #show a message that it's already in the list
					$included = $true #set included to true
				}
			}
			if(!$included) #if the select file is not already included under another folder.
			{
				$script:FilePaths +=  New-FilePathObject -Path $file #add to the listview.
			}
		}
	}
	ElseIF($userchoice -eq 'Folder')
	{
		#if folder
		$fso = New-Object -ComObject Scripting.FileSystemObject #create file system object
		$location = Get-FilePath #get folder location resuing get-filepath
		IF([string]::IsNullOrEmpty($location)){return}
		Foreach($Path in ($lvwFileList.Items).FullPath)#for each row in the listview.
		{
			IF($location -match (Format-RegexSafe $Path)) #if the selected location matches the path (i.e. if parent is a subfolder)
			{
				$userchoice = Show-UserPrompt -Message ('{0} is already in the list under {1}' -f $location, $Path) -Button3 'OK' #show a message that it's already in the list
				Return #stop processing
			}
		}
		$script:FilePaths +=  New-FilePathObject -Path $location #else add it to the listview
	}
	Else
	{#if cancel.
		Return
	}
	$script:backupSize = ($script:FilePaths|Where-Object{$_.Selected -eq $true}|Measure-Object -Property Size -Sum).Sum #calculate the new backup size.
	$bsize = [math]::Round($($script:backupSize/1kb),3) #round out new backup size to 3 places
	$lblrequiredSpace.Content = ("Required Space for Backup:`t{0} Gb" -f $bsize) #update label
	$Script:lvview = [Windows.Data.ListCollectionView]$script:FilePaths #update listcollectionview
	$lvwFileList.ItemsSource = $Script:lvview #assign view to form.
}
Function Start-Backup
{
	Add-Type -AssemblyName PresentationFramework
	$OS = Get-WmiObject -Class Win32_OperatingSystem -Property Caption|Select-Object -ExpandProperty Caption #get the OS type
	Show-ProgressBar -Status 'Initializing Backup' -ProgressValue 0 -SubStatus 'Creating Backup Folders' -SubProgressValue 0 #update progress bar
	#Create backup Folder.
	$SaveLoc = $txtSaveLoc.Text #define the saveloc variable as the selected folder/path in the backup location text field.
	$null = New-Item -Path $SaveLoc -ItemType Directory #create the directory

	Update-ProgressBar -Status 'Initialization Complete' -ProgressValue 0 -SubStatus ' ' -SubProgressValue 0 #update progress bar
	$i = 0 #counter value
	$FolderbackupCount = ($lvwFileList.SelectedItems|Where-Object{$_.Type -match 'Folder'}).Count #get a count of the root folders to be backedup.
	$FilebackupCount = ($lvwFileList.SelectedItems|Where-Object -FilterScript {
			$_.Type -Match 'File'
	}).Count #get a count of the selected files to be backed up.
	$totalBackupCount = $FolderbackupCount + $FilebackupCount #define the total number of items to be backed up - used for progress bar.
	
	IF($chk_networkDrives.ISChecked) #if backing up drive mappins
	{
		$totalBackupCount += 1 #add 1 to the step count 
	}
	IF($chk_Printers.ISChecked) #if backing up printer mappings
	{
		$totalBackupCount += 1
	}
	$step = 0
	Foreach($Row in $lvwFileList.SelectedItems|Where-Object{$_.Type -Match 'Folder'})
	{
		# for each folder selected for backup
		$i++
		New-Item -Path ('{0}\{1}' -f $SaveLoc, $Row.Name) -ItemType Directory #create a matching directory in the backup directory
		Update-ProgressBar -Status ('Creating Backup Folders step {0} of {1}' -f $step, $totalBackupCount) -ProgressValue 0 -SubStatus $Row.Name -SubProgressValue ($i/$FolderbackupCount*100) #update progress bar
	}
	If($chk_networkDrives.IsChecked) #if getting network drives
	{
		Update-ProgressBar -Status ('Collecting Network Drive Mappings step {0} of {1}' -f ($step+1), $totalBackupCount) -ProgressValue ($step/$totalBackupCount*100) -SubStatus ' ' -SubProgressValue 0 #update progress bar
		IF($script:Admin)
		{#if running as admin
			IF($Script:RemoteInfo.UserHive -eq $null)
			{#if unable to connect to userhive - show error and skip network drives and printers.
				$null = [Windows.MessageBox]::Show("Unable to access remote user's registry hive. Unable to populate mapped drives and printers.",'Registry Hive load Failed','OK')
				$chk_Printers.isChecked = $false
			}
			Else
			{#if userhive is present.
				$DriveLetters = $Script:RemoteInfo.UserHive.OpenSubkey('Network').GetSubKeyNames() #get the drive letters
				Foreach($drive in $DriveLetters) #foreach drive letter
				{
					$object = [PSCustomObject]@{
							Name=$Drive
							ProviderName=$Script:RemoteInfo.UserHive.OpenSubkey(('Network\{0}' -f $drive)).GetValue('RemotePath') #get the path it's mapped to
						}
					$object|Export-Csv -Path $SaveLoc\Drives.csv -Append -NoTypeInformation #export drive info to CSV
				}
			}
		}
		Else
		{#if not admin
			Get-WmiObject -Class 'Win32_MappedLogicalDisk' |
			Select-Object -Property Name, providername |
			Export-Csv -Path $SaveLoc\drives.csv -NoTypeInformation #get the mapped network drives and export the drive letter and path to a csv in the backup location
			$step ++ #increment step count
		}
	}
	If($chk_Printers.IsChecked) #if grabbing the printer mappings
	{
		Update-ProgressBar -Status ('Collecting Printer Mappings step {0} of {1}' -f ($step+1), $totalBackupCount) -ProgressValue ($step/$totalBackupCount*100) -SubStatus ' ' -SubProgressValue 0 #update progress bar
		IF($script:Admin)
		{#if running as admin (in case network drives weren't ticked)
			IF($Script:RemoteInfo.UserHive -eq $null) #check for userhive
			{
				$null = [Windows.MessageBox]::Show("Unable to access remote user's registry hive. Unable to populate mapped drives and printers.",'Registry Hive load Failed','OK')
			}
			Else
			{#get printer info from regedit.
				$printers = $Script:RemoteInfo.UserHive.OpenSubkey('Printers\Connections').GetSubKeyNames() -replace ',','\'
				$printers|Set-Content -Path $SaveLoc\Printers.txt
			}
		}
		Else{
			Get-WmiObject -Query "Select * FROM Win32_Printer WHERE Local=$false" |
			Select-Object -ExpandProperty Name |
			Set-Content -Path $SaveLoc\printers.txt  #get printer mappings using WMI and export to txt file in backup folder
			$step ++ #increment step count
		}
	}
	Foreach($Row in $lvwFileList.SelectedItems)
	{
		#foreach selected item in the list view
		Update-ProgressBar -Status ('Copying {0} step {1} of {2}' -f $Row.Name, ($step+1), $totalBackupCount) -ProgressValue ($step/$totalBackupCount*100) -SubStatus ' ' -SubProgressValue 0 #update progress bar
		
		IF($Row.Type -match 'Folder') #if Row is a folder
		{
			$Source = $Row.FullPath.tolower() #take the fullpath and make it lowercase
			$regexSaveLoc = Format-RegexSafe -String $SaveLoc #regex safe the save location 
			$files = Get-ChildItem -Path $Source -Recurse |Where-Object{$_.FullName -notmatch "$regexSaveLoc|.*\.(pst|ost)"} #get all sub folders and files in the location. excluding psts, osts and the selected save location.
			$FileCount = $files.Count #count the files (used for progress bar)
			$filenumber = 0 #counter
			$Destination = ('{0}\{1}' -f $SaveLoc, $Row.Name) #set the destination
			$logobject = [PSCustomObject]@{
				OriginalLoc = $Source
				Destination = $Destination
			}
				Export-Csv -InputObject $logobject -Path $SaveLoc\FileList_$script:Action.csv -Append -NoTypeInformation #append to the log file.

			Foreach($file in $files)
			{
				#for each of the sub folders and files.
				$Filename = $file.Fullname.tolower().replace($Source,'') #filename = the fullpath of the file with the source path removed.
				$DestinationFile = ($Destination+$Filename) #appened the file name to the destination folder to create the file's destination path.
				Update-ProgressBar -Status ('Copying {0} step {1} of {2}' -f $Row.Name, ($step+1), $totalBackupCount) -ProgressValue ($step/$totalBackupCount*100) -SubStatus ('Copying {0}' -f $file.FullName) -SubProgressValue ($filenumber/$FileCount *100) #update progress bar.
				Copy-File -Source $File.Fullname -Destination $DestinationFile
				$filenumber++ #increment counter
			}
		}
		Else
		{
			#row is a file.
			Update-ProgressBar -Status ('Copying {0} step {1} of {2}' -f $Row.Name, ($step+1), $totalBackupCount) -ProgressValue ($step/$totalBackupCount*100) -SubStatus ' ' -SubProgressValue 0 #update progress bar
			$DestinationFile = ('{0}\{1}' -f $SaveLoc, $Row.Name)
			Copy-File -Source $row.Fullpath -Destination $DestinationFile
		}
		$step++ #increment overall counter.
	}
	$script:Hash_ProgressBar.Form.Dispatcher.Invoke([Action][ScriptBlock]::Create({
				$script:Hash_ProgressBar.Form.Close()
	})) #close the progress bar
	IF($script:Admin)
	{#if running as admin
		IF((Get-PSDrive -Name HKU -ErrorAction SilentlyContinue) -ne $null) #check if a psdrive HKU exists
		{
			Remove-PSDrive -Name HKU #remove the PS Drive
			start-sleep -Seconds 1
			Reg.exe Unload ('HKU\{0}_remote' -f $Script:RemoteInfo.SSID) #unload the user hive.
		}
	}
}
Function Start-Restore
{
	#function for the restoration of data.
	$SaveLoc = $TxtSaveLoc.Text #set the saveloc varaible to the selected backup location
	$fileList = Import-Csv -Path $SaveLoc\FileList_Backup.csv #import the filelist csv (log file) from previous backup. - log consists of the original file location and where it was backed up to.

	$OS = Get-WmiObject -Class Win32_OperatingSystem -Property Caption|Select-Object -ExpandProperty Caption #get the OS type
	$SpecialFolders = @() #create a specialfolders array
	$SpecialFolders += [PSCustomObject]@{
		Name     = 'Desktop'
		FullPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
	}
	$SpecialFolders += [PSCustomObject]@{
		Name     = 'Documents'
		FullPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Documents)
	}
	$SpecialFolders += [PSCustomObject]@{
		Name     = 'Favorites'
		FullPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Favorites)
	}
	
	$count = $lvwFileList.SelectedItems.Count + 3 #create a count of actions (used for progress bar)
	$i = 0 #counter
	
	Foreach($Row in $lvwFileList.SelectedItems) #for each of the selected items.
	{
		$i++ #increment counter
		IF($SpecialFolders.Name -match ($Row.Name -replace 'Redirected') -and $Row.Type -eq 'Folder') #if the name of the item matches one of the special folders.
		{
			Update-ProgressBar -Status ('Restoring items from {0} Step {1} of {2}' -f $Row.Name, $i, $count) -ProgressValue 0 -SubStatus ' ' -SubProgressValue 0 #update progress bar
			$Destination = ($SpecialFolders|Where-Object -FilterScript {
					#destination equals the special folder path.
					$Row.Name -match $_.Name
			}).FullPath
			$Source = $Row.FullPath.ToLower() #source equals backup location to lower case
			$files = Get-ChildItem -Path $Source -Recurse #get all child files and folders.
			$filenumber = 0 #counter
			$FileCount = $files.count #get count of files (For progress bar)
			Foreach($file in $files)
			{
				#for each of the files
				$Filename = $file.Fullname.tolower().replace($Source,'') #remove the source path from the file name
				$DestinationFile = ($Destination+$Filename) #replace it with the destination path
				Update-ProgressBar -Status ('Restoring items from {0} Step {1} of {2}' -f $Row.Name, $i, $count) -ProgressValue ($i/$count*100) -SubStatus ('Copying {0}' -f $file.FullName) -SubProgressValue ($filenumber/$FileCount *100) #update progress bar
				Copy-File -Source $File.Fullname -Destination $DestinationFile
				$filenumber++  #increment counter
			}
		}
		ElseIF($Row.Type -eq 'Folder') #else if a normal folder
		{
			$Source = $Row.FullPath.ToLower()
			$Destination = ($fileList|Where-Object -FilterScript {
					$_.Destination -eq $Row.FullPath
			}).OriginalLoc -replace '\\\\(?:(SOE|Q\D*[012]*|PC|EQ)\D*\d{5,5}|localhost)\\C\$', 'C:' #set destination to the original folder location using the log of the backup.
			IF(-not (Test-Path -Path $Destination))
			{
				New-Item -Path $Destination -Force -ItemType Directory #create destination directory
			}
			$files = Get-ChildItem -Path $Source -Recurse #get child items from backup folder
			$filenumber = 0 #counter 
			$FileCount = $files.count #number of files to process (for progress bar)
			Foreach($file in $files)  #for each of the files.
			{
				$DestinationFile = ($fileList|Where-Object -FilterScript {
						$_.Destination -eq $file.fullName
				}).OriginalLoc
				Update-ProgressBar -Status ('Restoring items from {0} Step {1} of {2}' -f $Row.Name, $i, $count) -ProgressValue ($i/$count*100) -SubStatus ('Copying {0}' -f $file.FullName) -SubProgressValue ($filenumber/$FileCount *100) #update progress bar
				Copy-File -Source $File.Fullname -Destination $DestinationFile
				$filenumber++  #increment counter.
			}
		}
		ElseIF($Row.Type -eq 'File')
		{
			#if copying a file.
			Update-ProgressBar -Status ('Restoring {0} Step {1} of {2}' -f $Row.Name, $i, $count) -ProgressValue ($i/$count*100) -SubStatus ' ' -SubProgressValue 0 #update progress bar
			$Destination = ($fileList|Where-Object -FilterScript {
					$_.Destination -eq $Row.FullPath
			}).OriginalLoc -replace '\\\\(?:(SOE|Q\D*[012]*|PC|EQ)\D*\d{5,5}|localhost)\\C\$', 'C:' #get the original location of the file
			Copy-File -Source $row.FullPath -Destination $Destination
		}
	}
	IF($chk_networkDrives.IsChecked) 
	{
		#if restoring network drive mappings
		$i++ #Increment counter
		Update-ProgressBar -Status ('Restoring Network Drive Mappings Step {0} of {1}' -f $i, $count) -ProgressValue ($i/$count*100) -SubStatus ' ' -SubProgressValue 0 #update progress bar
		$drives = Import-Csv -Path $SaveLoc\Drives.csv #get list of network drives
		$drivecount = $drives.count #get count of drives to map
		$j = 0 #counter
		Foreach($drive in $drives)
		{
			#for each network drive
			$j++ #increment counter
			$letter = $drive.name.substring(0,1) #get the drive letter
			$Path = $drive.Providername #get the path
			Update-ProgressBar -Status ('Restoring Network Drive Mappings Step {0} of {1}' -f $i, $count) -ProgressValue ($i/$count*100) -SubStatus ('Mapping {0} to {1}' -f $letter, $Path) -SubProgressValue ($j/$drivecount*100) #update progress bar
			New-PSDrive -Persist -Name $letter -PSProvider FileSystem -Root $Path -Scope Global #map network drive. Jared Changed order of operations and edited Global scope call
		}
	}
	IF($chk_Printers.ischecked)
	{
		#if restoring printer mappings
		$i++ #increment counter
		Update-ProgressBar -Status ('Restoring Printer Mappings Step {0} of {1}' -f $i, $count) -ProgressValue ($i/$count*100) -SubStatus ' ' -SubProgressValue 0 #update progress bar
		$printers = Get-Content -Path $SaveLoc\Printers.txt #get printers.
		$ws_net = New-Object -ComObject wscript.network #create com object to map printers
		$printerCount = $printers.count #count the number of printers to map
		$j = 0 #counter
		Foreach($printer in $printers)
		{
			#for each printer
			$j++ #increment counter
			Update-ProgressBar -Status ('Restoring Printer Mappings Step {0} of {1}' -f $i, $count) -ProgressValue ($i/$count*100) -SubStatus ('Mapping {0}' -f $printer) -SubProgressValue ($j/$printerCount*100) #update progress bar
			$ws_net.addwindowsPrinterConnection($printer) #map printer
		}
	}
	$script:Hash_ProgressBar.Form.Dispatcher.Invoke([Action][ScriptBlock]::Create({
				$script:Hash_ProgressBar.Form.Close()
	})) #close the progress bar.
}
Function Start-RestoreAdmin
{
	Add-Type -AssemblyName PresentationFramework
	$SaveLoc = $TxtSaveLoc.Text #set the saveloc varaible to the selected backup location
	$fileList = Import-Csv -Path $SaveLoc\FileList_Backup.csv
	$Script:RemoteInfo = New-Object -TypeName RemoteInfo
	$Script:RemoteInfo.Computername = Get-ComputerNameGUI
	
	IF(-not (Test-Connection -ComputerName $Script:RemoteInfo.Computername -Count 1 -Quiet) -or -not (Test-Path -Path ('\\{0}\C$' -f $Script:RemoteInfo.Computername))) #check that remote machine is available.
	{
		[Windows.Messagebox]::Show(('Cannot connect to {0}' -f $Script:RemoteInfo.Computername))
		Return
	}

	IF($Script:RemoteInfo.Computername -match 'localhost') #if the computername is localhost the cim session needs to be created slightly differently than if it's a netbios name.
	{
		$CimSession = New-CimSession -ComputerName $Script:RemoteInfo.Computername
	}
	Else
	{
		$CimSession = New-CimSession -ComputerName ([net.dns]::GetHostByName($Script:RemoteInfo.Computername).Hostname)
	}

	$OS = Get-CimInstance -Query 'SELECT Caption FROM Win32_OperatingSystem' -CimSession $CimSession|Select-Object -ExpandProperty Caption #get OS type
	$profile = Get-SSID -Session $CimSession | Show-DataGrid -WindowTitle 'Select User' #select profile to restore data to.
	IF(-not $profile){return} #if no profile stop processing.
	
	#define proile info and admin share.
	$Script:RemoteInfo.ProfileFolder = '\\{0}\C$\Users\{1}' -f $Script:RemoteInfo.Computername,$profile.ProfileFolder 
	$Script:RemoteInfo.IsInUse = $profile.IsInUse
	$Script:RemoteInfo.ProfileName = $profile.ProfileName
	$Script:RemoteInfo.SSID = $profile.SSID
	$Adminshare = '\\{0}\C$' -f $Script:RemoteInfo.Computername
	
	$names = @('Desktop','Documents','Favorites') #user shell folder names
	IF($Script:RemoteInfo.IsInUse)
	{ #if profile is loaded.
		Try
		{ 
			$Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Users',$Script:RemoteInfo.Computername) #open HKey_Users on the remote machine.
			$Script:RemoteInfo.UserHive = $Registry.OpenSubKey($Script:RemoteInfo.SSID,$True)
		}
		Catch
		{
			Write-Error -Message $_
			[Windows.Messagebox]::Show('Failed to open remote Registry hive')
			Return
		}
	}
	Else
	{#if profile is not loaded, load the user hive into regedit on the local machine.
		$keyName = 'HKU\{0}_Remote' -f $Script:RemoteInfo.SSID
		$keyLocation = '{0}\ntuser.dat' -f $Script:RemoteInfo.ProfileFolder
		Try
		{
			Reg.exe Load $keyName $keyLocation
			New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_Users -ErrorAction Stop
			$Script:RemoteInfo.UserHive = get-item -Path ('HKU:\{0}_Remote' -f $Script:RemoteInfo.SSID)
		}
		Catch
		{
			Write-Error -Message $_
			[Windows.Messagebox]::Show('Failed to open remote Registry hive')
			Return
		}
	}
	$regkey = $Script:RemoteInfo.UserHive.OpenSubKey('Software\\microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders',$True) #open the user shell folders regkey.
	$Specialfolders = @() #create array
	ForEach($name in $names) 
	{
		$Specialfolders += [PsCustomObject]@{
			Name = $name
			FullPath = $regkey.GetValue($name) -replace ('C:\\Users\\{0}' -f $Script:RemoteInfo.ProfileName), $Script:RemoteInfo.ProfileFolder #get folder path and replace enviroment variable reference.
		}
	}
	$count = $lvwFileList.SelectedItems.Count
	$i = 0 #counter
	$RegexSafeUsers = '{0}\\Users\\|C:\\Users\\' -f (Format-RegexSafe -String $Adminshare) #create regex expression for profile path.
	Show-ProgressBar -Status 'Restoring Files' -ProgressValue 0 -SubStatus ' ' -SubProgressValue 0 #update progress bar
	Foreach($Row in $lvwFileList.SelectedItems) #for each of the selected items.
	{
		$i++ #increment counter
		IF($SpecialFolders.Name -match ($Row.Name -replace 'Redirected') -and $Row.Type -eq 'Folder') #if the name of the item matches one of the special folders.
		{
			Update-ProgressBar -Status ('Restoring items from {0} Step {1} of {2}' -f $Row.Name, $i, $count) -ProgressValue 0 -SubStatus ' ' -SubProgressValue 0 #update progress bar
			$Destination = ($SpecialFolders|Where-Object -FilterScript {
					#destination equals the special folder path.
					$Row.Name -match $_.Name
			}).FullPath
			IF($Destination -match ('({0}).*?\\' -f $regexsafeUsers))
			{
				$Destination = $Destination -replace ('({0}).*?\\' -f $regexsafeUsers),('{0}\' -f $Script:RemoteInfo.ProfileFolder)
			}
			$Source = $Row.FullPath.ToLower() #source equals backup location to lower case
			$files = Get-ChildItem -Path $Source -Recurse #get all child files and folders.
			$filenumber = 0 #counter
			$FileCount = $files.count #get count of files (For progress bar)
			Foreach($file in $files)
			{
				#for each of the files
				$Filename = $file.Fullname.tolower().replace($Source,'') #remove the source path from the file name
				$DestinationFile = ($Destination+$Filename) #replace it with the destination path
				Update-ProgressBar -Status ('Restoring items from {0} Step {1} of {2}' -f $Row.Name, $i, $count) -ProgressValue ($i/$count*100) -SubStatus ('Copying {0}' -f $file.FullName) -SubProgressValue ($filenumber/$FileCount *100) #update progress bar
				Copy-File -Source $File.Fullname -Destination $DestinationFile
				$filenumber++  #increment counter
			}
		}
		
		ElseIF($Row.Type -eq 'Folder') #else if a normal folder
		{
			$Source = $Row.FullPath.ToLower()
			$Destination = ($fileList|Where-Object -FilterScript {
					$_.Destination -eq $Row.FullPath
			}).OriginalLoc -replace '\\\\(?:(SOE|Q\D*[012]*|PC|EQ)\D*\d{5,5}|localhost)\\C\$|C:',$Adminshare #replace the \\pcnum\C$ or C: in the file path with the admin share for the selected PC.
			IF($Destination -match ('({0}).*?\\' -f $regexsafeUsers))
			{
				$Destination = $Destination -replace ('{0}.*?\\' -f $regexsafeUsers),('{0}\' -f $Script:RemoteInfo.ProfileFolder) #replace the user folder path with the current profile folder path.
			}
			New-Item -Path $Destination -Force -ItemType Directory #create destination directory
			$files = Get-ChildItem -Path $Source -Recurse #get child items from backup folder
			$filenumber = 0 #counter 
			$FileCount = $files.count #number of files to process (for progress bar)
			Foreach($file in $files)  #for each of the files.
			{
				$DestinationFile = ($fileList|Where-Object -FilterScript {
						$_.Destination -eq $file.fullName
				}).OriginalLoc -replace '\\\\(?:(SOE|Q\D*[012]*|PC|EQ)\D*\d{5,5}|localhost)\\C\$|C:',$Adminshare #replace the \\pcnum\C$ or C: in the file path with the admin share for the selected PC.
				IF($DestinationFile -match ('({0}).*?\\' -f $regexsafeUsers))
				{
					$DestinationFile = $DestinationFile -replace ('({0}).*?\\' -f $regexsafeUsers),('{0}\' -f $Script:RemoteInfo.ProfileFolder) #replace the user folder path with the current profile folder path.
				}
				Update-ProgressBar -Status ('Restoring items from {0} Step {1} of {2}' -f $Row.Name, $i, $count) -ProgressValue ($i/$count*100) -SubStatus ('Copying {0}' -f $file.FullName) -SubProgressValue ($filenumber/$FileCount *100) #update progress bar
				Copy-File -Source $File.Fullname -Destination $DestinationFile
				$filenumber++  #increment counter.
			}
		}
		ElseIF($Row.Type -eq 'File')
		{
			#if copying a file.
			Update-ProgressBar -Status ('Restoring {0} Step {1} of {2}' -f $Row.Name, $i, $count) -ProgressValue ($i/$count*100) -SubStatus ' ' -SubProgressValue 0 #update progress bar
			$Destination = ($fileList|Where-Object -FilterScript {
					$_.Destination -eq $Row.FullPath
			}).OriginalLoc -replace '\\\\(?:(SOE|Q\D*[012]*|PC|EQ)\D*\d{5,5}|localhost)\\C\$|C:',$Adminshare #replace the \\pcnum\C$ or C: in the file path with the admin share for the selected PC.
			IF($Destination -match ('({0}).*?\\' -f $regexsafeUsers))
			{
				$Destination = $Destination -replace ('({0}).*?\\' -f $regexsafeUsers),('{0}\' -f $Script:RemoteInfo.ProfileFolder) #replace the user folder path with the current profile folder path.
			}
			Copy-File -Source $row.FullPath -Destination $Destination
		}
	}
	IF(-not $Script:RemoteInfo.IsInUse)
	{#if the profile isn't loaded.
		Remove-PSDrive -Name HKU -Force #remove the PS Drive to regedit.
		Start-Sleep -Seconds 1
		reg.exe Unload ('HKU\{0}_Remote' -f $Script:RemoteInfo.SSID) #unload the User's hive.
	}
	$script:Hash_ProgressBar.Form.Dispatcher.Invoke([Action][ScriptBlock]::Create({
				$script:Hash_ProgressBar.Form.Close()
	})) #close the progress bar.
}
Function Start-BRProcess
{
	#function to call the process function for the selected option
	IF([string]::IsNullOrEmpty($TxtSaveLoc.text))
	{
		#check that a save location has been selected.
		$Message = ("An error occurred while trying to start the {0}`r`n`r`nPlease select a save location." -f $Script:Action)
		Show-UserPrompt -Message $Message -Button3 'Okay' #show error to client
		return #stop processing
	}
	IF([math]::Round(($script:BackupSize/1kb),3) -ge $Script:DriveSpace)
	{
		$Message = ("An error occurred while trying to start the {0}`r`n`r`nThere is not enough space on the destination drive." -f $Script:Action)
		Show-UserPrompt -Message $Message -Button3 'Okay' #show error to client
		return #stop processing
	}
	IF($Script:Action -eq 'Backup') #if we're running a backup
	{
		Start-Backup #call the backup function
	}
	ElseIF($Script:Action -eq 'Restore') #if we're running a restore
	{
		IF($script:Admin)
		{
			Start-RestoreAdmin
		}
		Else
		{
			Start-Restore #call the restore function
		}
	}
	Else
	{
		#shouldn't be possible to hit this bit but just in case
		$Message = ("An error occurred while trying to initiate the {0}`r`nThe application was unable determine which action was being taken.`r`nPlease restart the app and try again." -f $Script:Action)
		Show-UserPrompt -Message $Message -Button3 'Okay' #show message to the client
		return #return to front screen
	}
}
Function Initialize-Form
{
	#on content rendered (last event to occur when showing the form) occurs after all controls on the form have been rendered and displayed.
	[string]$Script:Action2 = Show-UserPrompt -Message 'Run A GPUPDATE?' -Button2 'Yes' -Button3 'No' #prompt user asking if they want to run GPUpdate
	[bool]$script:Admin = ([Security.Principal.WindowsPrincipal]([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    [string]$Script:Action = Show-UserPrompt -Message 'Are we running a Backup or Restore?' -Button2 'Restore' -Button3 'Backup' #prompt user on if we're running a backup or a restore
	[bool]$script:Admin = ([Security.Principal.WindowsPrincipal]([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	IF($admin)
	{
		IF($Script:Action2 -eq 'Yes') 
		{
			write-host "Inside Initialize-FormAdmin" -ForegroundColor Cyan
			#if they select Yes  
			Set-GPupdate #call the set gpupdate function
		}  
		If($Script:Action -eq 'Backup') 
		{
			#if they select backup
			Set-InitialStateAdminBackup #call the set initial stat backup function
		}
		ElseIf($Script:Action -eq 'Restore')
		{
			#else if they select restore
			Set-InitialStateRestore #set the initial state of the form for restores.
		}
	}
	Else{
		IF($Script:Action2 -eq 'Yes') 
		{
			write-host "Inside Initialize-Form" -ForegroundColor Cyan
			#if they select Yes  
			Set-GPupdate #call the set gpupdate function
		}  
		If($Script:Action -eq 'Backup') 
		{
			#if they select backup
			write-host "Inside Initialize-Form, Backup Pressed" -ForegroundColor Cyan
			Set-InitialStateBackup #call the set initial stat backup function
		}
		ElseIf($Script:Action -eq 'Restore')
		{
			#else if they select restore
			Set-InitialStateRestore #set the initial state of the form for restores.
		}
	}
}
#endregion Functions
#XAML For main screen.
[xml]$XAML = 
@'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    ResizeMode="CanResize"
    Width="800"
    Height="600"
    Title="Data Backup Tool"
    WindowStartupLocation="CenterScreen">
    <Grid Margin="0,0,0,0">
        <Label Width="Auto" Height="30" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Content="Backup Location (New or Existing):"/>
        <TextBox Name="TxtSaveLoc" Width="400" Height="30" HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top" IsReadOnly="True"/>
        <Button Name="btnSavBrowse" Width="60" Height="30" HorizontalAlignment="Left" Margin="420,40,0,0" VerticalAlignment="Top" Content="Browse"/>
        <Label Name="lblFreeSpace" Width="Auto" Height="30" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,10,0,0"/>
        <Label Name="lblRequiredSpace" Width="Auto" Height="30" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,40,0,0"/>
        <Label Width="200" Height="30" HorizontalAlignment="Left" Margin="10,80,0,0" VerticalAlignment="Top" Content="Files/Locations Selected for Backup:"/>
        <ListView Name="lvwFileList" Height="Auto" Width="Auto" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="10,110,200,50">
          <ListView.Resources>
            <Style TargetType="{x:Type ListViewItem}">
              <Setter Property="IsSelected" Value="{Binding Selected, Mode=TwoWay}"/>
            </Style>
          </ListView.Resources>
          <ListView.View>
            <GridView>
              <GridView.Columns>
                <GridViewColumn>
                  <GridViewColumn.CellTemplate>
                    <DataTemplate>
                      <CheckBox Tag="{Binding Name}" IsChecked="{Binding RelativeSource={RelativeSource AncestorType={x:Type ListViewItem}}, Path=IsSelected}"/>
                    </DataTemplate>
                  </GridViewColumn.CellTemplate>
                </GridViewColumn>
                <GridViewColumn DisplayMemberBinding="{Binding Name}" Header="Name"/>
                <GridViewColumn DisplayMemberBinding="{Binding FullPath}" Header="FullPath"/>
                <GridViewColumn DisplayMemberBinding="{Binding Type}" Header="Type"/>
                <GridViewColumn DisplayMemberBinding="{Binding Size}" Header="Size mb"/>
              </GridView.Columns>
            </GridView>
          </ListView.View>
        </ListView>
        <Button Name="btnAddLoc" Width="80" Height="30" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,110,100,0" Content="Add File"/>
        <StackPanel Margin="0,180,10,0" HorizontalAlignment="Right">
            <Label Content="Non File Based Options"/>
            <CheckBox Name="chk_networkDrives" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,2,0,0" IsChecked="True">Network Drive Mappings</CheckBox>
            <CheckBox Name="chk_Printers" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,2,0,0" IsChecked="True">Printer Mappings</CheckBox>
        </StackPanel>
        <Button Name="btnStart" Width="70" Height="30" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="10,420,0,10" Content="Start"/>
        <Button Name="btnReset" Width="70" Height="30" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="85,420,0,10" Content="Reset Form"/>
    </Grid>
</Window>
'@
$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $XAML)  #convert xaml 
try
{
	$Form = [Windows.Markup.XamlReader]::Load($reader) #load the form.
}
catch
{
	Write-Verbose -Message 'Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered.'
}

$TxtSaveLoc = $Form.FindName('TxtSaveLoc')
$lblFreeSpace = $Form.FindName('lblFreeSpace')
$lblRequiredSpace = $Form.FindName('lblRequiredSpace')
$lvwFileList = $Form.FindName('lvwFileList')
$btnSavBrowse = $Form.FindName('btnSavBrowse')
$btnAddLoc = $Form.FindName('btnAddLoc')
$btnStart = $Form.FindName('btnStart')
$btnReset = $Form.FindName('btnReset')
$chk_networkDrives = $Form.FindName('chk_networkDrives')
$chk_Printers = $Form.FindName('chk_Printers')

$btnSavBrowse.Add_click({
		Set-SaveLocation #on click of save location button call the set-savelocation function
})
$btnAddLoc.Add_Click({
		Add-BackupLocation #on click of the add File/Folder button call the add-backuplocation function
})
$btnStart.Add_Click({
		Start-BRProcess #on click of the start button call the start process function
})
$btnReset.Add_Click({
		$lvwFileList.ItemsSource = $null
		$TxtSaveLoc.Text = $null
		Initialize-Form
})
$lvwFileList.Add_SelectionChanged(
	{
		$script:backupSize = ($script:FilePaths|Where-Object{$_.Selected -eq $true}|Measure-Object -Property Size -Sum).Sum
		$bsize = [math]::Round($($script:backupSize/1kb),3) #round out new backup size to 3 places
		$lblrequiredSpace.Content = ("Required Space for Backup:`t{0} Gb" -f $bsize) #update label
	}
)
$Form.Add_ContentRendered({
		Initialize-Form
})
$null = $Form.showdialog() #show the form.