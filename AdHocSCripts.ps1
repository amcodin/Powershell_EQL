Powershell –

Adding NetworkDrives

================================

                #if restoring network drive mappings

                                #$i++ #Increment counter

                                #Update-ProgressBar -Status ('Restoring Network Drive Mappings Step {0} of {1}' -f $i, $count) -ProgressValue ($i/$count*100) -SubStatus ' ' -SubProgressValue 0 #update progress bar

                                $drives = Import-Csv -Path "C:\LocalData\Drives.csv" #get list of network drives

                                $drivecount = $drives.count #get count of drives to map

                                $j = 0 #counter

 

                                Foreach($drive in $drives)

                                {

                                                #for each network drive

                                                $j++ #increment counter

                                                $letter = $drive.name.substring(0,1) #get the drive letter

                                                $Path = $drive.Providername #get the path

                                New-PSDrive -Persist -Name $letter -PSProvider FileSystem -Root $Path -Scope Global #map network drive.

                                }

 

 

BulkGetSerialNumber

 

<# Written by Jared Vosters 20.10.23

Script bulk looks up serial numbers and spits out to file

#>

$Computers = Get-Content "C:\TEMP\GetSerialPC.txt"    ## Location of File to be imported

$MonintorInfo = "C:\TEMP\Successfully_GetSerialNum.csv"        ## Location of Output file for machines that were online, script successfully ran

$OfflineMachines = "C:\TEMP\Failed_GetSerialNum.csv"          ## Location of Output file for machines that were Offline, script failed to run

 

 

#Write-Output "PC Name,Issue" | Set-Content -Path $OfflineMachines

#$result = @()

 

ForEach($computer in $Computers){

    if (Test-WSMan -ComputerName $Computer -ErrorAction 'SilentlyContinue'){

        try {

            $CimSession = New-CimSession -ComputerName $Computer

 

            #$PcSerial = (Get-WmiObject -Class:Win32_ComputerSystem).Model

            #$PcSerial = Get-WmiObject win32_bios | select Serialnumber

                  

            $PcSerial = Get-WmiObject -ComputerName $computer -Class Win32_BIOS | select Serialnumber

           

            #Get-WmiObject -ComputerName $computer -Class Win32_BIOS -Filter ‘SerialNumber=“XXXXXX”’ | Select -Property PSComputerNam

            #Get-CimInstance

 

 

 

 

 

            Write-Output "$computer,$PcSerial" | Add-Content -Path $MonintorInfo

        }

        catch{

            Write-Error $PSItem.Exception

            if ($computer -ge 1){  ##was computernamechanged

               Continue

            }

            else {

                Write-Output "$computer,WMI ISSUE" | Add-Content -Path $OfflineMachines

                Continue

                #break

            }

        }

    }

    Else{

    Write-Output "$computer,Offline" | Add-Content -Path $OfflineMachines

    }

}

 

 

RemoteTriggerConfigMActions

 

<# Written by Jared Vosters & David Onions 19.10.23

# Triggers taken from https://learn.microsoft.com/en-us/mem/configmgr/develop/reference/core/clients/client-classes/triggerschedule-method-in-class-sms_client

 

Script will remotely trigger/start configuration manager actions, based on the trigger values, see link above.

 

-ForceSCCMActions.txt Modify the list of PC's targeted, $Computers. PC names MUST have the full domain name, e.g. SOE05327.myergon.local , PC72088.Energex.com.au, Q233216.mysparq.local, EQ71474.energyq.com.au

--txt file format is PCnameFQDN and space

 

-Successfully_Triggered_Devices is a list of devices after the script is ran, that were successfully ran

 

-Failed_Triggered_Devices is a list of devices the script couldnt find (offline, network issue, missing domain, etc) just rerun

 

#>

$Computers = Get-Content "C:\TEMP\ForceSCCMActions.txt"    ## Location of File to be imported

$MonintorInfo = "C:\TEMP\Successfully_Triggered_Devices.csv"        ## Location of Output file for machines that were online, script successfully ran

$OfflineMachines = "C:\TEMP\Failed_Triggered_Devices.csv"          ## Location of Output file for machines that were Offline, script failed to run

 

#Add or remove policies as you see fit, this is the 'optimal order' to kick off windows updates 1st

$Polices = {

    $SCCMActions = @(

        " {00000000-0000-0000-0000-000000000021}", #Machine Policy Assignments Request

        " {00000000-0000-0000-0000-000000000022}", #Machine Policy Evaluation

        " {00000000-0000-0000-0000-000000000101}", #Hardware Inventory Collection Cycle

        " {00000000-0000-0000-0000-000000000001}", #Hardware Inventory

        " {00000000-0000-0000-0000-000000000002}", #Software Inventory

        " {00000000-0000-0000-0000-000000000102}", #Software Inventory Collection Cycle

        " {00000000-0000-0000-0000-000000000108}"  #Software Updates Assignments Evaluation Cycle

        )

    }

 

#Write-Output "PC Name,Issue" | Set-Content -Path $OfflineMachines

#$result = @()

 

ForEach($computer in $Computers){

    if (Test-WSMan -ComputerName $Computer -ErrorAction 'SilentlyContinue'){

        try {

            $CimSession = New-CimSession -ComputerName $Computer

 

            foreach ($action in $SCCMActions) {

            Invoke-WMIMethod -Namespace rootccm -Class SMS_CLIENT -Name TriggerSchedule $action

            }

 

            Invoke-Command -ScriptBlock $Polices -ComputerName $computer 

                   

            Write-Output "$computer" | Add-Content -Path $MonintorInfo

        }

        catch{

            Write-Error $PSItem.Exception

            if ($computer -ge 1){  ##was computernamechanged

               Continue

            }

            else {

                Write-Output "$computer,WMI ISSUE" | Add-Content -Path $OfflineMachines

                Continue

                #break

            }

        }

    }

    Else{

    Write-Output "$computer,Offline" | Add-Content -Path $OfflineMachines

    }

}

MonitorExport

 

## Set File Variables Here

$Computers = Get-Content "C:\LocalData\MonitorExport\Computers.csv"           ## Location of CSV File to be imported

$MonintorInfo = "C:\LocalData\MonitorExport\Monitor_Info.csv"                 ## Location of Output file for machines that were online

$OfflineMachines = "C:\LocalData\MonitorExport\Offline_Machines.csv"          ## Location of Output file for machines that were Offline

 

Function ConvertTo-Char {

       

                                param

                                (

                                                $Array

                                )

                                $Output = ""

                                ForEach($char in $Array){$Output += [char]$char -join ""}

                                return $Output

                                }

Function Convert-MonitorManufacturer

{

<#

        .SYNOPSIS

        This should only be used by Get-RSMonitorInformation

        .DESCRIPTION

        Will translate the 3 letter code to the full name of the manufacturer, this should only be used by Get-RSMonitorInformation.

        .PARAMETER Manufacturer

        Enter the 3 letter manufacturer code.

        .EXAMPLE

        Convert-MonitorManufacturer -Manufacturer "PHL"

        # Return the translation of the 3 letter code to the full name of the manufacturer, in this example it will return Philips

        .LINK

https://github.com/rstolpe/MonitorInformation/blob/main/README.md

        .NOTES

        Author:         Robin Stolpe

        Mail:           robin@stolpe.io

        Twitter:        https://twitter.com/rstolpes

        Linkedin:       https://www.linkedin.com/in/rstolpe/

        Website/Blog:   https://stolpe.io

        GitHub:         https://github.com/rstolpe

        PSGallery:      https://www.powershellgallery.com/profiles/rstolpe

    #>

 


 

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, HelpMessage = "Enter the 3 letter manufacturer code")]

        [String]$Manufacturer

    )

 


 

    Switch ($Manufacturer)

    {

        ACI    {

            return "Asus"

        }

        ACR    {

            return "Acer"

        }

        ACT    {

            return "Targa"

        }

        ADI    {

            return "ADI Corporation"

        }

        AMW    {

            return "AMW"

        }

        AOC    {

            return "AOC"

        }

        API    {

            return "Acer"

        }

        APP    {

            return "Apple"

        }

        ART    {

            return "ArtMedia"

        }

        AST    {

            return "AST Research"

        }

        AUO    {

            return "AU Optronics"

        }

        BMM    {

            return "BMM"

        }

        BNQ    {

            return "BenQ"

        }

        BOE    {

            return "BOE Display Technology"

        }

        CPL    {

            return "Compal"

        }

        CPQ    {

            return "COMPAQ"

        }

        CTX    {

            return "Chuntex"

        }

        DEC    {

            return "Digital Equipment Corporation"

        }

        DEL    {

            return "Dell"

        }

        DPC    {

            return "Delta"

        }

        DWE    {

            return "Daewoo"

        }

        ECS    {

            return "ELITEGROUP"

        }

        EIZ    {

            return "EIZO"

        }

        EPI    {

            return "Envision"

        }

        FCM    {

            return "Funai"

        }

        FUS    {

            return "Fujitsu Siemens"

        }

        GSM    {

            return "LG (GoldStar)"

        }

        GWY    {

            return "Gateway"

        }

        HEI    {

            return "Hyundai Electronics"

        }

        HIQ    {

            return "Hyundai ImageQuest"

        }

        HIT    {

            return "Hitachi"

        }

        HSD    {

            return "Hannspree"

        }

        HSL    {

            return "Hansol"

        }

        HTC    {

            return "Hitachi / Nissei Sangyo"

        }

        HWP {

            return "Hewlett Packard (HP)"

        }

        HPN {

            return "Hewlett Packard (HP)"

        }

        IBM    {

            return "IBM"

        }

        ICL    {

            return "Fujitsu"

        }

        IFS    {

            return "InFocus"

        }

        IQT    {

            return "Hyundai"

        }

        IVM    {

            return "Idek Iiyama"

        }

        KDS    {

            return "KDS"

        }

        KFC    {

            return "KFC Computek"

        }

        LEN    {

            return "Lenovo"

        }

        LGD    {

            return "LG"

        }

        LKM    {

            return "ADLAS / AZALEA"

        }

        LNK    {

            return "LINK"

        }

        LPL    {

            return "LG Philips"

        }

        LTN    {

            return "Lite-On"

        }

        MAG    {

            return "MAG InnoVision"

        }

        MAX    {

            return "Maxdata"

        }

        MEI    {

            return "Panasonic"

        }

        MEL    {

            return "Mitsubishi"

        }

        MIR    {

            return "miro"

        }

        MTC    {

            return "MITAC"

        }

        NAN    {

            return "NANAO"

        }

        NEC    {

            return "NEC"

        }

        NOK    {

            return "Nokia"

        }

        NVD {

            return "Nvidia"

        }

        OQI    {

            return "OPTIQUEST"

        }

        PBN    {

            return "Packard Bell"

        }

        PCK    {

            return "Daewoo"

        }

        PDC    {

            return "Polaroid"

        }

        PGS    {

            return "Princeton Graphic Systems"

        }

        PHL    {

            return "Philips"

        }

        PRT    {

            return "Princeton"

        }

        REL    {

            return "Relisys"

        }

        SAM    {

            return "Samsung"

        }

        SEC    {

            return "Seiko Epson"

        }

        SMC    {

            return "Samtron"

        }

        SMI    {

            return "Smile"

        }

        SNI {

            return "Siemens"

        }

        SNY    {

            return "Sony"

        }

        SPT    {

            return "Sceptre"

        }

        SRC    {

            return "Shamrock"

        }

        STN    {

            return "Samtron"

        }

        STP    {

            return "Sceptre"

        }

        TAT {

            return "Tatung"

        }

        TRL    {

            return "Royal"

        }

        TSB    {

            return "Toshiba"

        }

        UNM    {

            return "Unisys"

        }

        VSC    {

            return "ViewSonic"

        }

        WTC    {

            return "Wen"

        }

        ZCM    {

            return "Zenith"

        }

        default {

            return $Manufacturer

        }

    }

}

 

 

Write-Output "PC Name,Issue" | Set-Content -Path $OfflineMachines

$result = @()

ForEach($computer in $Computers){

        if (Test-WSMan -ComputerName $Computer -ErrorAction 'SilentlyContinue'){

            try {

                $CimSession = New-CimSession -ComputerName $Computer

                $PcModel = (Get-WmiObject -Class:Win32_ComputerSystem).Model

                $VPN = $null

                $VPN =  Get-NetIPConfiguration -CimSession $CimSession | where InterfaceAlias -like "*vpn*"

                If ($VPN -eq $null){

                    $Monitors = Get-CimInstance -Query "Select * FROM WMIMonitorID" -Namespace root\wmi -cimsession $CimSession

                    $outputObj = New-Object PSObject -Property @{

                        Computer = $Computer

                    }

                    $outputObj|Add-Member -MemberType NoteProperty -Name "PC Model" -Value $PcModel  ##this line adds the PC Model

                    $i=0

                    ForEach ($Monitor in $Monitors) {

                    $i++

                                    $object = New-Object PSObject -Property @{

                                                Active = $Monitor.Active

                                                Manufacturer = Convert-MonitorManufacturer (ConvertTo-Char($Monitor.ManufacturerName)).Trim()

                                                Model = (ConvertTo-Char($Monitor.userfriendlyname)).Trim()

                                                    SerialNumber = (ConvertTo-Char($Monitor.serialnumberid)).Trim()

                                                WeekOfManufacture = $Monitor.WeekOfManufacture

                                                YearOfManufacture = $Monitor.YearOfManufacture

                                    }

 

                    If ($object.model -eq ""){

                    $Object.Model = "Laptop Screen"

                    }

                   

                    $outputObj|Add-Member -MemberType NoteProperty -Name "Model$i" -Value $object.Model #this line adds the name:value pair to the output.

                    }

                    

                    $outputObj|Add-Member -MemberType NoteProperty -Name MonitorCount -Value $i #this line adds the name:value pair to the output.

                    write-host $outputObj

                    $result += $outputObj

                    }

                    Else {

                    Write-Output "$computer,Connected to VPN" | Add-Content -Path $OfflineMachines

                    }

            }

            catch{

                Write-Error $PSItem.Exception

                if ($ComputerName -ge 1){

                    Continue

                }

                else {

                    Write-Output "$computer,WMI ISSUE" | Add-Content -Path $OfflineMachines

                    Continue

                    #break

                }

            }

        }

        Else{

            Write-Output "$computer,Offline" | Add-Content -Path $OfflineMachines

        }

    }

 

$result|Sort-Object -Descending -Property MonitorCount | Export-Csv -Path $MonintorInfo -NoTypeInformation