import os 
import sys 
install_office = itsm.getParameter('Install_Office')
install_editions = ",".join(itsm.getParameter('Install_Office_Edition'))
Install_Office_Channel = ",".join(itsm.getParameter('Install_Office_Channel'))
install_download_path = itsm.getParameter('Install_Office_Download_Path')
install_with_odt = itsm.getParameter('Install_Office_WITH_XML')
install_with_odt_path = itsm.getParameter('Install_Office_WITH_XML_PATH')
remove_office_only = itsm.getParameter('Remove_Office365_Only')
remove_office_only_edition = ",".join(itsm.getParameter('Remove_Office_Only_Edition'))
excluded_apps = ",".join(itsm.getParameter('Install_Office_Exclude_Apps'))
ps_content=r'''
<#
    .Name
    EZT-DeployO365

    .Version 
    0.5.2

    .SYNOPSIS
    Automates silent install or removal of Office 365 editions using the Office Deployment Tookit, allowing custom config XML generation and removal of existing/older Office installs

    .DESCRIPTION
       
    .Configurable Variables

    .Requirements
    - Powershell v3.0 or higher

    .EXAMPLE
    .\EZT-DeployO365.ps1

    .OUTPUTS
    System.Management.Automation.PSObject

    .Credits
     Install-Office365Suite        - https://github.com/mallockey/Install-Office365Suite
     Get-OfficeVersion             - https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls
     Write-Color                   - https://github.com/EvotecIT/PSWriteColor

    .NOTES
    Author: EZTechhelp
    Site  : https://www.eztechhelp.com
#> 

#############################################################################
#region Configurable Script Parameters
#############################################################################

#---------------------------------------------- 
#region Function Select Variables
#----------------------------------------------
$Install_Office = ''' + install_office + ''' #installs Office 365 using config values set in this script
$Install_Office_Edition = "''' + install_editions + '''" #Office edition to be installed
$Install_Office_Edition = $Install_Office_Edition.split(',') | Select -First 1
$Install_Office_Download_Path = "''' + install_download_path + '''" #directory where the Office 365 setup and ODT files should be downloaded
$Remove_Install_Office_Download_Path = 0 #enable to remove the folder and all downloaded files from the path defined in Install_Office_Download_Path after execution completes 

$Install_Office_WITH_ODT = ''' + install_with_odt + ''' #installs Office using an existing Office 365 ODT configuration XML file
$Install_Office_WITH_ODT_XMLFILE = "''' + install_with_odt_path + '''" #Full path to an existing Office365 ODT Configuration XML file to use with deployment.To config or create a configuration file, recommend visiting https://config.office.com

$Remove_Office_Only = ''' + remove_office_only + ''' #removes any detected office installations without installing another
$Remove_Office_Only_Edition = "''' + remove_office_only_edition + '''" #optionally removes only this edition of office

#The following variables are ignored if $Install_Office_WITH_ODT_XMLFILE is configured
$Install_Office_Channel = "''' + Install_Office_Channel + '''" #Defines which channel to use for installing Office
$Install_Office_Channel = $Install_Office_Channel.split(',') | Select -First 1
$Install_Office_Source_Path = "''' + itsm.getParameter('Install_Office_Source_Path') + '''" #Path where Office installation files will be saved and deployed from.
$Install_Office_Exclude_Apps = "''' + excluded_apps + '''"
$Install_Office_Exclude_Apps = $Install_Office_Exclude_Apps.split(',')
$Install_Office_Org_Name = "''' + itsm.getParameter('Install_Office_Org_Name') + '''"
$Install_Office_Shared_Computer_Licensing = ''' + itsm.getParameter('Install_Office_Shared_Computer_Licensing') + '''
$Install_Office_Remove_Previous_Intalls = ''' + itsm.getParameter('Install_Office_Remove_Previous_Installs') + '''
$Install_Office_Accept_EULA = ''' + itsm.getParameter('Install_Office_Accept_EULA') + '''
$Install_Office_Enable_Updates = ''' + itsm.getParameter('Install_Office_Enable_Updates') + '''
$Install_Office_Display_Install = ''' + itsm.getParameter('Install_Office_Display_Install') + '''
$Install_Office_AUTO_ACTIVATE = ''' + itsm.getParameter('Install_Office_AUTO_ACTIVATE') + '''
$Install_Office_FORCE_APPSHUTDOWN = ''' + itsm.getParameter('Install_Office_FORCE_APPSHUTDOWN') + '''
#---------------------------------------------- 
#endregion Function Select Variables
#----------------------------------------------

#---------------------------------------------- 
#region Global Variables - DO NOT CHANGE UNLESS YOU KNOW WHAT YOU'R DOING
#----------------------------------------------
$script:stopwatch = [system.diagnostics.stopwatch]::StartNew() #starts stopwatch timer 
$enablelogs = 1
$logfile_directory = "''' + itsm.getParameter('LogFile_Directory') + '''"
$Copy_SetupLog = ''' + itsm.getParameter('Copy_SetupLog') + '''
$LogFileAppend = $true #enables appending to existing log file. Disabled overwrites previous log file content
$logtime = $true
$logdateformat = 'MM/dd/yyyy h:mm:ss tt' # sets the date/time appearance format for log file and console messages
$update_modules = $false # enables checking for and updating all required modules for this script. Potentially adds a few seconds to total runtime but ensures all modules are the latest
$force_modules = $false # enables installing and importing of a module even if it is already. Should not be used unless troubleshooting module issues 
$Local_Test_Run = $false # enables using script on a local machine for testing. When launched, lauches a window (show-command) that presents all configurable variable options
#---------------------------------------------- 
#endregion Global Variables - DO NOT CHANGE UNLESS YOU KNOW WHAT YOU'R DOING
#----------------------------------------------

#############################################################################
#endregion Configurable Script Parameters
#############################################################################

#############################################################################
#region global functions - Functions that are commonly used across scripts and must be run first
#############################################################################
#---------------------------------------------- 
#region Request-Config Function
#----------------------------------------------
function Request-Config 
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][validateset("True","False")]$Install_Office,
        [validateset("O365BusinessRetail","O365ProPlusRetail")][string]$Install_Office_Edition = "O365BusinessRetail",
        [string]$Install_Office_Download_Path = "C:\\Office365Install",
        [string]$Install_Office_Source_Path = "Internet",
        [validateset("True","False")]
        [switch]$Install_Office_WITH_ODT = $false,
        [string]$Install_Office_WITH_ODT_XMLFILE,
        [validateset("True","False")]
        [switch]$Remove_Office_Only = $false,
        [string]$Remove_Office_Only_Edition,
        $Install_Office_Exclude_Apps = "Groove,Lync,Teams",
        [ValidateSet("Current","CurrentPreview","SemiAnnual","SemiAnnualPreview","BetaChannel","MonthlyEnterprise")]$Install_Office_Channel = "Current", #Defines which channel to use for installing Office         
        [string]$Install_Office_Org_Name = "EZTechhelp Company",
        [validateset("True","False")]
        [switch]$Install_Office_Remove_Previous_Intalls = $true,
        [validateset("True","False")]
        [switch]$Install_Office_Shared_Computer_Licensing = $false,
        [validateset("True","False")]
        [switch]$Install_Office_Accept_EULA = $true,
        [validateset("True","False")]
        [switch]$Install_Office_Enable_Updates = $true,        
        [validateset("True","False")]
        [switch]$Install_Office_Display_Install = $false,        
        [validateset("True","False")]
        [switch]$Install_Office_AUTO_ACTIVATE = $true,        
        [validateset("True","False")]
        [switch]$Install_Office_FORCE_APPSHUTDOWN = $true,
        [validateset("True","False")]
        [switch]$Copy_SetupLog = $true,
        [string]$logfile_directory = "c:\\logs"                
    )
    $ParameterList = (Get-Command -Name $MyInvocation.InvocationName).Parameters;
    foreach ($key in $ParameterList.keys)
    {
        $var = Get-Variable -Name $key -ErrorAction SilentlyContinue;
        $var
    }
}
if($Local_Test_Run){
  try
  {
    $result = Invoke-Expression (Show-Command {Request-Config} -Width 800 -NoCommonParameter -PassThru)
  }
  catch
  {
    write-host "[ERROR] An exception occured setting configuration values:`n | $($_.exception.message)`n | $($_.InvocationInfo.positionmessage)`n | $($_.ScriptStackTrace)`n" -ForegroundColor Red
    exit
  }

  foreach ($r in $result)
  {
    if($r.value -is [switch])
    {
      $value = $($r.value).ToBool()
    }
    else
    {
      $value = $($r.value)
    }
    Set-Variable -name $r.name -Value $value
  }
  $Install_Office_Exclude_Apps = $Install_Office_Exclude_Apps.split(',')
}
#---------------------------------------------- 
#endregion Request-Config Function
#----------------------------------------------

#---------------------------------------------- 
#region Get-ThisScriptInfo Function
#----------------------------------------------
function Get-ThisScriptInfo
{
  $Invocation = (Get-Variable MyInvocation -Scope 1).Value
  $ScriptPath = $PSCommandPath
  if(!$ScriptPath)
  {   
    $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand
    $thisScript = @{File = Get-ChildItem $ScriptPath; Contents = $Invocation.MyCommand}
  }
  else
  {$thisScript = @{File = Get-ChildItem $ScriptPath; Contents = $Invocation.MyCommand.ScriptContents}}
  If ($thisScript.Contents -Match '^\\s*\\<#([\\s\\S]*?)#\\>') 
  {$thisScript.Help = $Matches[1].Trim()}
  [RegEx]::Matches($thisScript.Help, "(^|[`r`n])\\s*\\.(.+)\\s*[`r`n]|$") | ForEach-Object {
    If ($Caption) 
    {$thisScript.$Caption = $thisScript.Help.SubString($Start, $_.Index - $Start)}
    $Caption = $_.Groups[2].ToString().Trim()
    $Start = $_.Index + $_.Length
  }
  $thisScript.Version = $thisScript.Version.Trim()
  $thisScript.Name = $thisScript.Name.Trim()
  $thisScript.credits = $thisScript.credits -split("`n") | ForEach-Object {$_.trim()}
  $thisScript.SYNOPSIS = $thisScript.SYNOPSIS -split("`n") | ForEach-Object {$_.trim()}
  $thisScript.Description = $thisScript.Description -split("`n") | ForEach-Object {$_.trim()}
  $thisScript.Notes = $thisScript.Notes -split("`n") | ForEach-Object {$_.trim()}
  $thisScript.Path = $thisScript.File.FullName; $thisScript.Folder = $thisScript.File.DirectoryName; $thisScript.BaseName = $thisScript.File.BaseName
  $thisScript.Arguments = (($Invocation.Line + ' ') -Replace ('^.*\\\\' + $thisScript.File.Name.Replace('.', '\\.') + "['"" ]"), '').Trim()
  [System.Collections.Generic.List[String]]$FX_NAMES = New-Object System.Collections.Generic.List[String]
  if(!([System.String]::IsNullOrWhiteSpace($thisScript.file)))
  { 
    Select-String -Path $thisScript.file -Pattern "function" |
    ForEach-Object {
      [System.Text.RegularExpressions.Regex] $regexp = New-Object Regex("(function)( +)([\\w-]+)")
      [System.Text.RegularExpressions.Match] $match = $regexp.Match("$_")
      if($match.Success)
      {
        $FX_NAMES.Add("$($match.Groups[3])")
      }  
    }
    $thisScript.functions = $FX_NAMES.ToArray()   
  }
  $Script_Temp_Folder = "$env:TEMP\\$($thisScript.Name)"
  if(!(Test-Path $Script_Temp_Folder))
  {
    try
    {$null = New-Item $Script_Temp_Folder -ItemType Directory -Force}
    catch
    {Write-EZLogs "[ERROR] Exception creating script temp directory $Script_Temp_Folder - $_" -ShowTime -color Red}
  }
  return $thisScript
}
$thisScript = Get-ThisScriptInfo
#---------------------------------------------- 
#endregion Get-ThisScriptInfo Function
#----------------------------------------------

#---------------------------------------------- 
#region Start EZLogs Function
#----------------------------------------------
function Start-EZLogs
{
  param (
    [switch]$Verboselog,
    [string]$Logfile_Directory,
    [string]$Logfile_Name,
    [string]$Script_Name,
    $thisScript,
    [string]$Script_Description,
    [string]$Script_Version,
    [string]$ScriptPath,
    [switch]$Start_Timer = $true
  )
  if(!$ScriptPath){$ScriptPath = $((Get-PSCallStack).ScriptName | where {$_ -notmatch '.psm1'} | select -First 1)}
  if(!$thisScript){$thisScript = Get-thisScriptinfo -ScriptPath $ScriptPath -No_Script_Temp_Folder}
  if($Start_Timer){$Global:globalstopwatch = [system.diagnostics.stopwatch]::StartNew()}
  if(!$logfile_name){$logfile_name = "$($thisScript.Name)-$($thisScript.Version).log"}
  if(!$Script_Name){$Script_Name = $($thisScript.Name)}
  if(!$Script_Description){$Script_Description = $($thisScript.SYNOPSIS)}
  if(!$Script_Version){$Script_Version = $($thisScript.Version)}
  if(!$logfile_Directory){$Logfile_Directory = $($thisScript.TempFolder)}
  $script:logfile = [System.IO.Path]::Combine($logfile_directory, $logfile_name)
  if (!(Test-Path -LiteralPath $logfile 2> $null))
  {$null = New-Item -Path $logfile_directory -ItemType directory -Force}
  $OriginalPref = $ProgressPreference
  $ProgressPreference = 'SilentlyContinue'
  $Computer_Info = Get-WmiObject Win32_ComputerSystem | Select-Object *
  $OS_Info = Get-CimInstance Win32_OperatingSystem | Select-Object *
  $CPU_Name = (Get-WmiObject Win32_Processor -Property 'Name').name
  $ProgressPreference = $OriginalPref
  $logheader = @"
`n###################### Logging Enabled ######################
Script Name          : $Script_Name
Synopsis             : $Script_Description
Log File             : $logfile
Version              : $Script_Version
Current Username     : $env:username
Powershell           : $($PSVersionTable.psversion)($($PSVersionTable.psedition))
Computer Name        : $env:computername
Operating System     : $($OS_Info.Caption)($($OS_Info.Version))
CPU                  : $($CPU_Name)
RAM                  : $([Math]::Round([int64]($computer_info.TotalPhysicalMemory)/1MB,2)) GB (Available: $([Math]::Round([int64]($OS_Info.FreePhysicalMemory)/1MB,2)) GB)
Manufacturer         : $($computer_info.Manufacturer)
Model                : $($computer_info.Model)
Serial Number        : $((Get-WmiObject Win32_BIOS | Select-Object SerialNumber).SerialNumber)
Domain               : $($computer_info.Domain)
Install Date         : $($OS_Info.InstallDate)
Last Boot Up Time    : $($OS_Info.LastBootUpTime)
Local Date/Time      : $($OS_Info.LocalDateTime)
Windows Directory    : $($OS_Info.WindowsDirectory)
###################### Logging Started - [$(Get-Date)] ##########################
"@

  Write-Output $logheader | Out-File -FilePath $logfile -Encoding unicode -Append
  Write-Host "#### Executing $($thisScript.Name) - v$($thisScript.Version) ####" -ForegroundColor Black -BackGroundColor yellow
  Write-Host " | $($thisScript.SYNOPSIS)"  
  Write-host " | Logging is enabled. Log file: $logfile"
  return $logfile
}
#---------------------------------------------- 
#endregion Start EZLogs Function
#----------------------------------------------

#---------------------------------------------- 
#region Load-Modules Function
#----------------------------------------------
function Load-Modules ($modules,$force,$update,$enablelogs) 
{
  #Make sure we can download and install modules through NuGet
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  if (Get-PackageProvider | Where-Object {$_.Name -eq 'Nuget'}) 
  {
    Write-Output ' | Required PackageProvider Nuget is installed.' -OutVariable message;if($enablelogs){$message | Out-File -FilePath $logfile -Encoding unicode -Append}
  }
  else
  {
    try
    {
      Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
      Register-PackageSource -Name nuget.org -Location https://www.nuget.org/api/v2 -ProviderName NuGet
    }
    catch
    {
      Write-Error "[Load-Module ERROR] $_`n" -ErrorVariable messageerror;if($enablelogs){$messageerror | Out-File -FilePath $logfile -Encoding unicode -Append}
    }
  }
  #Install latest version of PowerShellGet
  if (Get-Module 'PowershellGet' | Where-Object {$_.Version -lt '2.2.5'})
  {
    Write-Output ' | PowershellGet version too low, updating to 2.2.5' -OutVariable message;if($enablelogs){$message | Out-File -FilePath $logfile -Encoding unicode -Append}
    Install-Module -Name 'PowershellGet' -MinimumVersion '2.2.5' -Force 
  }
  if(-not [string]::IsNullOrEmpty($modules)){
    foreach ($m in $modules){  
      if (Get-Module | Where-Object {$_.Name -eq $m}){
        Write-Output " | Required Module $m is imported." -OutVariable message;if($enablelogs){$message | Out-File -FilePath $logfile -Encoding unicode -Append}
        if ($force){
          Write-Output " | Force parameter applied - Installing $m" -OutVariable message;if($enablelogs){$message | Out-File -FilePath $logfile -Encoding unicode -Append}
          Install-Module -Name $m -Scope AllUsers -Force -Verbose 
        }
      }
      else {
        #If module is not imported, but available on disk set module autoloading when needed/called 
        foreach($path in $env:PSModulePath -split ";"){
          if(Test-Path -literalpath $Path){
            $module_list += Get-ChildItem $path #using get-childitem against PSModulePath is much faster than using Get-Module -ListAvailable. Potential downside is it doesnt verify module is valid only that it exists
          }
        }
        if($module_list -match $m){
          Write-Output " | Required Module $m is available on disk." -OutVariable message;if($enablelogs){$message | Out-File -FilePath $logfile -Encoding unicode -Append}
          $PSModuleAutoLoadingPreference = 'ModuleQualified'
          if($update){
            Write-Output " | Updating module: $m" -OutVariable message;if($enablelogs){$message | Out-File -FilePath $logfile -Encoding unicode -Append}
            Update-Module -Name $m -Force -ErrorAction Continue
          }
          if($force){
            if($enablelogs){Write-Output " | Force parameter applied - Importing $m" | Out-File -FilePath $logfile -Encoding unicode -Append}
            Import-Module $m -Verbose -force -Scope Global
          }
        }
        else {
          #If module is not imported, not available on disk, but is in online gallery then install and import
          if (Find-Module -Name $m | Where-Object {$_.Name -eq $m}) {
            try{
              Install-Module -Name $m -Force -Verbose -Scope AllUsers -AllowClobber
              Import-Module $m -Verbose -force -Scope Global
            }
            catch{Write-Error "[Load-Module ERROR] $_" -ErrorVariable messageerror;if($enablelogs){$messageerror | Out-File -FilePath $logfile -Encoding unicode -Append}}      
          }
          else {
            #If module is not imported, not available and not in online gallery then abort
            Write-Error "[Load-Module ERROR] Required module $m not imported, not available and not in online gallery, exiting." -ErrorVariable messageerror;if($enablelogs){$messageerror | Out-File -FilePath $logfile -Encoding unicode -Append}
            EXIT 1
          }
        }
      }
    }
  }
}
#---------------------------------------------- 
#endregion Load-Modules Function
#----------------------------------------------

#---------------------------------------------- 
#region Write-EZLogs Function
#----------------------------------------------
function Write-EZLogs 
{
  [CmdletBinding(DefaultParameterSetName = 'text')]
  param (
    [string]$text,
    [switch]$VerboseDebug,
    [switch]$enablelogs = ([System.Convert]::ToBoolean($enablelogs)),
    [string]$logfile = $logfile,
    [switch]$Warning,
    $CatchError,
    [switch]$logOnly,
    [string]$DateTimeFormat = 'MM/dd/yyyy h:mm:ss tt',
    [ValidateSet('Black','Blue','Cyan','Gray','Green','Magenta','Red','White','Yellow','DarkBlue','DarkCyan','DarkGreen','DarkMagenta','DarkRed','DarkYellow')]
    [string]$color = 'white',
    [ValidateSet('Black','Blue','Cyan','Gray','Green','Magenta','Red','White','Yellow','DarkBlue','DarkCyan','DarkGreen','DarkMagenta','DarkRed','DarkYellow')]
    [string]$foregroundcolor,
    [switch]$showtime,
    [switch]$logtime,
    [switch]$NoNewLine,
    [int]$StartSpaces,
    [string]$Separator,
    [ValidateSet('Black','Blue','Cyan','Gray','Green','Magenta','Red','White','Yellow','DarkBlue','DarkCyan','DarkGreen','DarkMagenta','DarkRed','DarkYellow')]
    [string]$BackgroundColor,
    [int]$linesbefore,
    [int]$linesafter
  )
  if(!$logfile){$logfile = Start-EZlogs -thisScript $thisScript}
  if($showtime -and !$logtime){$logtime = $true}else{$logtime = $false}
  if($foregroundcolor){$color = $foregroundcolor}
  if($BackgroundColor){$BackgroundColor_param = $BackgroundColor}else{$BackgroundColor_param = $null}
  if($LinesBefore -ne 0){ for ($i = 0; $i -lt $LinesBefore; $i++) {write-host "`n" -NoNewline;if($enablelogs){write-output "" | Out-File -FilePath $logfile -Encoding unicode -Append}}}
  if(-not [string]::IsNullOrEmpty($CatchError)){$text = "[ERROR] $text`:`n | $($CatchError.exception.message)`n | $($CatchError.InvocationInfo.positionmessage)`n | $($CatchError.ScriptStackTrace)`n";$color = "red"}
  if($enablelogs)
  {
    if($VerboseDebug -and $warning)
    {
      $tmp = [System.IO.Path]::GetTempFileName();
      Write-Host -Object "[$([datetime]::Now.ToString($DateTimeFormat))] " -NoNewline;Write-Warning ($wrn = "$text");Write-Output "[$(Get-Date -Format $DateTimeFormat)] [WARNING] $wrn" | Out-File -FilePath $logfile -Encoding unicode -Append -Verbose:$VerboseDebug 4>$tmp
      $result = "[DEBUG] $(Get-Content $tmp)" | Out-File $logfile -Encoding unicode -Append;Remove-Item $tmp   
    }
    elseif($Warning)
    {
      if($logOnly)
      {
        if($showtime){
          Write-Output "[$(Get-Date -Format $DateTimeFormat)] " | Out-File -FilePath $logfile -Encoding unicode -Append -NoNewline
        }
        Write-Output "[WARNING] $text" | Out-File -FilePath $logfile -Encoding unicode -Append -NoNewline:$NoNewLine
      }
      else
      {
        if($showtime){
          Write-Host -Object "[$([datetime]::Now.ToString($DateTimeFormat))] " -NoNewline;if($enablelogs){"[$([datetime]::Now.ToString($DateTimeFormat))] " | Out-File -FilePath $logfile -Encoding unicode -Append -NoNewline}
        }
        Write-Warning ($wrn = "$text");Write-Output "[WARNING] $wrn" | Out-File -FilePath $logfile -Encoding unicode -Append
      }      
    }
    elseif($VerboseDebug)
    {
      if($showtime){
        Write-Host -Object "[$([datetime]::Now.ToString($DateTimeFormat))] " -NoNewline;if($enablelogs){"[$([datetime]::Now.ToString($DateTimeFormat))] " | Out-File -FilePath $logfile -Encoding unicode -Append -NoNewline}
      }
      if($BackGroundColor){
        Write-Host -Object "[DEBUG] $text" -ForegroundColor:Cyan -NoNewline:$NoNewLine -BackgroundColor:$BackGroundColor;if($enablelogs){"[DEBUG] $text" | Out-File -FilePath $logfile -Encoding unicode -Append -NoNewline:$NoNewLine}
      }else{
        Write-Host -Object "[DEBUG] $text" -ForegroundColor:Cyan -NoNewline:$NoNewLine;if($enablelogs){"[DEBUG] $text" | Out-File -FilePath $logfile -Encoding unicode -Append -NoNewline:$NoNewLine}
      }    
    }
    else
    {
      if($logOnly)
      {
        if($showtime){
          Write-Output "[$(Get-Date -Format $DateTimeFormat)] " | Out-File -FilePath $logfile -Encoding unicode -Append -NoNewline
        }
        Write-Output "$text" | Out-File -FilePath $logfile -Encoding unicode -Append -NoNewline:$NoNewLine 
      }
      else
      {
        if($showtime){
          Write-Host -Object "[$([datetime]::Now.ToString($DateTimeFormat))] " -NoNewline;if($enablelogs){"[$([datetime]::Now.ToString($DateTimeFormat))] " | Out-File -FilePath $logfile -Encoding unicode -Append -NoNewline}
        }
        if($BackGroundColor){
          Write-Host -Object $text -ForegroundColor:$Color -NoNewline:$NoNewLine -BackgroundColor:$BackGroundColor;if($enablelogs){$text | Out-File -FilePath $logfile -Encoding unicode -Append -NoNewline:$NoNewLine}
        }else{
          Write-Host -Object $text -ForegroundColor:$Color -NoNewline:$NoNewLine;if($enablelogs){$text | Out-File -FilePath $logfile -Encoding unicode -Append -NoNewline:$NoNewLine}
        }
      }
    }
  }
  else
  {
    if($warning)
    {
      if($showtime){
        Write-Host -Object "[$([datetime]::Now.ToString($DateTimeFormat))] " -NoNewline
      }
      Write-Warning ($wrn = "$text")
    }
    else
    {
      if($showtime){
        Write-Host -Object "[$([datetime]::Now.ToString($DateTimeFormat))] " -NoNewline
      }
      if($BackGroundColor){
        Write-Host -Object $text -ForegroundColor:$Color -NoNewline:$NoNewLine -BackgroundColor:$BackGroundColor
      }else{
        Write-Host -Object $text -ForegroundColor:$Color -NoNewline:$NoNewLine
      }    
    }     
  }
  if($LinesAfter -ne 0){ for ($i = 0; $i -lt $LinesAfter; $i++) {write-host "`n" -NoNewline;if($enablelogs){write-output "" | Out-File -FilePath $logfile -Encoding unicode -Append}}}
}
#---------------------------------------------- 
#endregion Write-EZLogs Function
#----------------------------------------------

#---------------------------------------------- 
#region Stop EZLogs
#----------------------------------------------
function Stop-EZLogs
{
  param (
    $ErrorSummary,
    [string]$logdateformat = 'MM/dd/yyyy h:mm:ss tt',
    [string]$logfile = $logfile,
    [switch]$logOnly,
    [switch]$enablelogs = $true,
    [switch]$stoptimer,
    [switch]$clearErrors
  )
  if($ErrorSummary)
  {
    Write-Output "`n`n[-----ALL ERRORS------]" | Out-File -FilePath $logfile -Encoding unicode -Append
    $e_index = 0
    foreach ($e in $ErrorSummary)
    {
      $e_index++
      Write-Output "[ERROR $e_index Message] =========================================================================`n$($e.exception.message)`n$($e.InvocationInfo.positionmessage)`n$($e.ScriptStackTrace)`n`n" | Out-File -FilePath $logfile -Encoding unicode -Append
    }
    if($clearErrors)
    {
      $error.Clear()
    }
  }
  if($logOnly){Write-Output "`n======== Total Script Execution Time ========" | Out-File -FilePath $logfile -Encoding unicode -Append}else{Write-EZLogs "`n======== Total Script Execution Time ========" -enablelogs:$enablelogs -LogTime:$false}
  if($logOnly){Write-Output "Hours        : $($globalstopwatch.elapsed.hours)`nMinutes      : $($globalstopwatch.elapsed.Minutes)`nSeconds      : $($globalstopwatch.elapsed.Seconds)`nMilliseconds : $($globalstopwatch.elapsed.Milliseconds)" | Out-File -FilePath $logfile -Encoding unicode -Append}else{Write-EZLogs "Hours        : $($globalstopwatch.elapsed.hours)`nMinutes      : $($globalstopwatch.elapsed.Minutes)`nSeconds      : $($globalstopwatch.elapsed.Seconds)`nMilliseconds : $($globalstopwatch.elapsed.Milliseconds)" -enablelogs:$enablelogs -LogTime:$false}
  if($stoptimer)
  {
    $($globalstopwatch.stop())
    $($globalstopwatch.reset()) 
  }
  Write-Output "###################### Logging Finished - [$(Get-Date -Format $logdateformat)] ######################`n" | Out-File -FilePath $logfile -Encoding unicode -Append
}  
#---------------------------------------------- 
#endregion Stop EZLogs
#----------------------------------------------

#---------------------------------------------- 
#region Use Run-As Function
#----------------------------------------------
function Use-RunAs 
{    
  # Check if script is running as Adminstrator and if not use RunAs 
  # Use Check Switch to check if admin 
  # http://gallery.technet.microsoft.com/scriptcenter/63fd1c0d-da57-4fb4-9645-ea52fc4f1dfb
    
  param([Switch]$Check) 
  $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator') 
  if ($Check) { return $IsAdmin }     
  if ($MyInvocation.ScriptName -ne '') 
  {  
    if (-not $IsAdmin)  
    {  
      write-ezlogs 'Script is not running as administrator, attempting to re-run with RunAs' -showtime -Warning
      try 
      {  
        $arg = "-file `"$($thisScript.File)`"" 
        Start-Process "$psHome\\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'  
      } 
      catch 
      { 
        write-ezlogs 'Failed to restart script with runas' -showtime -Warning
        break               
      } 
      exit # Quit this session of powershell 
    }  
  }  
  else  
  {  
    Write-EZLogs 'Script must be saved as a .ps1 file first' -showtime -LogFile $logfile -LinesAfter 1 -Warning  
    break  
  }  
}
#---------------------------------------------- 
#endregion Use Run-As Function
#----------------------------------------------

#############################################################################
#endregion global functions
#############################################################################

#############################################################################
#region Core functions - The primary functions specific to this script
#############################################################################

#---------------------------------------------- 
#region Get-OfficeVersion Function
#----------------------------------------------
Function Get-OfficeVersion 
{
  <#
      .Synopsis
      Gets the Office Version installed on the computer
      .DESCRIPTION
      This function will query the local or a remote computer and return the information about Office Products installed on the computer
      .NOTES   
      Name: Get-OfficeVersion
      Version: 1.0.5
      DateCreated: 2015-07-01
      DateUpdated: 2016-10-14
      .LINK
      https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts
      .PARAMETER ComputerName
      The computer or list of computers from which to query 
      .PARAMETER ShowAllInstalledProducts
      Will expand the output to include all installed Office products
      .EXAMPLE
      Get-OfficeVersion
    
      Will return the locally installed Office product
      .EXAMPLE
      Get-OfficeVersion -ComputerName client01,client02
    
      Will return the installed Office product on the remote computers
      .EXAMPLE
      Get-OfficeVersion | select *
    
      Will return the locally installed Office product with all of the available properties
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true, Position = 0)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    [switch]$ShowAllInstalledProducts,
    [System.Management.Automation.PSCredential]$Credentials
  )

  begin {
    $HKLM = [UInt32] '0x80000002'
    $HKCR = [UInt32] '0x80000000'

    $excelKeyPath = 'Excel\\DefaultIcon'
    $wordKeyPath = 'Word\\DefaultIcon'
   
    $installKeys = 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall',
    'SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall'

    $officeKeys = 'SOFTWARE\\Microsoft\\Office',
    'SOFTWARE\\Wow6432Node\\Microsoft\\Office'

    $defaultDisplaySet = 'DisplayName','Version', 'ComputerName'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
  }

  process {

    $results = new-object PSObject[] 0
    $MSexceptionList = 'mui','visio','project','proofing','visual'

    foreach ($computer in $ComputerName) {
      if ($Credentials) {
        $os = Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials
      } else {
        $os = Get-WMIObject win32_operatingsystem -computername $computer
      }

      $osArchitecture = $os.OSArchitecture

      if ($Credentials) {
        $regProv = Get-Wmiobject -list 'StdRegProv' -namespace root\\default -computername $computer -Credential $Credentials
      } else {
        $regProv = Get-Wmiobject -list 'StdRegProv' -namespace root\\default -computername $computer
      }

      [System.Collections.ArrayList]$VersionList = New-Object -TypeName System.Collections.ArrayList
      [System.Collections.ArrayList]$PathList = New-Object -TypeName System.Collections.ArrayList
      [System.Collections.ArrayList]$PackageList = New-Object -TypeName System.Collections.ArrayList
      [System.Collections.ArrayList]$ClickToRunPathList = New-Object -TypeName System.Collections.ArrayList
      [System.Collections.ArrayList]$ConfigItemList = New-Object -TypeName  System.Collections.ArrayList
      $ClickToRunList = new-object PSObject[] 0

      foreach ($regKey in $officeKeys) {
        $officeVersion = $regProv.EnumKey($HKLM, $regKey)
        foreach ($key in $officeVersion.sNames) {
          if ($key -match '\\d{2}\\.\\d') {
            if (!$VersionList.Contains($key)) {
              $AddItem = $VersionList.Add($key)
            }

            $path = join-path $regKey $key

            $configPath = join-path $path 'Common\\Config'
            $configItems = $regProv.EnumKey($HKLM, $configPath)
            if ($configItems) {
              foreach ($configId in $configItems.sNames) {
                if ($configId) {
                  $Add = $ConfigItemList.Add($configId.ToUpper())
                }
              }
            }

            $cltr = New-Object -TypeName PSObject
            $cltr | Add-Member -MemberType NoteProperty -Name InstallPath -Value ''
            $cltr | Add-Member -MemberType NoteProperty -Name UpdatesEnabled -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name UpdateUrl -Value ''
            $cltr | Add-Member -MemberType NoteProperty -Name ExcludedApps -Value ''
            $cltr | Add-Member -MemberType NoteProperty -Name TeamsAddon -Value ''
            $cltr | Add-Member -MemberType NoteProperty -Name Activate -Value ''
            $cltr | Add-Member -MemberType NoteProperty -Name SharedComputerLicensing -Value ''
            $cltr | Add-Member -MemberType NoteProperty -Name DeviceBasedLicensing -Value ''
            $cltr | Add-Member -MemberType NoteProperty -Name StreamingFinished -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name Platform -Value ''
            $cltr | Add-Member -MemberType NoteProperty -Name ClientCulture -Value ''
            
            $packagePath = join-path $path 'Common\\InstalledPackages'
            $clickToRunPath = join-path $path 'ClickToRun\\Configuration'
            $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, 'InstallationPath').sValue

            [string]$officeLangResourcePath = join-path  $path 'Common\\LanguageResources'
            $mainLangId = $regProv.GetDWORDValue($HKLM, $officeLangResourcePath, 'SKULanguage').uValue
            if ($mainLangId) {
              $mainlangCulture = [globalization.cultureinfo]::GetCultures('allCultures') | where {$_.LCID -eq $mainLangId}
              if ($mainlangCulture) {
                $cltr.ClientCulture = $mainlangCulture.Name
              }
            }

            [string]$officeLangPath = join-path  $path 'Common\\LanguageResources\\InstalledUIs'
            $langValues = $regProv.EnumValues($HKLM, $officeLangPath)
            if ($langValues) {
              foreach ($langValue in $langValues) {
                $langCulture = [globalization.cultureinfo]::GetCultures('allCultures') | where {$_.LCID -eq $langValue}
              } 
            }

            if ($virtualInstallPath) {

            } else {
              $clickToRunPath = join-path $regKey 'ClickToRun\\Configuration'
              $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, 'InstallationPath').sValue
            }

            if ($virtualInstallPath) {
              if (!$ClickToRunPathList.Contains($virtualInstallPath.ToUpper())) {
                $AddItem = $ClickToRunPathList.Add($virtualInstallPath.ToUpper())
              }

              $cltr.InstallPath = $virtualInstallPath
              $cltr.StreamingFinished = $regProv.GetStringValue($HKLM, $clickToRunPath, 'StreamingFinished').sValue
              $cltr.UpdatesEnabled = $regProv.GetStringValue($HKLM, $clickToRunPath, 'UpdatesEnabled').sValue
              $cltr.UpdateUrl = $regProv.GetStringValue($HKLM, $clickToRunPath, 'UpdateChannel').sValue
              $cltr.Platform = $regProv.GetStringValue($HKLM, $clickToRunPath, 'Platform').sValue
              $cltr.ClientCulture = $regProv.GetStringValue($HKLM, $clickToRunPath, 'ClientCulture').sValue
              $cltr.TeamsAddon = $regProv.GetStringValue($HKLM, $clickToRunPath, 'TeamsAddon').sValue
              $cltr.ExcludedApps = $regProv.GetStringValue($HKLM, $clickToRunPath, 'O365BusinessRetail.ExcludedApps').sValue
              $cltr.DeviceBasedLicensing = $regProv.GetStringValue($HKLM, $clickToRunPath, 'O365ProPlusRetail.DeviceBasedLicensing').sValue
              if(!$cltr.ExcludedApps)
              {
                $cltr.ExcludedApps = $regProv.GetStringValue($HKLM, $clickToRunPath, 'O365ProPlusRetail.ExcludedApps').sValue
              }
              if(!$cltr.DeviceBasedLicensing)
              {
                $cltr.DeviceBasedLicensing = $regProv.GetStringValue($HKLM, $clickToRunPath, 'O365BusinessRetail.DeviceBasedLicensing').sValue
              }
              $cltr.SharedComputerLicensing = $regProv.GetStringValue($HKLM, $clickToRunPath, 'SharedComputerLicensing').sValue
              $cltr.Activate = $regProv.GetStringValue($HKLM, $clickToRunPath, 'Activate').sValue
              $ClickToRunList += $cltr
            }

            $packageItems = $regProv.EnumKey($HKLM, $packagePath)
            $officeItems = $regProv.EnumKey($HKLM, $path)

            foreach ($itemKey in $officeItems.sNames) {
              $itemPath = join-path $path $itemKey
              $installRootPath = join-path $itemPath 'InstallRoot'

              $filePath = $regProv.GetStringValue($HKLM, $installRootPath, 'Path').sValue
              if (!$PathList.Contains($filePath)) {
                $AddItem = $PathList.Add($filePath)
              }
            }

            foreach ($packageGuid in $packageItems.sNames) {
              $packageItemPath = join-path $packagePath $packageGuid
              $packageName = $regProv.GetStringValue($HKLM, $packageItemPath, '').sValue
            
              if (!$PackageList.Contains($packageName)) {
                if ($packageName) {
                  $AddItem = $PackageList.Add($packageName.Replace(' ', '').ToLower())
                }
              }
            }

          }
        }
      }

      foreach ($regKey in $installKeys) {
        $keyList = new-object System.Collections.ArrayList
        $keys = $regProv.EnumKey($HKLM, $regKey)

        foreach ($key in $keys.sNames) {
          $path = join-path $regKey $key
          $installPath = $regProv.GetStringValue($HKLM, $path, 'InstallLocation').sValue
          if (!($installPath)) { continue }
          if ($installPath.Length -eq 0) { continue }

          $buildType = '64-Bit'
          if ($osArchitecture -eq '32-bit') {
            $buildType = '32-Bit'
          }

          if ($regKey.ToUpper().Contains('Wow6432Node'.ToUpper())) {
            $buildType = '32-Bit'
          }

          if ($key -match '{.{8}-.{4}-.{4}-1000-0000000FF1CE}') {
            $buildType = '64-Bit' 
          }

          if ($key -match '{.{8}-.{4}-.{4}-0000-0000000FF1CE}') {
            $buildType = '32-Bit' 
          }

          if ($modifyPath) {
            if ($modifyPath.ToLower().Contains('platform=x86')) {
              $buildType = '32-Bit'
            }

            if ($modifyPath.ToLower().Contains('platform=x64')) {
              $buildType = '64-Bit'
            }
          }

          $primaryOfficeProduct = $false
          $officeProduct = $false
          foreach ($officeInstallPath in $PathList) {
            if ($officeInstallPath) {
              try{
                $installReg = '^' + $installPath.Replace('\\', '\\\\')
                $installReg = $installReg.Replace('(', '\\(')
                $installReg = $installReg.Replace(')', '\\)')
                if ($officeInstallPath -match $installReg) { $officeProduct = $true }
              } catch {}
            }
          }

          if (!$officeProduct) { continue }
           
          $name = $regProv.GetStringValue($HKLM, $path, 'DisplayName').sValue          

          $primaryOfficeProduct = $true
          if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains('MICROSOFT OFFICE')) {
            foreach($exception in $MSexceptionList){
              if($name.ToLower() -match $exception.ToLower()){
                $primaryOfficeProduct = $false
              }
            }
          } else {
            $primaryOfficeProduct = $false
          }

          $clickToRunComponent = $regProv.GetDWORDValue($HKLM, $path, 'ClickToRunComponent').uValue
          $uninstallString = $regProv.GetStringValue($HKLM, $path, 'UninstallString').sValue
          if (!($clickToRunComponent)) {
            if ($uninstallString) {
              if ($uninstallString.Contains('OfficeClickToRun')) {
                $clickToRunComponent = $true
              }
            }
          }

          $modifyPath = $regProv.GetStringValue($HKLM, $path, 'ModifyPath').sValue 
          $version = $regProv.GetStringValue($HKLM, $path, 'DisplayVersion').sValue
          # Get version information
          $365_update_content = Invoke-RestMethod -Uri 'https://docs.microsoft.com/en-us/officeupdates/update-history-office365-proplus-by-date' -Method Get -UseBasicParsing

          # Cast current version to a version
          $Currentbuild = [Version]$version

          # Get the version using regex
          $null = $365_update_content -match "<a href=`"(?<Channel>.+?)`".+?>Version (?<Version>\\d{4}) \\(Build $($Currentbuild.Build)\\.$($Currentbuild.Revision)\\)"

          # Output the data
          $output = [PSCustomObject]@{
            Version = $Matches['Version']
          }
          $build = $output.Version
           
          $cltrUpdatedEnabled = $NULL
          $cltrUpdateUrl = $NULL
          $clientCulture = $NULL

          [string]$clickToRun = $false

          if ($clickToRunComponent) {
            $clickToRun = $true
            if ($name.ToUpper().Contains('MICROSOFT OFFICE')) {
              $primaryOfficeProduct = $true
            }

            foreach ($cltr in $ClickToRunList) {
              if ($cltr.InstallPath) {
                if ($cltr.InstallPath.ToUpper() -eq $installPath.ToUpper()) {
                  $cltrUpdatedEnabled = $cltr.UpdatesEnabled
                  $cltrUpdateUrl = $cltr.UpdateUrl
                  $cltrExcludedApps = $cltr.ExcludedApps
                  $cltrSharedComputerLicensing = $cltr.SharedComputerLicensing
                  $cltrActivate = $cltr.Activate
                  $cltrDeviceBasedLicensing = $cltr.DeviceBasedLicensing
                  $cltrTeamsAddon = $cltr.TeamsAddon                       
                  if ($cltr.Platform -eq 'x64') {
                    $buildType = '64-Bit' 
                  }
                  if ($cltr.Platform -eq 'x86') {
                    $buildType = '32-Bit' 
                  }
                  $clientCulture = $cltr.ClientCulture
                }
              }
            }
          }
           
          if (!$primaryOfficeProduct) {
            if (!$ShowAllInstalledProducts) {
              continue
            }
          }

          $object = New-Object PSObject -Property @{DisplayName = $name; Version = $version;Build = $build; InstallPath = $installPath; ClickToRun = $clickToRun; 
            Bitness = $buildType; ComputerName = $computer; ClickToRunUpdatesEnabled = $cltrUpdatedEnabled; ClickToRunUpdateUrl = $cltrUpdateUrl;
          ClientCulture = $clientCulture; KeyName = $key;ExcludedApps = $cltrExcludedApps;SharedComputerLicensing = $cltrSharedComputerLicensing;DeviceBasedLicensing = $cltrDeviceBasedLicensing ;Activate = $cltrActivate;TeamsAddon = $cltrTeamsAddon }
          $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
          $results += $object

        }
      }
    }

    $results = Get-Unique -InputObject $results 

    return $results
  }

}
#---------------------------------------------- 
#endregion Get-OfficeVersion Function
#----------------------------------------------

#---------------------------------------------- 
#region Test-URL Function
#----------------------------------------------
function Test-URL
{
  Param(
    $address,
    [switch]$TestConnection
  )
  $uri = $address -as [System.URI]
  if($uri.AbsoluteURI -ne $null -and $uri.Scheme -match 'http|https')
  {
    if($TestConnection)
    {
      Try
      {
        $HTTPRequest = [System.Net.WebRequest]::Create($address)
        $HTTPResponse = $HTTPRequest.GetResponse()
        $HTTPStatus = [Int]$HTTPResponse.StatusCode
        If($HTTPStatus -ne 200) {
          Return $False
        }
        $HTTPResponse.Close()
      }
      Catch
      {
        Return $False
      }	
      Return $True    
    }
    else
    {
      Return $true
    }    
  }
  else
  {
    return $false
  }
}
#---------------------------------------------- 
#endregion Test-URL Function
#----------------------------------------------

#---------------------------------------------- 
#region Invoke-FileDownload Function
#----------------------------------------------
function Invoke-FileDownload
{
  Param(
    [uri]$DownloadURL,
    [string]$Download_file_name,
    [string]$Destination_File_Path,
    [string]$Download_Directory,
    [switch]$Overwrite
  )
  Try
  {
    write-ezlogs ">> Initializing Download from: $DownloadURL" -showtime -color Cyan
    if($DownloadURL -match 'sharepoint.com')
    {
      write-ezlogs ' | Download URL is a Onedrive share link' -showtime
      if($DownloadURL -notmatch '&download=1')
      {
        $DownloadURL = "$DownloadURL&download=1"
      }
    }
    if($Destination_File_Path)
    {
      $Download_Directory = Split-Path $Destination_File_Path -Parent
      $download_file_name = Split-Path $Destination_File_Path -Leaf
    }
    $download_output_file = [System.IO.Path]::Combine($Download_Directory, $download_file_name)
    $Test_Download_Directory = Test-Path $Download_Directory -PathType Container
    $test_download_output_file = Test-Path $download_output_file -PathType Leaf
    if($test_download_output_file)
    {
      if($Overwrite)
      {
        write-ezlogs " | Overwriting existing download file: $download_output_file" -showtime
      }
      else
      {
        write-ezlogs " | File to download already exists : $download_output_file | Overwite option disabled, Skipping download" -showtime -Warning
        return $download_output_file
      }
    }
    elseif (!$Test_Download_Directory) 
    {
      write-ezlogs " | Creating destination directory: $Download_Directory" -ShowTime
      $null = New-Item $Download_Directory -ItemType Directory -Force
    }
    else
    {
      write-ezlogs " | Destination directory is valid: $Download_Directory" -ShowTime 
    }
    $start_time = Get-Date
    $null = Invoke-WebRequest -Uri $DownloadURL -OutFile $download_output_file -UseBasicParsing
    write-ezlogs " | Download Time taken for file $DownloadURL : $((Get-Date).Subtract($start_time).Seconds) second(s)" -ShowTime 
    $test_download_output_file = Test-Path $download_output_file
    if($test_download_output_file)
    {
      write-ezlogs " | File successfully downloaded to $download_output_file" -ShowTime -color Green
      return $download_output_file
    }
    else
    {
      write-ezlogs " | Unable to validate downloaded file: $download_output_file" -ShowTime -Warning
      return $false
    } 
  }
  catch
  {
    write-ezlogs "[ERROR] An exception occured downloading from $DownloadURL :`n | $($_.exception.message)`n | $($_.InvocationInfo.positionmessage)`n | $($_.ScriptStackTrace)`n" -Color Red -showtime
  }
}
#---------------------------------------------- 
#endregion Invoke-FileDownload Function
#----------------------------------------------

#---------------------------------------------- 
#region Install-Office365Suite Function
#----------------------------------------------
function Invoke-Office365Setup
{

  [CmdletBinding(DefaultParameterSetName = 'XMLFile')]
  Param(
    [Parameter(ParameterSetName = 'XMLFile')][ValidateNotNullOrEmpty()][String]$ConfigurationXMLFile,
    [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE','FALSE')]$AcceptEULA = 'TRUE',
    [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE','FALSE')]$FORCEAPPSHUTDOWN = 'FALSE',
    [Parameter(ParameterSetName = 'NoXML')][ValidateSet('Current','CurrentPreview','SemiAnnual','SemiAnnualPreview','BetaChannel','MonthlyEnterprise')]$Channel = 'Current',
    [Parameter(ParameterSetName = 'NoXML')][Switch]$DisplayInstall = $False,
    [Parameter(ParameterSetName = 'NoXML')][ValidateSet('Groove','Outlook','OneNote','Access','OneDrive','Publisher','Word','Excel','PowerPoint','Teams','Lync')][Array]$ExcludeApps,
    [Parameter(ParameterSetName = 'NoXML')][ValidateSet('64','32')]$OfficeArch = '64',
    [Parameter(ParameterSetName = 'NoXML')][ValidateSet('O365ProPlusRetail','O365BusinessRetail')]$OfficeEdition = 'O365ProPlusRetail',
    [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE','FALSE')]$SharedComputerLicensing = 'FALSE',
    [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE','FALSE')]$AUTOACTIVATE = 'TRUE',
    [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE','FALSE')]$EnableUpdates = 'TRUE',
    [Parameter(ParameterSetName = 'NoXML')][String]$OrgName,
    [Parameter(ParameterSetName = 'NoXML')][String]$SourcePath,
    [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE','FALSE')]$PinItemsToTaskbar = 'TRUE',
    [Parameter(ParameterSetName = 'NoXML')][Switch]$RemoveMSI = $true,
    [Parameter(ParameterSetName = 'NoXML')][Switch]$RemoveOnly = $false,
    [Parameter(ParameterSetName = 'NoXML')][ValidateSet('O365ProPlusRetail','O365BusinessRetail','All')]$RemoveOnlyEdition = 'O365ProPlusRetail',
    [String]$Copy_SetupLog_directory = $logfile_directory,
    [Switch]$Copy_SetupLog,
    [String]$OfficeInstallDownloadPath = $Install_Office_Download_Path,
    [switch]$Remove_Install_Office_Download_Path = $false
  )
  
  if($RemoveOnly)
  {
    write-ezlogs '#### Removing Office365 ####' -linesbefore 1 -color yellow
  }
  else
  {
    write-ezlogs '#### Installing Office365 ####' -linesbefore 1 -color yellow
  }
  Function Generate-XMLFile{
    write-ezlogs 'Creating XML configuration file based on supplied values' -showtime
    
    if($RemoveOnly)
    {
      if(!$RemoveOnlyEdition -or $RemoveOnlyEdition -eq 'All')
      {
        $RemoveAll = 'TRUE'
        $RemoveOnlyEdition = $null
        write-ezlogs " | Remove All: $RemoveAll" -showtime
      }
      else
      {
        $RemoveAll = 'FALSE'
        write-ezlogs " | Remove Office Edition: $RemoveOnlyEdition" -showtime
      }
      $OfficeXML = [XML]@"
  <Configuration>
    <Remove All="$RemoveAll">
        <Product ID="$RemoveOnlyEdition" >
            <Language ID="MatchOS" />
        </Product>
    </Remove>
    <Display Level="NONE" AcceptEULA="TRUE" />
  </Configuration>
"@        
    }
    else
    {
      write-ezlogs " | Office Edition: $OfficeEdition" -showtime
      If($ExcludeApps){
        write-ezlogs " | Excluded Apps: $ExcludeApps" -showtime
        $ExcludeApps | ForEach-Object{
          $ExcludeAppsString += "<ExcludeApp ID =`"$_`" />"
        }
      }
      If($OfficeArch){
        write-ezlogs " | Office Architecture: $OfficeArch" -showtime
        $OfficeArchString = "`"$OfficeArch`""
      }
      
      if($OrgName)
      {
        write-ezlogs " | Organization Name: $OrgName" -showtime
        $OrgName_String = "<AppSettings>
        <Setup Name=`"Company`" Value=`"$OrgName`" /></AppSettings>"      
      }else{
        $OrgName_String = $null
      }
      write-ezlogs " | Remove MSI: $RemoveMSI" -showtime
      If($RemoveMSI){
        $RemoveMSIString = '<RemoveMSI />'
      }Else{
        $RemoveMSIString = $Null
      }
      write-ezlogs " | Office Channel: $Channel" -showtime
      If($Channel){
        $ChannelString = "Channel=`"$Channel`""
      }Else{
        $ChannelString = $Null
      }
      if($SourcePath -eq 'Internet')
      {
        $SourcePathString = $Null
        write-ezlogs ' | Source Path: Internet - Microsoft CDN' -showtime
      }
      else
      {
        $SourcePath_Valid = Test-Path $SourcePath -PathType Container -ErrorAction SilentlyContinue
        If($SourcePath_Valid){
          write-ezlogs " | Source Path: $SourcePath" -showtime
          $SourcePathString = "SourcePath=`"$SourcePath`"" 
        }
        Else
        {
          write-ezlogs ' | Source Path: Internet - Microsoft CDN' -showtime
          $SourcePathString = $Null
        }        
      }
      write-ezlogs " | Display Install: $DisplayInstall" -showtime
      If($DisplayInstall){
        $SilentInstallString = 'Full'
      }Else{
        $SilentInstallString = 'None'
      }
      write-ezlogs " | Auto Activate: $AUTOACTIVATE" -showtime
      If($AUTOACTIVATE){
        $AUTOACTIVATE = '1'
      }Else{
        $AUTOACTIVATE = '0'
      }
      write-ezlogs " | Accept EULA: $AcceptEULA" -showtime
      write-ezlogs " | Shared Computer Licensing (RDP installs): $SharedComputerlicensing" -showtime
      If($SharedComputerlicensing){
        $SharedComputerlicensing = '1'
      }Else{
        $SharedComputerlicensing = '0'
      }
      write-ezlogs " | ForceAppShutdown: $FORCEAPPSHUTDOWN" -showtime
      write-ezlogs " | PinIconsToTaskbar: $PinItemsToTaskbar" -showtime
      write-ezlogs " | Enable Updates: $EnableUpdates" -showtime
      #XML data that will be used for the download/install
      $OfficeXML = [XML]@"
  <Configuration>
    <Add OfficeClientEdition=$OfficeArchString $ChannelString $SourcePathString  >
      <Product ID="$OfficeEdition">
        <Language ID="MatchOS" />
        $ExcludeAppsString
      </Product>
    </Add>  
    <Property Name="PinIconsToTaskbar" Value="$PinItemsToTaskbar" />
    <Property Name="AUTOACTIVATE" Value="$AUTOACTIVATE" />
    <Property Name="FORCEAPPSHUTDOWN" Value="$FORCEAPPSHUTDOWN" />
    <Property Name="SharedComputerLicensing" Value="$SharedComputerlicensing" />
    <Display Level="$SilentInstallString" AcceptEULA="$AcceptEULA" />
    <Updates Enabled="$EnableUpdates" />
    $RemoveMSIString
     <AppSettings>
     $OrgName_String
     </AppSettings>
  </Configuration>
"@
    }
    #Save the XML file
    $OfficeXML.Save("$OfficeInstallDownloadPath\\OfficeInstall.xml")
    write-ezlogs ">> XML File Generated: $OfficeInstallDownloadPath\\OfficeInstall.xml" -showtime -color cyan
    Return "$OfficeInstallDownloadPath\\OfficeInstall.xml"
  }

  Function Get-ODTURL {
    $ODTDLLink = ''
    $MSWebPage = (Invoke-WebRequest 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117' -UseBasicParsing).Content
  
    #Thank you reddit user, u/sizzlr for this addition.
    Foreach ($m in $MSWebPage)
    {
      If($m -match 'url=(https://.*officedeploymenttool.*\\.exe)'){
      $ODTDLLink = $matches[1]}
    }
    if(Test-URL $ODTDLLink)
    {
      Return $ODTDLLink
    }
    else
    {
      Return $false
    }
  
  }

  $VerbosePreference = 'SilentlyContinue'
  $ErrorActionPreference = 'Stop'

  If(-Not(Test-Path $OfficeInstallDownloadPath )){
    $Null = New-Item -Path $OfficeInstallDownloadPath  -ItemType Directory -ErrorAction Stop
  }

  If(!($ConfigurationXMLFile))
  { 
    $ConfigurationXMLFile = Generate-XMLFile #If the user didn't specify with -ConfigurationXMLFile param, we make one
  }
  Else
  {
    if(Test-URL($ConfigurationXMLFile))
    {
      write-ezlogs 'XML file path provided is a web URL' -showtime
      $Download_FileName = '365ODT-Config.xml'
      $ConfigurationXMLFile = Invoke-FileDownload -DownloadURL $ConfigurationXMLFile -Download_Directory $OfficeInstallDownloadPath -Download_file_name $Download_FileName -Overwrite
    }    
    elseif(!(Test-Path $ConfigurationXMLFile))
    {
      Write-ezlogs 'The configuration XML file is not a valid file. Please check the path and try again' -showtime -Warning
      Stop-EZLogs -ErrorSummary -stoptimer -clearErrors -logOnly
      Exit
    }
    else
    {
      write-ezlogs "Using existing XML configuration file: $ConfigurationXMLFile" -showtime
    }
  }

  #Get the ODT Download link
  $ODTInstallLink = Get-ODTURL

  #Download the Office Deployment Tool
  write-ezlogs '>> Downloading the Office Deployment Toolkit' -showtime -color Cyan
  if($ODTInstallLink -ne $false)
  {
    write-ezlogs " | ODT Install Link: $ODTInstallLink" -showtime
    Try
    {
      #$null = Invoke-WebRequest -Uri $ODTInstallLink -OutFile "$OfficeInstallDownloadPath\\ODTSetup.exe" 
      $ODTInstallFile = Invoke-FileDownload -DownloadURL $ODTInstallLink -Destination_File_Path "$OfficeInstallDownloadPath\\ODTSetup.exe" -Overwrite
    }
    Catch
    {
      write-ezlogs "[ERROR] An exception occured downloading ODT from $ODTInstallLink :`n | $($_.exception.message)`n | $($_.InvocationInfo.positionmessage)`n | $($_.ScriptStackTrace)`n" -color Red -showtime
      Stop-Ezlogs -ErrorSummary -logOnly -stoptimer -clearErrors -enablelogs:([System.Convert]::ToBoolean($enablelogs))
      Exit
    }
  }
  else
  {
    write-ezlogs '[ERROR] Unable to get the ODT download link' -color Red -showtime
    Stop-Ezlogs -ErrorSummary -logOnly -stoptimer -clearErrors -enablelogs:([System.Convert]::ToBoolean($enablelogs))
    Exit
  }
  
  #Run the Office Deployment Tool setup
  Try
  {
    write-ezlogs ' | Running the Office Deployment Toolkit Setup' -showtime
    Start-Process $ODTInstallFile -ArgumentList "/quiet /extract:$OfficeInstallDownloadPath" -Wait
  }
  Catch
  {
    write-ezlogs "[ERROR] An exception occured running ODT:`n | $($_.exception.message)`n | $($_.InvocationInfo.positionmessage)`n | $($_.ScriptStackTrace)`n" -color Red -showtime
  }
  #Run the O365 install
  Try
  {
    write-ezlogs '>> Downloading and Running Office 365 Setup' -showtime -color Cyan
    $OfficeInstaller = "$OfficeInstallDownloadPath\\setup.exe"
    $OfficeArguments = "/configure `"$ConfigurationXMLFile`""
    write-ezlogs " | Office Setup Command: $OfficeInstaller $OfficeArguments" -showtime 
    $ODTLogFile = "$env:COMPUTERNAME-$(Get-Date -f 'yyyyMMdd')"
    $Proc = Start-Process -FilePath $OfficeInstaller -ArgumentList $OfficeArguments -Wait -PassThru -WindowStyle Hidden
    $proc | Wait-Process -Timeout 400 -ErrorAction Continue -ErrorVariable timeouted
    if ($timeouted)
    {
      # terminate the process
      $proc | Stop-Process 
      Write-ezlogs  "[ERROR] Process failed to finish before the timeout period and was canceled. Removing downloaded files and exiting. $timeouted" -Color red -showtime
      $Null = Remove-Item $OfficeInstallDownloadPath -Recurse -Force
      Stop-Ezlogs -ErrorSummary -logOnly -stoptimer -clearErrors -enablelogs:([System.Convert]::ToBoolean($enablelogs))
      exit
    }
    elseif ($proc.ExitCode -ne 0)
    {
      Write-ezlogs "Unexpected process exit code ($($proc.ExitCode)). Halting further actions" -showtime -Warning
      Stop-Ezlogs -ErrorSummary -logOnly -stoptimer -clearErrors -enablelogs:([System.Convert]::ToBoolean($enablelogs))
      exit
    }    
  }
  Catch
  {
    write-ezlogs "[ERROR] An exception occured while running Office 365 setup:`n | $($_.exception.message)`n | $($_.InvocationInfo.positionmessage)`n | $($_.ScriptStackTrace)`n" -color Red -showtime
  }
  
  $OfficeInstalled = $False
  $Office_Installs = Get-OfficeVersion -ShowAllInstalledProducts | select *
  Foreach ($Key in $Office_Installs ) 
  {
    If($Key.('KeyName') -like "*$OfficeEdition*" -or $Key.('DisplayName') -like '*Microsoft 365*') 
    {
      $OfficeVersionInstalled = "$($Key.('DisplayName')) - $($Key.('Version'))"
      $OfficeInstalled = $True
    }
  }

  If($OfficeInstalled -and !$RemoveOnly)
  {
    write-ezlogs "[SUCCESS] $($OfficeVersionInstalled) installed successfully" -showtime -color Green
    write-ezlogs $($Office_Installs | out-string)
  }
  elseif(!$RemoveOnly)
  {
    write-ezlogs 'Office 365 was not detected after the install ran. Check the log files to get more information' -showtime -warning
  }  
  if($RemoveOnly -and !$OfficeInstalled)
  {
    write-ezlogs "[SUCCESS] Office 365 ($($RemoveOnlyEdition)) was removed successfully" -showtime -color Green
  }
  elseif($RemoveOnly)
  {
    write-ezlogs "Office 365 ($($RemoveOnlyEdition)) was detected after removal ran. Check the log files to get more information" -showtime -warning
  }
  if($Copy_SetupLog)
  {
    $ODTLogFileFound = Get-ChildItem $env:temp -Recurse -Filter "$ODTLogFile*.log"  -ErrorAction SilentlyContinue
    if($ODTLogFileFound)
    {
      write-ezlogs "Copying Office 365 Setup log ($($ODTLogFileFound.FullName)) to directory ($Copy_SetupLog_directory)" -showtime
      foreach($log in $ODTLogFileFound){
        $null = Copy-Item $log -Destination $Copy_SetupLog_directory -Force
      }
    }
    else
    {
      write-ezlogs "Unable to find an Office 365 Setup log file matching $ODTLogFile" -showtime -Warning
    }
  }
  if($Remove_Install_Office_Download_Path)
  {
    write-ezlogs ">> Removing all files downloaded and created at $OfficeInstallDownloadPath" -showtime -color Cyan
    try
    {
      $null = Remove-item $OfficeInstallDownloadPath -Recurse -Force
    }
    catch
    {
      write-ezlogs "[ERROR] An exception occurred when removing $OfficeInstallDownloadPath`n | $($_.exception.message)`n | $($_.InvocationInfo.positionmessage)`n | $($_.ScriptStackTrace)`n" -color red -showtime
    }
  }
}
#---------------------------------------------- 
#endregion Install-Office365Suite Function
#----------------------------------------------

#############################################################################
#endregion Core functions
#############################################################################


#############################################################################
#region Execution and Output - Functions or Code that executes required actions and/or performs output 
#############################################################################

#---------------------------------------------- 
#region Execute Get-OfficeVersion
#----------------------------------------------
Use-RunAs
$thisScript = Get-ThisScriptInfo
$logfile = Start-Ezlogs -Logfile_Directory:$logfile_directory -Start_Timer -thisScript $thisScript
Load-Modules -modules $Required_modules -force:$force_modules -update:$update_modules -enablelogs:([System.Convert]::ToBoolean($enablelogs))
$OfficeInstalled = $False
write-ezlogs '#### Checking Office Installations ####' -linesbefore 1 -color yellow
$Office_Installs = Get-OfficeVersion -ShowAllInstalledProducts | select ComputerName,DisplayName,KeyName,Version,Build,Bitness,ClientCulture,ClickToRun,ClickToRunUpdatesEnabled,ClickToRunUpdateUrl,Activate,DeviceBasedLicensing,SharedComputerLicensing,InstallPath,TeamsAddOn,ExcludedApps

if($Office_Installs)
{
  write-ezlogs "---- Office Installation(s) found ----`n$($($Office_Installs | out-string).trim())`n"
  Foreach ($Key in $Office_Installs ) 
  {
    If($Install_Office_Edition -and $Key.('KeyName') -like "*$Install_Office_Edition*") 
    {
      $OfficeVersionInstalled = "Name: $($Key.('DisplayName')) - Version: $($Key.('Version')) - Edition: $($Key.('KeyName'))"
      $OfficeInstalled = $True
    }  
  } 
}
else
{
  write-ezlogs 'No valid Office installations were found' -showtime
}
#---------------------------------------------- 
#endregion Execute Get-OfficeVersion
#----------------------------------------------

#---------------------------------------------- 
#region Execute Install-Office365Suite
#----------------------------------------------
if(([System.Convert]::ToBoolean($Remove_Office_Only)))
{
  $RemoveOfficeInstalled = $false
  if(!$Remove_Office_Only_Edition)
  {
    $Remove_Office_Only_Edition = 'All'
  }
  Foreach ($Key in $Office_Installs ) 
  {
    If($Remove_Office_Only_Edition -and $Key.('KeyName') -like "*$Remove_Office_Only_Edition*") 
    {
      $OfficeVersionInstalled = "Name: $($Key.('DisplayName')) - Version: $($Key.('Version')) - Edition: $($Key.('KeyName'))"
      $RemoveOfficeInstalled = $True
    }
    elseif($Office_Installs -and $Remove_Office_Only_Edition -eq 'All')
    {
      $RemoveOfficeInstalled = $true
    }
    else
    {
      $RemoveOfficeInstalled = $false
    }  
  }  
  if($RemoveOfficeInstalled)
  {
    write-ezlogs " | Office Edition(s) to remove: $Remove_Office_Only_Edition" -showtime
    Invoke-Office365Setup -RemoveOnly:([System.Convert]::ToBoolean($Remove_Office_Only)) -RemoveOnlyEdition:$Remove_Office_Only_Edition -Copy_SetupLog:([System.Convert]::ToBoolean($Copy_SetupLog)) -OfficeInstallDownloadPath:$Install_Office_Download_Path
  }
  else
  {
    write-ezlogs 'The edition of Office specified to remove was not found' -showtime -warning
  } 
}
elseif(([System.Convert]::ToBoolean($Install_Office)))
{
  if ($OfficeInstalled)
  {
    write-ezlogs 'The Edition of Office specified to install is already installed on this machine | Skipping any further actions' -showtime -warning
    break
  }
  else
  {
    if (([System.Convert]::ToBoolean($Install_Office_WITH_ODT)) -and $Install_Office_WITH_ODT_XMLFILE)
    {
      Invoke-Office365Setup -ConfigurationXMLFile $Install_Office_WITH_ODT_XMLFILE -OfficeInstallDownloadPath:$Install_Office_Download_Path -Copy_SetupLog:([System.Convert]::ToBoolean($Copy_SetupLog)) -Remove_Install_Office_Download_Path:([System.Convert]::ToBoolean($Remove_Install_Office_Download_Path))
    }
    else
    {
      Invoke-Office365Setup -AcceptEULA:([System.Convert]::ToBoolean($Install_Office_Accept_EULA)) -SharedComputerLicensing:([System.Convert]::ToBoolean($Install_Office_Shared_Computer_Licensing)) -FORCEAPPSHUTDOWN:([System.Convert]::ToBoolean($Install_Office_FORCE_APPSHUTDOWN)) -ExcludeApps $Install_Office_Exclude_Apps -OrgName $Install_Office_Org_Name -AUTOACTIVATE:([System.Convert]::ToBoolean($Install_Office_AUTO_ACTIVATE)) -EnableUpdates:([System.Convert]::ToBoolean($Install_Office_Enable_Updates)) -OfficeEdition $Install_Office_Edition -RemoveMSI:([System.Convert]::ToBoolean($Install_Office_Remove_Previous_Intalls)) -DisplayInstall:([System.Convert]::ToBoolean($Install_Office_Display_Install)) -Copy_SetupLog:([System.Convert]::ToBoolean($Copy_SetupLog)) -OfficeInstallDownloadPath:$Install_Office_Download_Path -SourcePath:$Install_Office_Source_Path -Channel:$Install_Office_Channel
    }
  }
}
#---------------------------------------------- 
#endregion Execute Install-Office365Suite
#----------------------------------------------

#---------------------------------------------- 
#region End Logging
#----------------------------------------------
Stop-EZLogs -ErrorSummary $error -logOnly -stoptimer -clearErrors -enablelogs:([System.Convert]::ToBoolean($enablelogs))
#---------------------------------------------- 
#endregion End Logging
#----------------------------------------------
#############################################################################
#endregion Execution and Output Functions
#############################################################################
'''

print ("iTarian RMM - Executing Powershell Script")

def ecmd(command):
    import ctypes
    from subprocess import PIPE, Popen
    
    class disable_file_system_redirection:
        _disable = ctypes.windll.kernel32.Wow64DisableWow64FsRedirection
        _revert = ctypes.windll.kernel32.Wow64RevertWow64FsRedirection
        def __enter__(self):
            self.old_value = ctypes.c_long()
            self.success = self._disable(ctypes.byref(self.old_value))
        def __exit__(self, type, value, traceback):
            if self.success:
                self._revert(self.old_value)
    
    with disable_file_system_redirection():
        obj = Popen(command, shell = True, stdout = PIPE, stderr = PIPE)
    out, err = obj.communicate()
    ret=obj.returncode
    if ret==0:
        if out:
			return out.strip()
        else:
            return ret
    else:
        if err:
            return err.strip()
        else:
            return ret

file_name='EZT-DeployO365.ps1'
file_path=os.path.join(os.environ['TEMP'], file_name)
with open(file_path, 'wb') as wr:
    wr.write(ps_content)

ecmd('powershell "Set-ExecutionPolicy Bypass"')
print ecmd('powershell "%s"'%file_path)

os.remove(file_path)