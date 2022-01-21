# EZT-DeployO365

**Synopsis**

Automates silent install or removal of Office 365 editions using the Office Deployment Tookit, allowing custom config XML generation,removal of existing/older Office installs and more
* * * 
**Features**

- Automatically downloads required ODT and Office 365 setup files
- Silently deploys Office 365 installs using ODT
- Detects all current installed Office products and editions
- Ability to create ODT XML file programmatically from configured variables
- Ability to remove/upgrade previous Office installs and other options available with ODT
- Ability to remove all or specified editions of Office 365 that are already installed
- Ability to provide existing custom ODT XML from local path, direct download web URL, or Onedrive public access share link
- Ability to automatically copy Office 365 setup logs to specified log directory (revives the LoggingPath option that was deprecated in ODT)
- Ability to specify an existing Office installation share to use for deployment or download from Microsoft CDN
- Detailed log output and error handling for troubleshooting problematic deployments


**Notes/Requirements**

- The script should be **run as the LocalSystem User** as it requires local admin privileges. It will work under logged in user if said user is a local admin. 
- This is a **powershell script**. As such this will only work on **Windows** endpoints. It is untested on Powershell Core. 
- Requires **Powershell v3** or higher. 
- Tested on **Windows 10**. Not compatible with Windows 7 as Office 365 is no longer supported on it

## Available Versions

**[Self Executable TEST]()**  

- Use this if you are a QA tester for this project or wish to test on a single machine
- Its packaged as a self-executable for quick and easy testing and mobility
- Configuration variables are set via a UI pop-up window that opens upon launching

**[Powershell Source Code]()** 

- Powershell only version, main script

**[Python Source Code]()**

- Python only version that is also the source code for the iTarian Procedure

**[iTarian Procedure]()**

- Encoded Python version exported from iTarian. Use this to quickly import into your iTarian procedures

## Installation and Configuration

### Installation for iTarian Procedures

1. **Download the ITSM procedure** 
2. Within your ITSM portal, import the procedure under **Configuration Templates - Procedures**
3. Configure desired **procedure name, alert settings**..etc
4. Configure the **default parameters** for the procedure from the **Parameters tab** of the script. See **Configuration Parameters** below for explanations of each parameter
5. Click **Save - Ready to Review - Approve** to finish. **Assign to a profile** and optionally a schedule of your liking
6. **(Recommended)** Run the new procedure on a single **test machine** to ensure its working or configured as expected

#### iTarian Configuration

- This script can be configured by editing the **parameter options** within the iTarian RMM procedure 

#### Powershell Configuration

- If you wish to use the pure Powershell shell script version, use the configuration variables located in the region **Configurable Script Parameters** located near the top of the script 

### Configuration Parameters/Variables

Be sure to read and understand what each config option does, as some are dependant on others, or some negate others..etc

_**Note: 1 = Enabled, 0 = Disabled**_


#### Install Office Configuration

-  **InstallOffice**
   - Default: 0
   - Installs Office 365 using config values set in this script. 
   - If this option and RemoveOffice_Only are disabled, script only returns any detected Office installations  

-  **InstallOffice_Download_Path**
   - Default: C:\Office365Install
   - Directory where the Office 365 setup and ODT files should be downloaded.
   
-  **InstallOffice_Edition**
   - Default: O365BusinessRetail
   - Available Options: O365BusinessRetail, O365ProPlusRetail (Pick 1 Only)
   - Edition of Office 365 that will be installed and/or verified

-  **Install_Office_Source_Path**
   - Default: Internet
   - Path where Office installation files will be deployed from
   - Defaults to downloading from the internet (Microsoft CDN) if a valid path is not provided
   - Good to use if deploying to many clients via a central network share to prevent internet bandwidth saturation

-  **InstallOffice_WITH_ODT**
   - Default: 0
   - Enables using an existing ODT XML configuration file for Office installs
   - If enabled, "InstallOffice_WITH_ODT_XMLFILE" must be configured
   - If enabled, any further "InstallOffice_" options are ignored as they will be provided by the XML config

-  **InstallOffice_WITH_ODT_XMLFILE**
   - Default: C:\Office365Install\365ODT-Config.xml
   - Path to an existing ODT XML configuration file to use for Office Installs
   - Requires enabling "InstallOffice_WITH_ODT", otherwise is ignored
   - Local/UNC paths, direct download web URLs, and Onedrive public access share link URLs are accepted
   - If provided a URL to a valid XML file, it will be downloaded automatically to the directory specified in "InstallOffice_Download_Path"
   - Onedrive share links permissions must be public access links (anyone who has the link is allowed access)
     - Recommend creating links that expire as soon as you are done with deployment
   - To create and customize your own configuration file, highly recommend visiting https://config.office.com 

-  **Install_Office_Channel**
   - Default: Current
   - Available Options: Current,CurrentPreview,SemiAnnual,SemiAnnualPreview,BetaChannel,MonthlyEnterprise (Pick 1 Only)
   - Defines which channel to use for installing Office
   
-  **InstallOffice_Accept_EULA**
   - Default: 1
   - Enables automatically accepting EULA for Office 365 installs

-  **InstallOffice_Auto_Activate**
   - Default: 1
   - Enables automatic activation of Office Installs

-  **InstallOffice_Display_Install**
   - Default: 0
   - Enables displaying UI while Installing Office. Leave disabled for silent installs

-  **InstallOffice_Enable_Updates**
   - Default: 1
   - Enables automatic updates for Office installs

-  **InstallOffice_Exclude_Apps**
   - Default: Groove,Lync
   - Available Options: Groove,Teams,Lync,Access,Excel,Onedrive,Onenote,Outlook,PowerPoint,Publisher,Word
   - Specifies the applications that should be excluded from installing with Office, comma seperated

-  **InstallOffice_Force_APPSHUTDOWN**
   - Default: 1
   - Enables force closing any open Office apps or processes when installing Office

-  **InstallOffice_Org_Name**
   - Default: My Organization
   - Specifies your organization name for Office installs

-  **InstallOffice_Remove_Previous_Installs**
   - Default: 1
   - Enables removal of any previous non-365 (MSI installed) versions of Office found when installing Office 365

-  **InstallOffice_Shared_Computer_Licensing**
   - Default: 0
   - Enables configuring office to use Shared Computer Licensing during Office installs
   - This is primarily used for installing Office on shared systems such as RDP servers. Note: this requires a special license from Microsoft to use

#### Remove Office Configuration

-  **RemoveOffice_Only**
   - Default: Disabled
   - Removes/uninstalls existing Office 365 installs found
   - If enabled, no installs or other options configured will occur. Will only remove Office 365 installations. Will not remove MSI based or older Office installations

-  **RemoveOffice_Only_Edition**
   - Default: All
   - Specifies the edition of Office to remove when RemoveOffice_Only is enabled
   - Available Options: O365BusinessRetail, O365ProPlusRetail, All (Pick 1)
   - Requires enabling "RemoveOffice_Only"
   - All (default) removes any 365 edition of Office (Including Visio, Project..etc). Will not remove non-365 licensed editions (OEM/VL editions)

#### Log Configuration

-  **LogFile_Directory** 
   - Default: "C:\Logs\"
   - Directory where log file should be created
   - Log file contains detailed status, error, warning and other messages used for troubleshooting or auditing
   
-  **Copy_SetupLogs**
   - Default: 1
   - Copies Office 365 setup logs to the log directory specified in "LogFile_Directory"
   - This provides similiar functionality to the LoggingPath ODT XML option that was deprecated by Microsoft
