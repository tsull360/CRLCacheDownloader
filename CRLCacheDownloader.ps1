<#
.SYNOPSIS
    Automate the downloading of DISA CRL files.

.DESCRIPTION
    CRLCacheDownloader.ps1 is a small PowerShell script used to auomate the retrieval of 
    Certificate Revocation Lists (CRL) from DISA.

.PARAMETER CRLSource
    HTTP url of CRL sources.

.PARAMETER CRLPath
    Local location CRL files will be saved to.

.PARAMETER TempPath
    Temporary script working location.
.PARAMETER ProxyConfig
    If required, specify URL of local network proxy.
    
.PARAMETER EVLog
    Enables the logging of script actions to the local event viewer
    
.PARAMETER Email
    Enables the sending of an email message containing script actions.
    
.PARAMETER SMTPServer
    Mail server to send email message to.
    
.PARAMETER MailTo
    Recipient of email message.
    
.PARAMETER MailFrom
    Sender of email status message.
    
.NOTES
    Author: Tim Sullivan
    Version: 1.0
    Contact: tsull360@live.com
    Date: 14/11/2018
    Name: CRLCacheDownloader.ps1

    CHANGE LOG
    Version 1.0: Initial Release

.EXAMPLE
    .\CRLCacheDownloader.ps1
#>

#Parameters used throughout the script are defined below. Some are mandatory, others are optional.
[CmdletBinding(SupportsShouldProcess=$true)]
param
(
#This is the URL of the CRL file itself.
[Parameter(Mandatory=$false)]
[String]$CRLSource = "http://crl.disa.mil/getcrlzip?ALL+CRL+ZIP",

#This is the location on the network where the extracted CRL's should be saved to.
[Parameter(Mandatory=$false)]
[String]$CRLPath = "C:\inetpub\wwwroot\crl",

#This is the temp location where the download file is saved to.
[Parameter(Mandatory=$false)]
[String]$TempPath = "C:\Working",

#If a proxy is used on the network, the URL specified here will be used for communications.
[Parameter(Mandatory=$false)]
[String]$ProxyConfig,

#Set this to $true to write to the event log, or $false to skip.
[Parameter(Mandatory=$false)]
[Boolean]$EVLog = $true,

#Set this to $true to send an email, or $false to skip.
[Parameter(Mandatory=$false)]
[Boolean]$Email = $false,

#Specify the address of the mail server to send messages to.
[Parameter(Mandatory=$false)]
[String]$SMTPServer,

#Enter the email address of the person who will recieve the emails.
[Parameter(Mandatory=$false)]
[String]$MailTo,

#enter the email address that should appear as the sender of the message.
[Parameter(Mandatory=$false)]
[String]$MailFrom

)

#We can now begin our script. The BEGIN section contains prepatory tasks used laster
#in the script.
BEGIN
{
#A hashtable is used to hold various messages generated during the script run.
    [HashTable]$StatusHash =@{
    "UnzipStatus"="Not Done";
    "DownloadStatus"="Not Done";
    "CRLPath"="Unknown";
    "TempPath"="Unknown"
    }

    #This function uses our parameters to pull down the CRL file from the internet.
    Function Get-CRL
    {
        Write-Verbose "In Get-CRL Function..."
        Write-Verbose "Attempting to download from $CRLSource and save the contents to $CRLPath"
        If($ProxyConfig -notlike "")
        {
            Write-Verbose "Web Proxy parameter passed. Will use this to download."
            $Proxy = new-object System.Net.WebProxy($ProxyConfig,$true)
            $WebProxy.Proxy = $Proxy
            try
            {
                Write-Verbose "Downloading file via proxy.."
                $WebProxy.DownloadFiles($CRLSource, $TempPath+"\ALLCRLZIP.zip")
                Write-Verbose "File downloaded!"
                $StatusHash.Set_Item("DownloadStatus","Success")
            }
            catch
            {
                Write-Verbose "Error downloading file via proxy. Error: $_.Exception.Message"
                $StatusHash.Set_Item("DownloadStatus","Error: $_.Exception.Message")
            }
        }
        Else
        {
            Write-Verbose "No proxy specified. Will direct download."
            $WebClient = New-Object System.Net.WebClient
            try
            {
                Write-Verbose "Downloading file directly..."
                $WebClient.DownloadFile($CRLSource, $TempPath+"\ALLCRLZIP.zip")
                Write-Verbose "File downloaded!"
                $StatusHash.Set_Item("DownloadStatus","Success")
            }
            catch
            {
                Write-Verbose "Error downloading file directly. Error: $_.Exception.Message"
                $StatusHash.Set_Item("DownloadStatus","Error: $_.Exception.Message")
            }
        }
    }

    #This function uses a comobject to extract the downloaded ZIP file containing the CRL's.
    Function CRL-Extract
    {
        $Shell_App=New-Object -ComObject Shell.Application
        $ExtractFile = $Shell_App.Namespace($TempPath+"\ALLCRLZIP.zip")
        $ExtractPath = $Shell_App.Namespace($CRLPath)
        try
        {
            Write-Verbose "Extracting file..."
            $ExtractPath.Copyhere($ExtractFile.items(), 0x14)
            Write-Verbose "File extracted!"
            $StatusHash.Set_Item("UnzipStatus","Success")
        }
        catch
        {
            Write-Verbose "Error extracting file! Error: $_.Exception.Message"
            $StatusHash.Set_Item("UnzipStatus","Error: $_.Exception.Message")
        }
    }

    #Here we test to make sure our CRL path exists. If not, we create it.
    If(!(Test-Path $CRlPath))
    {
        Write-Verbose "Local CRL path not found, creating."
        New-Item -ItemType Directory $CRlPath
        $StatusHash.Set_Item("CRLPath","Good")
    }
    Else
    {
        $StatusHash.Set_Item("CRLPath","Good")
    }

    #Here we test to make sure the specified temp path exists. If not, we create it.
    If (!(Test-Path $TempPath))
    {
        Write-Verbose "Local temp path not found, creating."
        New-Item -ItemType Directory $TempPath
        $Statushash.Set_Item("TempPath","Good")
    }
    Else
    {
        $Statushash.Set_Item("TempPath","Good")
    }

    If (Test-Path $TempPath+"\ALLCRLZIP.zip")
    {
        Remove-Item $TempPath+"\ALLCRLZIP.zip"
    }
}

#The Process block is where our work takes place.
PROCESS
{
Write-Verbose "CRL Cache Downloader Script"
Write-Verbose ""
Write-Verbose "Beginning work..."
Write-Verbose "Attempting to download CRL files..."
Get-CRL
Write-Verbose "CRL download complete."
Write-Verbose "Attempting to extract..."
CRL-Extract
Write-Verbose "Extraction task complete."
}

#Finally the ed block is where post processing and cleanup takes place.
END
{
Write-Verbose "Cleaning up temp files..."
If (Test-Path $TempPath+"\ALLCRLZIP.zip")
{
    Remove-Item $TempPath+"\ALLCRLZIP.zip"
}
Write-Verbose "Temp files cleaned up."

$LogMSG = @"
CRL Cache Downloader Script

Variables Provided
CRL Source: $CRLSource
CRL Path: $CRLPath
Temp Path: $TempPath
Proxy Configuration: $ProxyConfig
EVLog: $EVLog
Email: $Email
Mail Server: $SMTPServer
Mail Recipient: $MailTo
Mail Sender: $MailFrom

Task Status
Local CRL Path Status: $($StatusHash.Get_Item("CRLPath"))
Local Temp Path Status: $($StatusHash.Get_Item("TempPath"))
CRL Download Status: $($StatusHash.Get_Item("DownloadStatus"))
File Unzip Status: $($StatusHash.Get_Item("UnzipStatus"))

CRL Download Status Complete
"@

If ($EVLog -eq $true)
{
    Write-Verbose "Writing event log entry..."
    #Create event log source, if it does not already exist.
    if ([System.Diagnostics.EventLog]::SourceExists("CRLDownload") -eq $false) 
    {
        [System.Diagnostics.EventLog]::CreateEventSource("CRLDownload","Application")
    }
    Write-EventLog -LogName "Application" -EntryType Information -EventId 0815 -Source CRLDownload -Message $LogMSG
    Write-Verbose "Event log entry recorded."
}

            
If ($email -eq $true)
{
    Write-Verbose "Sending mail message..."
    try
    {
        Send-MailMessage -To $MailTo -Subject "CRL Cache Downloader" -Body $LogMsg -SmtpServer $SMTPServer -From $MailFrom
    }
    catch
    {
        Write-Verbose "Error sending mail message. Error: $_.Exception.Message"
    }
    Write-Verbose "Mail message sent."
}

Write-Verbose ""
Write-Verbose "Script Complete!"
}