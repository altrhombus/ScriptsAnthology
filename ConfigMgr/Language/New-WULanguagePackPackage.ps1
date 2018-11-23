<#
.SYNOPSIS
    Create Language Pack packages in ConfigMgr.
.DESCRIPTION
    This script will download Language Packs from Windows Update and create packages for them in ConfigMgr.
.PARAMETER SiteServer
    Site server where the SMS Provider is installed.
.PARAMETER PackageSourcePath
    Root path where Language Pack package source files will be stored.
.PARAMETER XmlSourcePath
    Path where the XML file containing download URLs is stored.
.EXAMPLE
    .\New-WULanguagePackPackage.ps1 -SiteServer "CM01" -PackageSourcePath "\\CM01\Sources\OSD\LanguagePacks\Windows10" -XmlSourcePath ".\lp_download.xml"
.NOTES
    FileName:    New-CMLanguagePackPackage.ps1
    Author:      Jacob Thornberry
    Contact:     @altrhombus
    Created:     2018-11-20
    Updated:     2018-11-20
    
    Version history:
    1.0.0 - (2018-11-20) Script created
    1.0.1 - (2018-11-23) Add ConfigMgr Package creation
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [parameter(Mandatory=$true, HelpMessage="Site server where the SMS Provider is installed.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({Test-Connection -ComputerName $_ -Count 1 -Quiet})]
    [string]$SiteServer,

    [parameter(Mandatory=$true, HelpMessage="Root path where Language Pack package source files will be stored.")]
    [ValidateNotNullOrEmpty()]
    [string]$PackageSourcePath,

    [parameter(Mandatory=$true, HelpMessage="Full path of the XML file.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({Test-Path $_ })]
    [string]$XmlSourcePath
)

Begin {
    # Determine SiteCode from WMI
    try {
        Write-Verbose -Message "Determining Site Code for Site server: '$($SiteServer)'"
        $SiteCodeObjects = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $SiteServer -ErrorAction Stop
        foreach ($SiteCodeObject in $SiteCodeObjects) {
            if ($SiteCodeObject.ProviderForLocalSite -eq $true) {
                $SiteCode = $SiteCodeObject.SiteCode
                Write-Verbose -Message "Site Code: $($SiteCode)"
            }
        }
    }
    catch [System.UnauthorizedAccessException] {
        Write-Warning -Message "Access denied" ; break
    }
    catch [System.Exception] {
        Write-Warning -Message "Unable to determine Site Code" ; break
    }

    # Load ConfigMgr module
    try {
        $SiteDrive = $SiteCode + ":"
        Import-Module -Name (Join-Path -Path (($env:SMS_ADMIN_UI_PATH).Substring(0, $env:SMS_ADMIN_UI_PATH.Length-5)) -ChildPath "\ConfigurationManager.psd1") -Force -ErrorAction Stop -Verbose:$false
        if ((Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue | Measure-Object).Count -ne 1) {
            New-PSDrive -Name $SiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root $SiteServer -ErrorAction Stop -Verbose:$false | Out-Null
        }
    }
    catch [System.UnauthorizedAccessException] {
        Write-Warning -Message "Access denied" ; break
    }
    catch {
        Write-Warning -Message "$($_.Exception.Message). Line: $($_.InvocationInfo.ScriptLineNumber)" ; break
    }

    # Determine current location
    $CurrentLocation = $PSScriptRoot

    # Disable Fast parameter usage check for Lazy properties
    $CMPSSuppressFastNotUsedCheck = $true
}
Process {
    [xml]$source = Get-Content $XmlSourcePath
    $totalSource = $source.LanguagePacks.LanguagePack.Count
    
    $i = 0
    foreach ($lp in $source.LanguagePacks.LanguagePack)
    {
        try {
            $subFolderPath = Join-Path -Path $PackageSourcePath -ChildPath ($lp.FriendlyBuild + "\" + $lp.Architecture + "\" + $lp.Locale) 
            if (!(Test-Path -Path $subFolderPath)) {
                Write-Verbose -Message "Creating folder: $($subFolderPath)"
                New-Item -Path $subFolderPath -ItemType Directory -Force -ErrorAction Stop -Verbose:$false | Out-Null
            }
        }
        catch {
            Write-Warning -Message "Unable to create subfolders. Error message: $($_.Exception.Message)" ; break
        }

        Write-Progress -Activity "Creating Language Packs from Windows Update" -Status "Processing $($lp.Locale) language pack for $($lp.Product) $($lp.FriendlyBuild) ($($lp.Architecture))" -PercentComplete (($i / $totalSource) * 100)
        
        try {
            Start-BitsTransfer -Destination $subFolderPath -Source $lp.DownloadPath -Description "Downloading $($lp.DownloadPath)" -ErrorAction Continue
        }
        catch {
            Write-Warning -Message "$($_.Exception.Message). Line: $($_.InvocationInfo.ScriptLineNumber)" ; break
        }

        try {
            Set-Location -Path $SiteDrive -ErrorAction Stop -Verbose:$false

            $LanguagePackageName = -join@("Language Pack - ", $lp.Product, " ", $lp.FriendlyBuild, " ", $lp.Architecture)
            Write-Verbose -Message "Creating Language Pack package: $($LanguagePackageName)"
            $LanguagePackPackage = New-CMPackage -Name $LanguagePackageName -Language $lp.Locale -Version $lp.Build -Path $subFolderPath -ErrorAction Stop -Verbose:$false

            Set-Location -Path $CurrentLocation
        }
        catch {
            Write-Warning -Message "Unable to create Language Pack package. Error message: $($_.Exception.Message)" ; break
        }
        $i++
    }
}