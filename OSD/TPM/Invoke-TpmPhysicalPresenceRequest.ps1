<#
    .SYNOPSIS
        Requests a TPM operation that requires physical presence.
    
    .DESCRIPTION
        This script attempts to request a TPM operation, such as enabling, disabling, activating, deactivating, 
        or clearing. If no TPM operation is specified, the script will attempt to enable, activate, and clear
        the TPM. More information about the various TPM operations can be found at:
        https://msdn.microsoft.com/en-us/library/windows/desktop/aa376478(v=vs.85).aspx
    
    .PARAMETER RequestedTpmOperation
        Specify the requested TPM operation that requires physical presence. 

    .EXAMPLE
        # Enable + Activate + Clear the TPM
        .\Invoke-TpmPhysicalPresenceRequest.ps1 -RequestedTpmOperation EnableAndActivateAndClear
    
    .NOTES
        FileName:   Invoke-TpmPhysicalPresenceRequest.ps1
        Author:     Jacob Thornberry
        Contact:    @altrhombus
        Created:    2018-04-20
        Updated:    2018-06-01

        Version history:
        1.0.0 - (2018-04-20) Script created
        2.0.0 - (2018-05-31) Major rewrite (/salute) for code clarity
        2.0.1 - (2018-06-01) Enhance logging capability
#>
[CmdletBinding(SupportsShouldProcess = $true)]
Param(
    [Parameter(mandatory=$False,HelpMessage="The TPM operation requested.")]
    [ValidateSet('Nothing','Enable','Disable','Activate','Deactivate','Clear','EnableAndActivate','DeactivateAndDisable','AllowTPMOwnerInstallation','PreventTPMOwnerInstallation','EnableAndActivateAndAllowTPMOwnerInstallation','DeactivateAndDisableAndPreventTPMOwnerInstallation','ClearAndEnableAndActivate','EnableAndActivateAndClear','EnableAndActivateAndClearAndEnableAndActivate')]
    [String]$RequestedTpmOperation
)

Begin {
    # Load Microsoft.SMS.TSEnvironment COM object
	try {
		$TSEnvironment = New-Object -ComObject Microsoft.SMS.TSEnvironment -ErrorAction Stop
	}
	catch [System.Exception] {
		Write-Warning -Message "Unable to construct Microsoft.SMS.TSEnvironment object"; exit 1
	}
}

Process {
    switch ($RequestedTpmOperation) {
        "Nothing" { $tpmOperation = "0" }
        "Enable" { $tpmOperation = "1" }
        "Disable" { $tpmOperation = "2" }
        "Activate"  { $tpmOperation = "3" }
        "Deactivate" { $tpmOperation = "4" }
        "Clear" { $tpmOperation = "5" }
        "EnableAndActivate" { $tpmOperation = "6" }
        "DeactivateAndDisable" { $tpmOperation = "7" }
        "AllowTPMOwnerInstallation" { $tpmOperation = "8" }
        "PreventTPMOwnerInstallation" { $tpmOperation = "9" }
        "EnableAndActivateAndAllowTPMOwnerInstallation" { $tpmOperation = "10" }
        "DeactivateAndDisableAndPreventTPMOwnerInstallation" { $tpmOperation = "11" }
        "ClearAndEnableAndActivate" { $tpmOperation = "14" }
        "EnableAndActivateAndClear" { $tpmOperation = "21" }
        "EnableAndActivateAndClearAndEnableAndActivate" { $tpmOperation = "22" }
        Default { $tpmOperation = "21" }
    }

    function Write-CMLogEntry {
        param (
            [parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
            [ValidateNotNullOrEmpty()]
            [string]$Value,
            [parameter(Mandatory = $true, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
            [ValidateNotNullOrEmpty()]
            [ValidateSet("1", "2", "3")]
            [string]$Severity,
            [parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
            [ValidateNotNullOrEmpty()]
            [string]$FileName = "TpmPhysicalPresenceRequest.log"
        )
        # Determine log file location
        $LogFilePath = Join-Path -Path $LogsDirectory -ChildPath $FileName
        
        # Construct time stamp for log entry
        $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), "+", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
        
        # Construct date for log entry
        $Date = (Get-Date -Format "yyyy-MM-dd")
        
        # Construct context for log entry
        $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
        
        # Construct final log entry
        $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""TpmPhysicalPresenceRequest"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
        
        # Add value to log file
        try {
            Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
        }
        catch [System.Exception] {
            Write-Warning -Message "Unable to append log entry to TpmPhysicalPresenceRequest.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
        }
    }

    function Set-TPMPhysicalPresence {
        param (
            [parameter(Mandatory = $true, HelpMessage = "Integer value that specifies the requested TPM operation that requires physical presence.")]
            [ValidateNotNullOrEmpty()]
            [ValidateSet("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "14", "21", "22")]
            [string]$tpmOperationValue
        )

        try {
            # Try to connect to the Win32_Tpm namespace
            $tpm = Get-WmiObject -Namespace "root\cimv2\Security\MicrosoftTpm" -Class Win32_Tpm

            # Write out some information about the TPM
            Write-CMLogEntry -Value "Found a $($tpm.ManufacturerIdTxt) TPM supporting Specification Versions $($tpm.SpecVersion)." -Severity 1
            Write-CMLogEntry -Value "TPM Activation status is $($tpm.IsActivated_InitialValue)" -Severity 1
            Write-CMLogEntry -Value "TPM Enabled status is $($tpm.IsEnabled_InitialValue)" -Severity 1
            Write-CMLogEntry -Value "TPM Ownership status is $($tpm.IsOwned_InitialValue)" -Severity 1

            # Request the new physical presence
            $tpmRequestProcess = $tpm.SetPhysicalPresenceRequest($tpmOperationValue)
            Write-CMLogEntry -Value "Completed the requested TPM operation with error code $($tpmRequestProcess.ReturnValue)" -Severity 1
        }
        catch {
            Write-CMLogEntry -Value "Failed to complete the requested TPM operation with error code $($tpmRequestProcess.ReturnValue)" -Severity 3
        }
    }

    $LogsDirectory = $Script:TSEnvironment.Value("_SMSTSLogPath")

    Write-CMLogEntry -Value "TPM Physical Presence Request Version 2.0.1" -Severity 1

    Set-TPMPhysicalPresence($tpmOperation)

}