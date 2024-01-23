<#
    .SYNOPSIS
        Install PowerShell on a remote machine.

    .DESCRIPTION
        This script will install PowerShell on many remote machines.

        First download the latest PowerShell .MSI package from the Microsoft
        website for use with this script.

        The script follows these steps for each computer:
        1. Copy the .MSI file to the remote computer
        2. Install the .MSI file on the remote computer
        3. Test the installation with a PowerShell remoting connection
           to the remote computer

        When the .MSI package fails to install, a log file is created
        that contains the computer names that failed.

    .PARAMETER ImportFile
        A .CSV file containing the 'ComputerName' property with all the computer
        names where the .MSI package needs to be installed.

    .PARAMETER PackagePath
        Path to the PowerShell .MSI file that will be installed on the remote
        computers.

    .PARAMETER DestinationFolder
        The destination folder where the .MSI package will be stored on the
        remote computer.

    .PARAMETER PowerShellEndpointVersion
        The PowerShell version to test the remote connection.

    .PARAMETER FailedInstallLogFile
        The folder where the log file will be stored that contains the computer
        names where the installation failed.

    .LINK
        https://4sysops.com/archives/how-to-install-and-upgrade-to-powershell-71/#rtoc-4

    .EXAMPLE
        msiexec.exe /package "c:\Temp\PowerShell-7.1.0-win-x64.msi" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1

        Local installation

    .EXAMPLE
        Start-Process 'msiexec.exe' -ArgumentList '/package "c:\Temp\PowerShell-7.1.0-win-x64.msi" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1' -Wait

        Install remotely
#>

Param (
    [String]$ImportFile = 'T:\Test\Brecht\PowerShell\PowerShell.7 not compatible clients 20240123.csv',
    [String]$PackagePath = 'C:\Users\bgijbels\Downloads\PowerShell-7.4.1-win-x64.msi',
    [String]$DestinationFolder = 'c:\Temp',
    [String]$PowerShellEndpointVersion = 'PowerShell.7',
    [String]$FailedInstallLogFile = 'T:\Test\Brecht\PowerShell\FailedInstalls.csv'
)

Begin {
    $VerbosePreference = 'Continue'

    #region Test
    if (-not (Test-Path -LiteralPath $PackagePath -Type Leaf)) {
        throw "Source package '$PackagePath' not found"
    }

    if (-not (Test-Path -LiteralPath $ImportFile -Type Leaf)) {
        throw "ImportFile '$ImportFile' not found"
    }
    #endregion

    #region Get computer names from file
    $computerNames = (Import-Csv -LiteralPath $ImportFile).ComputerName |
    Sort-Object -Unique

    Write-Verbose "Found '$($computerNames.Count)' unique computer names"

    if (-not $computerNames) {
        Write-Verbose "No computer names found in the import file '$ImportFile'"
        exit
    }
    #endregion

    $packageItem = Get-Item -LiteralPath $PackagePath
    $destinationPathUnc = $DestinationFolder.Replace(':', '$')
    $failedInstalls = @()

    $argumentList = '/package "{0}\{1}" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1' -f $DestinationFolder, $packageItem.Name
}

Process {
    foreach (
        $computer in
        $computerNames
    ) {
        try {
            #region Copy package to remote computer
            try {
                $testParams = @{
                    LiteralPath = "\\$computer\$destinationPathUnc\$($packageItem.Name)"
                    PathType    = 'Leaf'
                }
                if (-not (Test-Path @testParams)) {
                    Write-Verbose "'$computer' copy package to computer"

                    $copyParams = @{
                        LiteralPath = $packageItem.FullName
                        Destination = $testParams.LiteralPath
                        ErrorAction = 'Stop'
                    }
                    Copy-Item @copyParams
                }
                else {
                    Write-Verbose "'$computer' package already on computer"
                }
            }
            catch {
                throw "Failed to copy package to remote computer: $_"
            }
            #endregion

            try {
                $testParams = @{
                    LiteralPath = "\\$computer\$destinationPathUnc\$($packageItem.BaseName) - installed.txt"
                    PathType    = 'Leaf'
                }
                if (Test-Path @testParams) {
                    Write-Verbose "'$computer' package already installed"
                    Continue
                }

                #region Install package on remote computer
                Write-Verbose "'$computer' install package"

                $invokeParams = @{
                    ComputerName = $computer
                    ScriptBlock  = {
                        Start-Process 'msiexec.exe' -ArgumentList $using:argumentList -Wait
                    }
                    ErrorAction  = 'SilentlyContinue'
                }
                Invoke-Command @invokeParams
                #endregion

                #region Test installation
                $installed = $false
                $maxCount = 15
                $currentCount = 0

                while (
                    (-not $installed) -and
                    ($currentCount -lt $maxCount)
                ) {
                    $currentCount++

                    Write-Verbose "'$computer' try connecting ($currentCount\$maxCount)"
                    Start-Sleep -Seconds 1

                    try {
                        $sessionParams = @{
                            ComputerName      = $computer
                            ConfigurationName = $PowerShellEndpointVersion
                            ErrorAction       = 'Stop'
                        }
                        New-PSSession @sessionParams
                        $installed = $true
                    }
                    catch {
                        $lastError = $_
                        $Error.RemoveAt(0)
                    }
                }

                if ($installed) {
                    Write-Verbose "'$computer' package installed successfully"

                    'Installed' | Out-File -LiteralPath $testParams.LiteralPath
                }
                else {
                    $failedInstalls += @{
                        Date         = Get-Date
                        ComputerName = $computer
                    }
                    throw "Package not installed: $lastError"
                }
                #endregion
            }
            catch {
                throw "Failed installing package: $_"
            }
        }
        catch {
            Write-Warning "'$computer' Failed installing package '$($packageItem.Name)': $_"
        }
    }

    #region Report results
    if ($failedInstalls) {
        Write-Warning "Check log file '$FailedInstallLogFile' for failures"
        $failedInstalls | Export-Csv -Path $FailedInstallLogFile
    }
    else {
        Write-Verbose "No failures, all successful"
    }
    #endregion
}