#Requires -Version 5.0

Clear-Host
# load config file
$workingDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$config = Get-Content $(Join-Path $workingDir "config.json") -ErrorAction Stop | ConvertFrom-Json

Configuration $config.name {

    param (
        [string]
        $FilePath,

        [string]
        $ProductId,

        [string]
        $EnrollmentToken
    )

    Import-DscResource -ModuleName "PSDesiredStateConfiguration"

    Package LogscaleCollectorInstallation {
        Ensure = "Present"
        Path = $FilePath
        Name = "Humio Log Collector"
        ProductId = $ProductId
    }

    Script CheckEnrolled {

        DependsOn = "[Package]LogscaleCollectorInstallation"

        GetScript = {
            $servicePath = (Get-CimInstance Win32_Service -Filter 'Name = "HumioLogCollector"').PathName
            $configPathMatch = $servicePath | Select-String -Pattern '(?:-{1,2}cfg)\s+"([^"]+)"'
            if ($configPathMatch) {
                return $configPathMatch.Matches.Groups[1].Value.Trim()
            } else {
                return $False
            }
        }

        TestScript = {

            # get InstalledLocation
            $configPath = [scriptblock]::Create($GetSCript).Invoke()
            if (!$configPath) {
                return $False
            }

            # load config file and extract parameters from the file
            $configData = Get-Content $configPath
            
            # $dataDirectoryMatch = $configData | Select-String -Pattern 'dataDirectory:\s+(.+)'
            # if ($dataDirectoryMatch) {
            #     $dataDirectoryValue = $dataDirectoryMatch.Matches.Groups[1].Value.Trim()
            # }

            # $fleetManagementMatch = $configData | Select-String -Pattern 'fleetManagement:\s+(.+)'
            # if ($fleetManagementMatch) {
            #     $fleetManagementValue = $fleetManagementMatch.Matches.Groups[1].Value.Trim()
            # }

            # $urlMatch = $configData | Select-String -Pattern 'url:\s+(.+)'
            # if ($urlMatch) {
            #     $urlValue = $urlMatch.Matches.Groups[1].Value.Trim()
            # }

            $modeMatch = $configData | Select-String -Pattern 'mode:\s+(.+)'
            if ($modeMatch) {
                $modeValue = $modeMatch.Matches.Groups[1].Value.Trim()
            }


            # is the collector enrolled
            if ($modeValue -ilike "full") {
                Write-Verbose "Client already enrolled."
                return $true
            }

            Write-Verbose "Client not enrolled yet."
            return $false
            
        }

        SetScript = {
            
            # get config path
            $servicePath = (Get-CimInstance Win32_Service -Filter 'Name = "HumioLogCollector"').PathName
            $executePathMatch = $servicePath | Select-String -Pattern "^(.*?)\s*--?cfg"
            if ($executePathMatch) {
                $executePath = $executePathMatch.Matches.Groups[1].Value.Trim()
            } else {
                return $False
            }

            # enroll client
            Write-Verbose "Enroll client..."
            Start-Process -FilePath $executePath -ArgumentList "enroll $($using:enrollmentToken)" -Wait

        }

    }

}

function Get-MsiProductCode {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Path
    )
    
    begin {
        $windowsInstaller = New-Object -ComObject WindowsInstaller.Installer
    }
    
    process {
        try {
            $msiDatabase = $windowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $windowsInstaller, @($Path, 0))
            
            $query = "SELECT Value FROM Property WHERE Property='ProductCode'"
            $view = $msiDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $msiDatabase, ($query))
            $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null)
            
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
            $productCode = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
        
            if ($null -ne $productCode) {
                return [string] $productCode.Replace("{", "").Replace("}", "").Trim()
            } 
        } catch {
            Write-Error "Fehler beim Lesen des ProductCode: $_"
        }  
    }
}

# get ProductCode from msi package
[string] $ProductId = Get-MsiProductCode -Path $config.path_to_msi

Write-Host ""
Write-Host "      ___                              ___  "
Write-Host "     (o o)                            (o o) "
Write-Host "    (  V  )   LogScale DSC Creator   (  V  )"
Write-Host "    --m-m------------------------------m-m--"
Write-Host "    "
Write-Host "  (c) 2023 - Sebastian Selig - mylogscale.com"
Write-Host ""
Write-Host ""
Write-Host -ForegroundColor White "[-] Config: " -NoNewline 
Write-Host -ForegroundColor Yellow $(Join-Path $workingDir "config.json")
Write-Host -ForegroundColor White "[-] Start compiling DSC with the following parameters..."
Write-Host ""
Write-Host -ForegroundColor White "[-] Installationfile: " -NoNewline 
Write-Host -ForegroundColor Yellow $config.path_to_msi
Write-Host -ForegroundColor White "[-] ProductCode: " -NoNewLine
Write-Host -ForegroundColor Yellow $ProductId
Write-Host -ForegroundColor White "[-] EnrollmentToken: " -NoNewLine
Write-Host -ForegroundColor Yellow $config.enrollmentToken
Write-Host ""
Write-Host -ForegroundColor Green "DSC is loaded..."

#########################################################################################
# Create DSC configuration file
$arguments = @{
    FilePath = $config.path_to_msi
    ProductId = $ProductId 
    EnrollmentToken = $config.enrollmentToken
}
$output = Invoke-Expression -Command "$($config.name) @arguments"
#########################################################################################
Write-Host -ForegroundColor Green "DSC is compiled..."
Write-Host ""
Write-Host -ForegroundColor White "[-] DSC Configuration saved: " -NoNewLine
Write-Host -ForegroundColor Cyan $output

#########################################################################################
# Create Powershell file to run in Group Policy
@"
Start-DscConfiguration "$($config.share)\$($config.name)" -Wait -Force
"@ | Out-File -FilePath "$($workingDir)\DSC_LogScale_$($config.name).ps1" -Force
#########################################################################################

Write-Host -ForegroundColor White "[-] Powershell file for the GPO is stored here: " -NoNewline 
Write-Host -ForegroundColor Cyan "$($workingDir)\DSC_LogScale_$($config.name).ps1"
Write-Host ""

# Start-DscConfiguration InstallAndEnrollLogscaleCollector -Verbose -Wait -Force