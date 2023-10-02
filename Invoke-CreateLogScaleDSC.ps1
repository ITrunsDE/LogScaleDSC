#Requires -Version 5.0
[float] $version = 0.1

Clear-Host
# load config file
$workingDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$config = Get-Content $(Join-Path $workingDir "config.json") -ErrorAction Stop | ConvertFrom-Json

Configuration InstallLogScaleCollector {

    param (
        [string]
        $Name,

        [string]
        $FilePath,

        [string]
        $ProductId,

        [string]
        $EnrollmentToken
    )

    Import-DscResource -ModuleName "PSDesiredStateConfiguration"

    Node $Name {

        Package LogscaleCollectorInstallation {
            Ensure    = "Present"
            Path      = $FilePath
            Name      = "Humio Log Collector"
            ProductId = $ProductId
        }

        Script CheckEnrolled {

            DependsOn  = "[Package]LogscaleCollectorInstallation"

            GetScript  = {
                $servicePath = (Get-CimInstance Win32_Service -Filter 'Name = "HumioLogCollector"').PathName
                $configPathMatch = $servicePath | Select-String -Pattern '(?:-{1,2}cfg)\s+"([^"]+)"'
                if ($configPathMatch) {
                    return $configPathMatch.Matches.Groups[1].Value.Trim()
                }
                else {
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

            SetScript  = {
                
                # get config path
                $servicePath = (Get-CimInstance Win32_Service -Filter 'Name = "HumioLogCollector"').PathName
                $executePathMatch = $servicePath | Select-String -Pattern "^(.*?)\s*--?cfg"
                if ($executePathMatch) {
                    $executePath = $executePathMatch.Matches.Groups[1].Value.Trim()
                }
                else {
                    return $False
                }

                # enroll client
                Write-Verbose "Enroll client..."
                Start-Process -FilePath $executePath -ArgumentList "enroll $($EnrollmentToken)" -Wait

            }

        }
    }
    
}

function Get-MsiProductCode {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
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
        }
        catch {
            Write-Error "Fehler beim Lesen des ProductCode: $_"
        }  
    }
}

Write-Host -ForegroundColor White ""
Write-Host -ForegroundColor White "      ___                                  ___  "
Write-Host -ForegroundColor White "     (o o)                                (o o) "
Write-Host -ForegroundColor White "    (  V  )   LogScale DSC Creator v$($version)  (  V  )"
Write-Host -ForegroundColor White "    --m-m----------------------------------m-m--"
Write-Host -ForegroundColor White "    "
Write-Host -ForegroundColor White "    (c) 2023 - Sebastian Selig - mylogscale.com"
Write-Host -ForegroundColor White ""
Write-Host -ForegroundColor White ""
Write-Host -ForegroundColor White "[-] Config: " -NoNewline 
Write-Host -ForegroundColor Yellow $(Join-Path $workingDir "config.json")
Write-Host -ForegroundColor White "[-] Start compiling DSC with the following parameters..."
Write-Host ""
Write-Host -ForegroundColor White "[-] Using installation file: " -NoNewline 
Write-Host -ForegroundColor Yellow $config.installation_file
Write-Host ""

#########################################################################################
# Create folder if necessary 
$null = New-Item "$($workingDir)\GroupPolicy" -ItemType Directory -ErrorAction SilentlyContinue

# get ProductId for comparing in DSC script
[string] $ProductId = Get-MsiProductCode -Path $config.installation_file

foreach ($section in $config.sections) {
    
    Write-Host -ForegroundColor White "[-] Section: " -NoNewline
    Write-Host -ForegroundColor Yellow $section.name

    Write-Host -ForegroundColor White "[-] => Create DSC configuration for section"
    $output = InstallLogScaleCollector -Name $section.name -FilePath $config.installation_file -ProductId $ProductId -EnrollmentToken $section.enrollmentToken

    Write-Host -ForegroundColor White "[-] => Create Powershell start script"
    @"
Start-DscConfiguration "$($config.dsc_share)\InstallLogScaleCollector" -ComputerName $($section.name) -Wait -Force
"@ | Out-File -FilePath "$($workingDir)\GroupPolicy\DSC_LogScale_$($section.name).ps1" -Force
    
    Write-Host -ForegroundColor White "[-] => Powershell start script is stored here: " -NoNewline 
    Write-Host -ForegroundColor Cyan "$($workingDir)\DSC_LogScale_$($section.name).ps1"
    Write-Host ""

}

#########################################################################################
Write-Host ""
Write-Host -ForegroundColor White "The files are created for each section. Now you just have to create a group policy for each section and add the corresponding script in the Powershell Scripts section under Computer Configuration/Windows Settings/...."
Write-Host ""
Write-Host -ForegroundColor White "The folder " -NoNewLine
Write-Host -ForegroundColor Yellow "InstallLogScaleCollector" -NoNewline
Write-Host -ForegroundColor White " must be copied into the directory " -NoNewline
Write-Host -ForegroundColor Yellow $config.dsc_share -NoNewline
Write-Host "."
Write-Host ""
Write-Host -ForegroundColor White "For more information visit " -NoNewLine
Write-Host -ForegroundColor Yellow "https://mylogscale.com" -NoNewline
Write-Host ""
Write-Host ""