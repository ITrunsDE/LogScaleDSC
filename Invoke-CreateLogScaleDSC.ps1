param (
    [switch]$UseSigning
)

#Requires -Version 5.0
[float] $version = 0.2

Clear-Host

# load module and configuration file
$workingDir = Split-Path $MyInvocation.MyCommand.Path -Parent
Import-Module $(Join-Path $workingDir "LogScaleDSC.psm1")

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

    Node localhost {

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
                # Write-Verbose "Client: $($executePath)"
                # Write-Verbose "EnrollmentToken: $($using:EnrollmentToken)"
                # Write-Verbose "enroll $($using:EnrollmentToken)"
                Start-Process -FilePath $executePath -ArgumentList "enroll $($using:EnrollmentToken)" -Wait -Verbose

            }

        }

        Service LogScaleService {
            Name        = "HumioLogCollector"
            StartupType = "Automatic"
            Ensure      = "Present"
            DependsOn   = "[Script]CheckEnrolled"
        }
    }   
}

# Create folder if necessary 
$null = New-Item "$($workingDir)\DSC" -ItemType Directory -ErrorAction SilentlyContinue
$null = New-Item "$($workingDir)\GroupPolicy" -ItemType Directory -ErrorAction SilentlyContinue

# get ProductId for comparing in DSC script
[string] $ProductId = Get-MsiInformation -Path $config.installation_file -Type "ProductCode"
[string] $ProductName = Get-MsiInformation -Path $config.installation_file -Type "ProductName"
[string] $ProductVersion = Get-MsiInformation -Path $config.installation_file -Type "ProductVersion"

Write-Host -ForegroundColor White ""
Write-Host -ForegroundColor White "      ___                                  ___  "
Write-Host -ForegroundColor White "     (o o)                                (o o) "
Write-Host -ForegroundColor White "    (  V  )   LogScale DSC Creator v$($version)  (  V  )"
Write-Host -ForegroundColor White "    --m-m----------------------------------m-m--"
Write-Host -ForegroundColor White "    "
Write-Host -ForegroundColor White "    (c) 2023 - Sebastian Selig - mylogscale.com"
Write-Host -ForegroundColor White ""
Write-Host -ForegroundColor White ""
Write-Host -ForegroundColor White "[i] Using Config file: " -NoNewline 
Write-Host -ForegroundColor Yellow $(Join-Path $workingDir "config.json")
Write-Host -ForegroundColor White "[+] Start compiling DSC with the following parameters..."
Write-Host ""
Write-Host -ForegroundColor White "[i] Using installation file: " -NoNewline 
Write-Host -ForegroundColor Yellow $config.installation_file
Write-Host -ForegroundColor White "[i] Found:" -NoNewline 
Write-Host -ForegroundColor Yellow "$ProductName $($ProductVersion.Trim())"
Write-Host ""

# get signing certificate
$signCert = Get-SigningCertificate
$signScripts = $false
if ($UseSigning -and $signCert) {
    $signScripts = $true
    Write-Host -ForegroundColor White "[i] Using this certificate for the signature"
    Write-Host -ForegroundColor White "[+] => Thumbprint: " -NoNewline
    Write-Host -ForegroundColor Yellow $signCert.Thumbprint
    Write-Host -ForegroundColor White "[+] => Subject: " -NoNewline
    Write-Host -ForegroundColor Yellow $signCert.Subject
    Write-Host -ForegroundColor White "[+] => Issuer: " -NoNewline
    Write-Host -ForegroundColor Yellow $signCert.Issuer
    Write-Host -ForegroundColor White "[+] => Valid until: " -NoNewline
    Write-Host -ForegroundColor Yellow $signCert.NotAfter
    Write-Host -ForegroundColor White "[+] => Friendly name: " -NoNewline
    Write-Host -ForegroundColor Yellow $signCert.FriendlyName
    Write-Host ""
}
elseif ($UseSigning) {
    Write-Host -ForegroundColor Red "[e] No code signing certificate found"
    Write-Host ""
}
elseif ($signCert) {
    Write-Host -ForegroundColor White "[i] Found a signing certificate"
    Write-Host -ForegroundColor White "[+] => Thumbprint: " -NoNewline
    Write-Host -ForegroundColor Yellow $signCert.Thumbprint
    Write-Host -ForegroundColor White "[+] => Subject: " -NoNewline
    Write-Host -ForegroundColor Yellow $signCert.Subject
    Write-Host -ForegroundColor White "[+] => Issuer: " -NoNewline
    Write-Host -ForegroundColor Yellow $signCert.Issuer
    Write-Host -ForegroundColor White "[+] => Valid until: " -NoNewline
    Write-Host -ForegroundColor Yellow $signCert.NotAfter
    Write-Host -ForegroundColor White "[+] => Friendly name: " -NoNewline
    Write-Host -ForegroundColor Yellow $signCert.FriendlyName
    Write-Host -ForegroundColor DarkGreen "[i] To use this signing certificate run again with " -NoNewline
    Write-Host -ForegroundColor Cyan "-UseSigning"
    Write-Host ""
}

#########################################################################################
foreach ($section in $config.sections) {
    
    Write-Host -ForegroundColor White "[i] Section: " -NoNewline
    Write-Host -ForegroundColor Yellow $section.name

    Write-Host -ForegroundColor White "[+] => Create DSC configuration " -NoNewline
    $outputPath = Join-Path "DSC" $section.name
    $null = InstallLogScaleCollector -Name $section.name -FilePath $config.installation_file -ProductId $ProductId -EnrollmentToken $($section.enrollmentToken) -OutputPath $outputPath
    Write-Host -ForegroundColor Cyan "-" $outputPath

    # sign the DSC files with the code signing certificate
    if ($signScripts) {
        Write-Host -ForegroundColor White "[+] => Sign DSC configuration"
        $null = Set-AuthenticodeSignature "$($outputPath)\localhost.mof" $signCert
    }

    Write-Host -ForegroundColor White "[+] => Create PowerShell start script " -NoNewline
    @"
Start-DscConfiguration "$($config.dsc_share)\$($outputPath)" -Wait -Force
"@ | Out-File -FilePath "$($workingDir)\GroupPolicy\DSC_LogScale_$($section.name).ps1" -Force
    
    Write-Host -ForegroundColor Cyan "- GroupPolicy\DSC_LogScale_$($section.name).ps1"
    
    # sign the scripts with the code signing certificate
    if ($signScripts) {
        Write-Host -ForegroundColor White "[+] => Sign PowerShell start script with certificate"
        $null = Set-AuthenticodeSignature "$($workingDir)\GroupPolicy\DSC_LogScale_$($section.name).ps1" $signCert
    }
    
    Write-Host ""

}

#########################################################################################
Write-Host ""
Write-Host -ForegroundColor White "The files are created for each section. Now you just have to create a group policy for each section and add the corresponding script in the Powershell Scripts section under Computer Configuration/Windows Settings/...."
Write-Host ""
Write-Host -ForegroundColor White "The folder " -NoNewLine
Write-Host -ForegroundColor Yellow "DSC" -NoNewline
Write-Host -ForegroundColor White " must be copied into the directory " -NoNewline
Write-Host -ForegroundColor Yellow $config.dsc_share -NoNewline
Write-Host "."
Write-Host ""
Write-Host -ForegroundColor White "For more information visit " -NoNewLine
Write-Host -ForegroundColor Yellow "https://mylogscale.com" -NoNewline
Write-Host ""
Write-Host ""