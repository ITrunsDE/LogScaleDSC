function Get-MsiInformation {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $Path,

        [Parameter(Mandatory = $true)]
        [string]
        $Type

    )
    
    begin {
        $windowsInstaller = New-Object -ComObject WindowsInstaller.Installer
    }
    
    process {
        try {
            $msiDatabase = $windowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $windowsInstaller, @($Path, 0))
            
            $query = "SELECT Value FROM Property WHERE Property='$($Type)'"
            $view = $msiDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $msiDatabase, ($query))
            $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null)
            
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
            $productCode = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
        
            if ($null -ne $productCode) {
                return [string] $productCode.Replace("{", "").Replace("}", "").Trim()
            } 
        }
        catch {
            Write-Error "Fehler beim Lesen des $($Type): $_"
        }  
    }
}

Export-ModuleMember -Function *