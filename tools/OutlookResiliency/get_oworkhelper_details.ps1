# Get detailed oWorkhelper registry information
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins\oWorkHelper"

Write-Host "=========================================="
Write-Host "oWorkhelper Registry Details"
Write-Host "=========================================="
Write-Host ""

if (Test-Path -Path $regPath) {
    Write-Host "Path: $regPath"
    Write-Host "Status: FOUND"
    Write-Host ""

    # Get all properties
    $allProps = Get-ItemProperty -Path $regPath

    Write-Host "All Properties:"
    foreach ($prop in $allProps.PSObject.Properties) {
        if ($prop.Name -ne "PSPath" -and $prop.Name -ne "PSParentPath" -and $prop.Name -ne "PSChildName" -and $prop.Name -ne "PSDrive" -and $prop.Name -ne "PSProvider") {
            Write-Host "  $($prop.Name) = $($prop.Value)"
        }
    }

    Write-Host ""
    Write-Host "Interpretation:"

    # LoadBehavior
    if ($allProps.LoadBehavior) {
        $lb = $allProps.LoadBehavior
        $lbDesc = switch ($lb) {
            3 { "Startup Load (Normal)" }
            9 { "Disabled (Slow Load)" }
            16 { "Load on Demand" }
            2 { "Depends on Setting" }
            8 { "Load Failed" }
            0 { "No Auto Load" }
            default { "Unknown Value" }
        }
        Write-Host "  LoadBehavior: $lb = $lbDesc"
    } else {
        Write-Host "  LoadBehavior: NOT SET"
    }

    # Manifest
    if ($allProps.Manifest) {
        Write-Host "  Manifest Path: $($allProps.Manifest)"
        Write-Host "    Does it exist? $(if (Test-Path $allProps.Manifest) { 'YES' } else { 'NO - FILE NOT FOUND' })"
    } else {
        Write-Host "  Manifest: NOT SET"
    }

} else {
    Write-Host "oWorkhelper not found in HKCU\Software\Microsoft\Office\16.0\Outlook\Addins"
}

Write-Host ""
Write-Host "=========================================="
Write-Host "Check Resiliency DisabledItems"
Write-Host "=========================================="
Write-Host ""

$disabledPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems"
if (Test-Path -Path $disabledPath) {
    $disabled = Get-ItemProperty -Path $disabledPath
    $foundDisabled = $false
    foreach ($prop in $disabled.PSObject.Properties) {
        if ($prop.Value -match "oWorkHelper") {
            Write-Host "FOUND IN DISABLED ITEMS:"
            Write-Host "  $($prop.Name) = $($prop.Value)"
            $foundDisabled = $true
        }
    }
    if (-not $foundDisabled) {
        Write-Host "oWorkhelper NOT in DisabledItems (good sign)"
    }
} else {
    Write-Host "DisabledItems registry not found (first run?)"
}

Write-Host ""
Write-Host "=========================================="
Write-Host "Next Steps"
Write-Host "=========================================="
Write-Host ""
Write-Host "1. If LoadBehavior = 3, that's normal startup load"
Write-Host "2. If LoadBehavior = 9, add-in is disabled by Outlook"
Write-Host "3. If LoadBehavior = 16, add-in is set to load on demand"
Write-Host "4. If Manifest file doesn't exist, that's a deployment problem"
Write-Host ""
