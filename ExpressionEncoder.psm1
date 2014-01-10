function Ensure-Encoder {
    Add-Type -Path @{
        'AMD64'="${env:ProgramFiles(x86)}\Microsoft Expression\Encoder 4\SDK\Microsoft.Expression.Encoder.dll"
        'x86'="$env:ProgramFiles\Microsoft Expression\Encoder 4\SDK\Microsoft.Expression.Encoder.dll"
    }[$env:PROCESSOR_ARCHITECTURE]
}

function Invoke-EncodeXESC($FileName) {
    Ensure-Encoder
    msg "Encoding $FileName" -Level 1
    err {
        $MediaItem = New-Object Microsoft.Expression.Encoder.MediaItem $FileName
        $Job = New-Object Microsoft.Expression.Encoder.Job
        $Job.MediaItems.Add($MediaItem)
        <# Doesn't fire until after Encode returns, TODO alternate event registration
        Register-ObjectEvent $Job EncodeProgress -Action { param($Sender, $e)
            msg "Encoder Progress $($e.Progress)" -Level 1
        } | Out-Null
        #>
        msg "Encoding output $(Split-Path -Parent $FileName)" -Level 1
        $Job.OutputDirectory = $env:TEMP # Split-Path -Parent $FileName
        
        $job.DefaultMediaOutputFileName = "$([Guid]::NewGuid().ToString()).wmv"
        $Job.Encode()
        $outFile = "$($Job.ActualOutputDirectory)\$($job.DefaultMediaOutputFileName)"
        $Job.Dispose()
        copy-file "$outFile" "$FileName.wmv"

        "$FileName.wmv"
    } 'Encode-XESC'
}

function Start-Video([switch]$ExitAfter) {
    write-host "Starting Encoder..."
    Ensure-Encoder
    $job = new-object Microsoft.Expression.Encoder.ScreenCapture.ScreenCaptureJob
    $job.OutputPath = gs ScreencastDir
    $job.Start()
    write-host "Now Recording"
    $p = New-Pipe 'EECaptureTask' '.'
    $p.Recv()
    $job.Stop()
    write-host "File: $($job.ScreenCaptureFileName)"
    $p = New-Pipe 'EECaptureTaskRet' '.'
    $p.Send($job.ScreenCaptureFileName)
    if ($ExitAfter) {
        [Environment]::Exit(0)
    }
}

function Stop-Video {
    $p = New-Pipe 'EECaptureTask' '.'
    $p.Send('stop') | out-null
    $p = New-Pipe 'EECaptureTaskRet' '.'
    $p.Recv()
}

# Open a 32-Bit shell that will run .NET v3.5 binaries on v4 forcefully
function Invoke-Net4x86($Command) {
    Start-Process @{'amd64'="$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell_net4.exe"
                    'x86'=  "$env:windir\System32\WindowsPowerShell\v1.0\powershell_net4.exe"
    }[$env:PROCESSOR_ARCHITECTURE] -ArgumentList "-NoExit -WindowStyle Minimized -Command Register-EngineEvent Shell.Startup -Action {$Command} | select -exp Command"
}

function Invoke-ConfigureForcedNet4PowerShell {
    foreach ($PSRoot in @("$env:windir\System32\WindowsPowerShell\v1.0", "$env:windir\syswow64\WindowsPowerShell\v1.0")) {
        if (Test-Path $PSRoot) {
            # Helpfully the ISE in PS 3.0 has exactly the configuration we need.
            $ISE = "$PSRoot\powershell_ise.exe"
            $PS = "$PSRoot\powershell.exe"
            $PSNet4 = "$PSRoot\powershell_net4.exe"
            if (!(Test-Path $PSNet4)) {
                if (!(Test-Elevated)) { Throw "Elevation required to configure .Net4 forced shell" }
                copy $PS $PSNet4
                copy "$ISE.config" "$PSNet4.config"
                "$PSNet4 -NoProfile Set-ExecutionPolicy Unrestricted -Force" | iex 
            }
        }
    }
}

function Install-ExpressionEncoder {
    Invoke-ConfigureForcedNet4PowerShell
    # Kick off the EE install.
    $Installer = gs Packages.Package | ? Name -eq ExpressionEncoder4 | select -exp Location
    "$Installer -q" | iex
    # Find setup
    $xs = $null
    do {
        $xs = ps "XSetup" -ErrorAction SilentlyContinue | select -first 1
        sleep 1
    } while ($xs -eq $null)
    # found setup
    do {
        $xs = ps "XSetup" -ErrorAction SilentlyContinue | select -first 1
        sleep 1
    } while ($xs -ne $null)
    # Done
}