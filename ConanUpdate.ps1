[CmdletBinding()]
param(
    [String]$SteamDIR = "C:\Data\Steam",
    [String]$ConanDIR = "C:\Data\ConanServer",
    [String]$BackupDIR = "C:\Data\ConanSave",
    [String]$Proc = "ConanSandboxServer-Win64-Test",
    [String]$NumOfBackupToKeep = 5,
    [String]$MultiHome = "x.x.x.x",
    [String]$ProcName = "Conan Exiles - press Ctrl+C to shutdown",
    [Switch]$NoUpdate,
    [Switch]$NoStart,
    [Switch]$RaidOn,
    [Switch]$RaidOff
)

#***** Methods *****

function Write-Log
{
    param($log)
    $date = Get-Date -f "dd.MM.yyyy"
    ("{0} --> {1}" -f (Get-Date), $log) | Out-File ("{0}\{1}_ConanUpdate.log" -f $scriptDirectory, $date) -Append
    Write-Host ("{0} - {1}" -f $date, $log) -ForegroundColor Green
}

function Get-ScriptDirectory
{
    $scriptInvocation = (Get-Variable MyInvocation -Scope 1).Value
    return Split-Path $scriptInvocation.MyCommand.Path
}

function Stop-ConanProcess
{
    try
    {
        if (Get-Process $Proc -ErrorAction SilentlyContinue)
        {
            Write-Log -Log "1/4 - Stopping process"   
            $Wshell = New-Object -ComObject wscript.shell;
            $Wshell.AppActivate($ProcName) | Out-Null
            Start-Sleep -Seconds 1
            $Wshell.SendKeys('^c') | Out-Null
            Start-Sleep -Seconds 10        
            if (Get-Process $Proc -ErrorAction SilentlyContinue)
            { 
                Stop-Process -Name $Proc
                Write-Log -Log "1/4 - Process stopped forcefully"
                Start-Sleep -Milliseconds 5000
            }        
            else
            {
                Write-Log -Log "1/4 - Process stopped gracefully"
            }
        }
        else
        {
            Write-Log -Log "1/4 - Process is not running"    
        }
    }
    catch
    {
        Write-Log -Log "An error occurred while stopping the process"
    }
}

function Start-ConanBackup
{
    if ([System.IO.File]::Exists("{0}\last.txt" -f $BackupDIR))
    {
        $last = [Convert]::ToInt32((Get-Content ("{0}\last.txt" -f $BackupDIR)))
    }
    else
    {
        $last = 0
    }
    $last++
    Write-Log -Log "2/4 - Removing old backup"
    foreach ($dir in (Get-ChildItem $BackupDIR | Where {$_.PSIsContainer}))
    {
        if ($dir.Name.Length -eq 1)
        {
            $dNum = [Convert]::ToInt32($dir.Name)
            if ($dNum -le ($last - $NumOfBackupToKeep))
            {
                Remove-Item ("{0}\{1}" -f $BackupDIR, $dir.Name) -Force -Recurse
            }
        }
    }

    Write-Log -Log "2/4 - Creating backup"
    if ([System.IO.Directory]::Exists("{0}\ConanSandbox\Saved" -f $ConanDIR))
    {
        Copy-Item ("{0}\ConanSandbox\Saved" -f $ConanDIR) -destination ("{0}\{1}" -f $BackupDIR, $last.ToString()) -recurse
        Start-Sleep -Milliseconds 1000
        Write-Log -Log "2/4 - Backup created successfully"
        $last.ToString() | Out-File ("{0}\last.txt" -f $BackupDIR)
    }
    else
    {
        Write-Log -Log "2/4 - Catastrophic failure - Saved folder doesn't exist"
        Exit(-1)
    }
}

function Update-Conan
{
    Write-Log -Log "3/4 - Starting update"
    $startCmd = ("{0}\steamcmd.exe" -f $SteamDIR)
    $args = '+login anonymous', ("+force_install_dir {0}" -f $ConanDIR), '+app_update 443030', 'validate', '+quit'
    Start-Process -FilePath $startCmd -ArgumentList $args -Wait
    Write-Log -Log "3/4 - Update completed"
}

function Set-ConanRaid
{
    param($Active)
    $settingsFile = ("{0}\ConanSandbox\Saved\Config\WindowsServer\ServerSettings.ini" -f $ConanDIR)
    if ([System.IO.File]::Exists($settingsFile))
    {        
        if ($Active)
        {
            (Get-Content $settingsFile) | Foreach-Object {$_ -replace '^CanDamagePlayerOwnedStructures=.+$', 
            "CanDamagePlayerOwnedStructures=True"} | Set-Content $settingsFile
        }
        else
        {
            (Get-Content $settingsFile) | Foreach-Object {$_ -replace '^CanDamagePlayerOwnedStructures=.+$', 
            "CanDamagePlayerOwnedStructures=False"} | Set-Content $settingsFile
        }
    }
}

function Start-ConanServer
{
    Write-Log -Log "4/4 - Setting raid policy"
    if ($RaidOn.IsPresent)
    {
        Set-ConanRaid -Active $true
    }
    elseif ($RaidOff.IsPresent)
    {
        Set-ConanRaid -Active $false
    }
    Write-Log -Log "4/4 - Starting server"
    if (!$NoStart.IsPresent)
    {
        $startCmd = ("{0}\ConanSandboxServer.exe" -f $ConanDIR)
        $args = ("ConanSandbox?Multihome={0}?listen?" -f $MultiHome), '-nosteamclient', '-game', '-server', '-log'
        Start-Process -FilePath $startCmd -ArgumentList $args
    }
    Write-Log -Log "4/4 - Server started"
}

#***** Main *****

$scriptDirectory = Get-ScriptDirectory
Write-Log -Log "---------- Start update script ----------"

#Stopping process 
#Find a better way to stop the server when the "unbelievable that it's not already there" RCON is implemented
Stop-ConanProcess

#Launch backup
Start-ConanBackup

#Start update
if (!$NoUpdate.IsPresent)
{
    Update-Conan
}
else
{
    Write-Log -Log "3/4 - Update not needed"
}

#Launch game server
Start-ConanServer

Write-Log -Log "---------- End update script ----------"