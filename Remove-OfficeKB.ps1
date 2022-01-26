# Функция удаляет обновление для MS Office KB5002104, KB5002099 
# https://www.devhut.net/access-lock-file-issues/
# Example: Remove-OfficeKB -Computername Comp1 -Result

function Remove-OfficeKB { # Удаление обновлений MS Office 2013-2016
        param(
          [parameter(Mandatory=$true)][string]$ComputerName,
          [switch]$Result
        )
# Версии Office 
$2013x32 = '{90150000-0011-0000-0000-0000000FF1CE}'
$2013x64 = '{90150000-0011-0000-1000-0000000FF1CE}'
$2016x32 = '{90160000-0011-0000-0000-0000000FF1CE}'
$2016x64 = '{90160000-0011-0000-1000-0000000FF1CE}'
# GUID обновлений
# Office 2013 x32 KB5002104
$KB5002104x32 = '{8FE4AEF3-DE32-4A09-9302-BB30F9088699}'
# Office 2013 x64 KB5002104
$KB5002104x64 = '{AC593D32-2D34-48A1-82B5-52FC0CFDA409}'
# Office 2016 x32 KB5002099    
$KB5002099x32 = '{BA36399C-CF0F-4368-8327-7D35302BF0BB}'
# Office 2016 x64 KB5002099    
$KB5002099x64 = '{127B2615-3D07-4189-B91C-44A04FB7A55F}'

# Собираем информацию об установленном пакете Office
$Session = New-PSSession -ComputerName $ComputerName
$OfficeVersion = Invoke-Command -Session $Session -ScriptBlock {
# Ищем 64х...
Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{901?0000-001?-0000-?000-0000000FF1CE}"
# ...или 32х
Get-ItemProperty "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{901?0000-001?-0000-?000-0000000FF1CE}"}
Remove-PSSession -Session $Session

if ($OfficeVersion.PSChildName -like $2013x32) {
    Write-Host "Office 2013 x32" $OfficeVersion.DisplayVersion
    $Session = New-PSSession -ComputerName $ComputerName
    Invoke-Command -Session $Session -ScriptBlock {
        Start-Process msiexec.exe -ArgumentList "/package $using:2013x32 /uninstall $using:KB5002104x32 /qn /l c:\office.log /norestart" -Wait -NoNewWindow
        if ($using:Result) {gc c:\office.log | ? {$_.trim() -ne "" }}}
        Remove-PSSession -Session $Session}

    elseif ($OfficeVersion.PSChildName -like $2013x64) {
        Write-Host "Office 2013 x64" $OfficeVersion.DisplayVersion
        $Session = New-PSSession -ComputerName $ComputerName
        Invoke-Command -Session $Session -ScriptBlock {
            Start-Process msiexec.exe "/package $using:2013x64 /uninstall $using:KB5002104x64 /qn /l c:\office.log /norestart" -Wait -NoNewWindow
            if ($using:Result) {gc c:\office.log | ? {$_.trim() -ne "" }}}
            Remove-PSSession -Session $Session}

        elseif ($OfficeVersion.PSChildName -like $2016x32) {
            Write-Host "Office 2016 x32" $OfficeVersion.DisplayVersion
            $Session = New-PSSession -ComputerName $ComputerName
            Invoke-Command -Session $Session -ScriptBlock {
                Start-Process msiexec.exe "/package $using:2016x32 /uninstall $using:KB5002099x32 /qn /l c:\office.log /norestart" -Wait -NoNewWindow
                if ($using:Result) {gc c:\office.log | ? {$_.trim() -ne "" }}}
                Remove-PSSession -Session $Session}

            elseif ($OfficeVersion.PSChildName -like $2016x64) {
                Write-host "Office 2016 x64" $OfficeVersion.DisplayVersion
                $Session = New-PSSession -ComputerName $ComputerName
                Invoke-Command -Session $Session -ScriptBlock {
                    Start-Process msiexec.exe "/package $using:2016x64 /uninstall $using:KB5002099x64 /qn /l c:\office.log /norestart" -Wait -NoNewWindow
                    if ($using:Result) {gc c:\office.log | ? {$_.trim() -ne "" }}}
                    Remove-PSSession -Session $Session}
        }