param (
    [int] $ServerPort = 2002,              # or whatever port you like
    [string] $WorkingFolder = "$env:TEMP"  # for temporary and lock files
)


$LibreOfficePath = (Get-ItemProperty -Path 'HKLM:SOFTWARE\LibreOffice\UNO\InstallPath' -Name '(default)' -ErrorAction Stop).'(default)'

# IMPORTANT: Folder where soffice.exe is installed MUST be added to PATH and ...
if (-not ($env:Path.Contains($LibreOfficePath)))
{
    if (-not ($env:Path.EndsWith(';')))
    {
        $env:Path += ';'
    }

    $env:Path += $LibreOfficePath
}

# ... UNO_PATH environment variables before launching the server!!!
if (-not $env:UNO_PATH)
{
    $env:UNO_PATH = $LibreOfficePath
}


$SOfficeExe = Join-Path $LibreOfficePath 'soffice.exe'

if (-not (Test-Path $SOfficeExe))
{
    Write-Error -Message 'LibreOffice launcher not found.' -RecommendedAction 'Install latest version of LibreOffice (64 bit) from https://www.libreoffice.org/'
}
else
{
    Set-Location -Path $WorkingFolder

    # Launch soffice.exe in server mode
    & $SOfficeExe '--headless' "--accept=socket,host=localhost,port=$ServerPort,tcpNoDelay=1;urp" '--nodefault' '--nofirststartwizard' '--nolockcheck' '--nologo' '--norestore' '--invisible'
}

