param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64"
)

$ErrorActionPreference = "Stop"

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$projectPath = Join-Path $root "VMWV.Modern\VMWV.Modern.csproj"
$artifactsDir = Join-Path $root "artifacts"
$publishDir = Join-Path $artifactsDir "publish"
$installerDir = Join-Path $artifactsDir "installer"
$stageDir = Join-Path $installerDir "stage"
$outputExe = Join-Path $installerDir "Install-VVC.exe"
$outputZip = Join-Path $installerDir "Install-VVC.zip"
$sedFile = Join-Path $installerDir "installer.sed"

Write-Host "[Build] Cleaning artifacts"
Remove-Item -Path $publishDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path $installerDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $publishDir -Force | Out-Null
New-Item -ItemType Directory -Path $stageDir -Force | Out-Null

Write-Host "[Build] Publishing app"
dotnet publish $projectPath -c $Configuration -r $Runtime --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:PublishTrimmed=false -o $publishDir

Write-Host "[Build] Preparing installer stage"
$payloadZip = Join-Path $stageDir "payload.zip"
Compress-Archive -Path (Join-Path $publishDir "*") -DestinationPath $payloadZip -Force

Copy-Item -Path (Join-Path $root "build-tools\installer\install.ps1") -Destination (Join-Path $stageDir "install.ps1") -Force
Copy-Item -Path (Join-Path $root "build-tools\installer\uninstall.ps1") -Destination (Join-Path $stageDir "uninstall.ps1") -Force

@"
@echo off
powershell.exe -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
"@ | Set-Content -Path (Join-Path $stageDir "install.cmd") -Encoding ascii

Compress-Archive -Path (Join-Path $stageDir "*") -DestinationPath $outputZip -Force
Write-Host "[Build] ZIP package created: $outputZip"

$iexpressCandidates = @(
    (Join-Path $env:SystemRoot "System32\iexpress.exe"),
    (Join-Path $env:SystemRoot "SysWOW64\iexpress.exe")
)

$iexpress = $iexpressCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iexpress) {
    $iexpressCommand = Get-Command iexpress.exe -ErrorAction SilentlyContinue
    if ($iexpressCommand) {
        $iexpress = $iexpressCommand.Source
    }
}

if (-not $iexpress) {
    Write-Warning "IExpress was not found. Installer EXE was not created. Use ZIP package instead."
    exit 0
}

Write-Host "[Build] Generating IExpress SED"
$sed = @"
[Version]
Class=IEXPRESS
SEDVersion=3

[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=Installation complete.
TargetName=$outputExe
FriendlyName=VVC Installer
AppLaunched=install.cmd
PostInstallCmd=<None>
AdminQuietInstCmd=
UserQuietInstCmd=
SourceFiles=SourceFiles
SelfDelete=0
FILE0=payload.zip
FILE1=install.ps1
FILE2=uninstall.ps1
FILE3=install.cmd

[SourceFiles]
SourceFiles0=$stageDir

[SourceFiles0]
%FILE0%=
%FILE1%=
%FILE2%=
%FILE3%=
"@
$sed | Set-Content -Path $sedFile -Encoding ascii

Write-Host "[Build] Building EXE installer via IExpress"
& $iexpress /N $sedFile | Out-Null

if (Test-Path $outputExe) {
    Write-Host "[Build] EXE installer created: $outputExe"
} else {
    Write-Warning "IExpress did not produce EXE. Use ZIP package: $outputZip"
}
