$ErrorActionPreference = "Stop"

$SkillRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$FontSource = Join-Path $SkillRoot "assets\fonts"
$UserFonts = Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Fonts"
$RegistryPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"

New-Item -ItemType Directory -Force -Path $UserFonts | Out-Null

$Fonts = @(
  @{ File = "BaiWuchangKeke-Thin.ttf"; Name = "白无常可可体 细 (TrueType)" },
  @{ File = "BaiWuchangKeke-Regular.ttf"; Name = "白无常可可体 常规 (TrueType)" },
  @{ File = "BaiWuchangKeke-Bold.ttf"; Name = "白无常可可体 粗 (TrueType)" }
)

foreach ($Font in $Fonts) {
  $Source = Join-Path $FontSource $Font.File
  $Target = Join-Path $UserFonts $Font.File
  if (!(Test-Path -LiteralPath $Source)) {
    throw "Missing packaged font: $Source"
  }
  Copy-Item -LiteralPath $Source -Destination $Target -Force
  New-ItemProperty -Path $RegistryPath -Name $Font.Name -Value $Target -PropertyType String -Force | Out-Null
  Write-Host "Installed $($Font.File)"
}

Write-Host "BaiWuchangKeke fonts installed for current user. Restart PowerPoint if it was already open."
