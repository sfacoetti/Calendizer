# Uso: .\release.ps1 1.0.1   (crea tag + release su GitHub con lo zip; serve gh: https://cli.github.com)
param([Parameter(Mandatory)][string]$Version)
$out = "$PSScriptRoot\dist\release"
Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
dotnet publish "$PSScriptRoot\Host\Calendizer.vbproj" -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -p:Version=$Version -o $out
if ($LASTEXITCODE) { throw "publish fallito" }
$zip = "$PSScriptRoot\dist\Calendizer.zip"
Remove-Item $zip -ErrorAction SilentlyContinue
Compress-Archive "$out\*" $zip
gh release create "v$Version" $zip --title "v$Version" --generate-notes
