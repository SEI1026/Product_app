# PowerShellスクリプトでバージョン更新
param(
    [Parameter(Mandatory=$true)]
    [string]$NewVersion
)

$file = "src\utils\version_checker.py"
$encoding = [System.Text.Encoding]::UTF8

# ファイルを読み込み
$content = Get-Content $file -Encoding UTF8

# バージョン行を更新
$updated = $content -replace 'CURRENT_VERSION = ".*"', "CURRENT_VERSION = `"$NewVersion`""

# ファイルに書き込み
[System.IO.File]::WriteAllLines($file, $updated, $encoding)

# 確認
$verificationLine = $updated | Where-Object { $_ -match 'CURRENT_VERSION = ' }
Write-Host "Updated: $verificationLine"
