# update-git-hash-local.ps1

<#
To use this with local git commits you need to put the following
in .git/hook/post-commit:

#!/bin/sh
powershell -ExecutionPolicy Bypass -File scripts/update-git-hash-local.ps1

#>

# Get commit hash (short form)
$commitHash = git rev-parse --short HEAD
Write-Host "Current commit hash: $commitHash"

# Get files changed in the latest commit (ignore deleted files)
$changedFiles = git diff-tree --no-commit-id --name-only -r HEAD | Where-Object {
    Test-Path $_
}

$commitAmended = $false

foreach ($file in $changedFiles) {
    Write-Host "Scanning file: $targetFile"
    $lines = Get-Content $file
    $lineFound = $false
    $newLines = $lines | ForEach-Object {
        # Match only lines at the beginning: $GitHash = "<something>"
        if (-not $lineFound -and $_ -match '^\$GitHash\s*=\s*".*"$') {
            $lineFound = $true
            '$GitHash = "' + $commitHash + '"'
        } else {
            $_
        }
    }
    if ($lineFound) {
        Write-Host "Updating $targetFile with the new commit hash"
        Set-Content -Path $file -Value $newLines
        git add $file
        $commitAmended = $true
    }
}

if ($commitAmended) {
    git commit --amend --no-edit
}
