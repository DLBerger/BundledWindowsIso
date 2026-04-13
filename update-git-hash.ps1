# update-git-hash.ps1

# Get commit hash (short form)
$commitHash = git rev-parse --short HEAD

# Get list of files changed in last commit (not deleted)
$changedFiles = git diff-tree --no-commit-id --name-only -r HEAD | Where-Object {
    Test-Path $_
}

foreach ($file in $changedFiles) {
    $lines = Get-Content $file
    $lineFound = $false
    $newLines = $lines | ForEach-Object {
        if (-not $lineFound -and $_ -match '^#git-hash\s*=\s*.*$') {
            $lineFound = $true
            'git-hash = "' + $commitHash + '"'
        } else {
            $_
        }
    }
    if ($lineFound) {
        Set-Content -Path $file -Value $newLines
        # Stage and amend ONLY if something changed
        git add $file
        $commitAmended = $true
    }
}

# If any file was amended, amend the last commit
if ($commitAmended) {
    git commit --amend --no-edit
}
